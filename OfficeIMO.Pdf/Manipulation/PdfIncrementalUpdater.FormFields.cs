using System.Globalization;

namespace OfficeIMO.Pdf;

public static partial class PdfIncrementalUpdater {
    private const int IncrementalRadioButtonFlag = 1 << 15;
    private const int IncrementalMultilineFlag = 1 << 12;
    private const int IncrementalPasswordFlag = 1 << 13;
    private const int IncrementalCombFlag = 1 << 24;

    /// <summary>
    /// Appends a simple AcroForm field-value revision to a PDF byte array without rewriting the existing bytes.
    /// </summary>
    public static byte[] UpdateFormFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues, bool keepNeedAppearances = true) {
        return UpdateFormFields(pdf, fieldValues, new PdfIncrementalFormFieldUpdateOptions {
            KeepNeedAppearances = keepNeedAppearances,
            GenerateAppearanceStreams = !keepNeedAppearances
        });
    }

    /// <summary>
    /// Appends a simple AcroForm field-value revision to a PDF byte array without rewriting the existing bytes.
    /// </summary>
    public static byte[] UpdateFormFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues, PdfIncrementalFormFieldUpdateOptions? options) {
        Guard.NotNull(pdf, nameof(pdf));
        ValidateFieldValues(fieldValues);
        PdfIncrementalFormFieldUpdateOptions effectiveOptions = options ?? new PdfIncrementalFormFieldUpdateOptions();

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf);
        ValidateAppendOnlyFormInput(security, fieldValues.Keys);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        if (!security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? rootObject) ||
            rootObject.Value is not PdfDictionary catalog ||
            !catalog.Items.TryGetValue("AcroForm", out PdfObject? acroFormObject) ||
            ResolveDictionary(objects, acroFormObject) is not PdfDictionary acroForm ||
            !acroForm.Items.TryGetValue("Fields", out PdfObject? fieldsObject) ||
            ResolveObject(objects, fieldsObject) is not PdfArray fields) {
            throw new ArgumentException("PDF does not contain a readable AcroForm field tree.", nameof(pdf));
        }

        var remaining = new HashSet<string>(fieldValues.Keys, StringComparer.Ordinal);
        var changedObjectNumbers = new HashSet<int>();
        int inheritedFlags = 0;
        int nextObjectNumber = objects.Keys.Count == 0 ? 1 : objects.Keys.Max() + 1;
        int helveticaFontObjectNumber = 0;
        int? fieldsContainerObjectNumber = fieldsObject is PdfReference fieldsReference ? fieldsReference.ObjectNumber : null;
        for (int i = 0; i < fields.Items.Count; i++) {
            UpdateFormField(
                objects,
                fields.Items[i],
                fieldsContainerObjectNumber,
                null,
                null,
                inheritedFlags,
                null,
                null,
                null,
                fieldValues,
                remaining,
                changedObjectNumbers,
                effectiveOptions,
                new HashSet<int>(),
                ref nextObjectNumber,
                ref helveticaFontObjectNumber);
        }

        if (remaining.Count > 0) {
            throw new ArgumentException("PDF form field was not found: " + string.Join(", ", remaining), nameof(fieldValues));
        }

        acroForm.Items["NeedAppearances"] = new PdfBoolean(effectiveOptions.KeepNeedAppearances);
        if (acroFormObject is PdfReference acroFormReference) {
            changedObjectNumbers.Add(acroFormReference.ObjectNumber);
        } else {
            changedObjectNumbers.Add(security.RootObjectNumber.Value);
        }

        if (changedObjectNumbers.Count == 0) {
            throw new ArgumentException("No supported AcroForm fields were updated.", nameof(fieldValues));
        }

        return AppendIncrementalObjects(pdf, objects, security, trailerRaw, changedObjectNumbers);
    }

    /// <summary>Appends a simple AcroForm field-value revision to a readable PDF stream.</summary>
    public static byte[] UpdateFormFields(Stream input, IReadOnlyDictionary<string, string> fieldValues, bool keepNeedAppearances = true) {
        return UpdateFormFields(input, fieldValues, new PdfIncrementalFormFieldUpdateOptions {
            KeepNeedAppearances = keepNeedAppearances,
            GenerateAppearanceStreams = !keepNeedAppearances
        });
    }

    /// <summary>Appends a simple AcroForm field-value revision to a readable PDF stream.</summary>
    public static byte[] UpdateFormFields(Stream input, IReadOnlyDictionary<string, string> fieldValues, PdfIncrementalFormFieldUpdateOptions? options) {
        Guard.NotNull(input, nameof(input));
        if (!input.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(input));
        }

        using var buffer = new MemoryStream();
        input.CopyTo(buffer);
        return UpdateFormFields(buffer.ToArray(), fieldValues, options);
    }

    /// <summary>Appends a simple AcroForm field-value revision to a PDF file and writes the result to <paramref name="outputPath"/>.</summary>
    public static void UpdateFormFields(string inputPath, string outputPath, IReadOnlyDictionary<string, string> fieldValues, bool keepNeedAppearances = true) {
        UpdateFormFields(inputPath, outputPath, fieldValues, new PdfIncrementalFormFieldUpdateOptions {
            KeepNeedAppearances = keepNeedAppearances,
            GenerateAppearanceStreams = !keepNeedAppearances
        });
    }

    /// <summary>Appends a simple AcroForm field-value revision to a PDF file and writes the result to <paramref name="outputPath"/>.</summary>
    public static void UpdateFormFields(string inputPath, string outputPath, IReadOnlyDictionary<string, string> fieldValues, PdfIncrementalFormFieldUpdateOptions? options) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNullOrWhiteSpace(outputPath, nameof(outputPath));
        File.WriteAllBytes(outputPath, UpdateFormFields(File.ReadAllBytes(inputPath), fieldValues, options));
    }

    private static void ValidateAppendOnlyFormInput(PdfDocumentSecurityInfo security, IEnumerable<string> fieldNames) {
        PdfAppendOnlyMutationReport report = BuildAppendOnlyMutationReport(security);
        if (!report.SupportedActions.Contains("FormFill", StringComparer.Ordinal)) {
            throw new NotSupportedException("Incremental form field updates are not supported for this PDF: " + string.Join(", ", report.Blockers));
        }

        string? lockedFieldName = GetFirstLockedFormFieldName(security, fieldNames);
        if (lockedFieldName is not null) {
            throw new NotSupportedException("Incremental form field updates are not supported for field " + lockedFieldName + " because it is locked by a signature field lock.");
        }
    }

    private static void ValidateFieldValues(IReadOnlyDictionary<string, string> fieldValues) {
        Guard.NotNull(fieldValues, nameof(fieldValues));
        if (fieldValues.Count == 0) {
            throw new ArgumentException("At least one form field value must be provided.", nameof(fieldValues));
        }

        foreach (KeyValuePair<string, string> entry in fieldValues) {
            if (string.IsNullOrWhiteSpace(entry.Key)) {
                throw new ArgumentException("Form field names cannot be empty.", nameof(fieldValues));
            }
        }
    }

    private static void UpdateFormField(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject fieldObject,
        int? containingObjectNumber,
        string? parentName,
        string? inheritedFieldType,
        int inheritedFlags,
        int? inheritedQuadding,
        int? inheritedMaxLength,
        PdfDictionary? inheritedDefaultResources,
        IReadOnlyDictionary<string, string> fieldValues,
        HashSet<string> remaining,
        HashSet<int> changedObjectNumbers,
        PdfIncrementalFormFieldUpdateOptions options,
        HashSet<int> visited,
        ref int nextObjectNumber,
        ref int helveticaFontObjectNumber) {
        int? objectNumber = null;
        if (fieldObject is PdfReference reference) {
            objectNumber = reference.ObjectNumber;
            if (!visited.Add(reference.ObjectNumber)) {
                return;
            }
        }

        if (ResolveObject(objects, fieldObject) is not PdfDictionary field) {
            return;
        }

        string? partialName = TryReadText(objects, field, "T");
        string? fullName = CombineFieldName(parentName, partialName);
        string? fieldType = TryReadName(objects, field, "FT") ?? inheritedFieldType;
        int fieldFlags = ReadFieldFlags(objects, field, inheritedFlags);
        int? fieldQuadding = ReadFieldQuadding(objects, field, inheritedQuadding);
        int? fieldMaxLength = ReadFieldMaxLength(objects, field, inheritedMaxLength);
        PdfDictionary? defaultResources = TryReadDefaultResources(objects, field) ?? inheritedDefaultResources;

        if (fullName is not null && remaining.Contains(fullName) && fieldValues.TryGetValue(fullName, out string? value)) {
            string actualValue = value ?? string.Empty;
            SetIncrementalFieldValue(field, fieldType, actualValue);
            if (objectNumber.HasValue) {
                changedObjectNumbers.Add(objectNumber.Value);
            } else if (containingObjectNumber.HasValue) {
                changedObjectNumbers.Add(containingObjectNumber.Value);
            }

            if (options.GenerateAppearanceStreams) {
                if (string.Equals(fieldType, "Btn", StringComparison.Ordinal)) {
                    bool isRadioButtonGroup = (fieldFlags & IncrementalRadioButtonFlag) != 0;
                    string name = IsOffButtonValue(actualValue) ? "Off" : actualValue;
                    SetIncrementalWidgetAppearanceStates(objects, field, name, isRadioButtonGroup, changedObjectNumbers, new HashSet<int>(), ref nextObjectNumber);
                } else {
                    SetIncrementalTextWidgetAppearances(objects, field, actualValue, fieldFlags, fieldQuadding, fieldMaxLength, defaultResources, changedObjectNumbers, new HashSet<int>(), ref nextObjectNumber, ref helveticaFontObjectNumber);
                }
            }

            remaining.Remove(fullName);
        }

        if (!field.Items.TryGetValue("Kids", out PdfObject? kidsObject) ||
            ResolveObject(objects, kidsObject) is not PdfArray kids) {
            return;
        }

        for (int i = 0; i < kids.Items.Count; i++) {
            int? kidsContainerObjectNumber = kidsObject is PdfReference kidsReference ? kidsReference.ObjectNumber : objectNumber;
            UpdateFormField(objects, kids.Items[i], kidsContainerObjectNumber, fullName, fieldType, fieldFlags, fieldQuadding, fieldMaxLength, defaultResources, fieldValues, remaining, changedObjectNumbers, options, visited, ref nextObjectNumber, ref helveticaFontObjectNumber);
        }
    }

    private static void SetIncrementalFieldValue(PdfDictionary field, string? fieldType, string value) {
        if (string.Equals(fieldType, "Btn", StringComparison.Ordinal)) {
            string name = IsOffButtonValue(value) ? "Off" : value;
            field.Items["V"] = new PdfName(name);
            field.Items["AS"] = new PdfName(name);
            return;
        }

        field.Items["V"] = new PdfStringObj(value, useTextStringEncoding: true);
    }

    private static bool IsOffButtonValue(string value) =>
        string.IsNullOrWhiteSpace(value) ||
        string.Equals(value, "false", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "off", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "0", StringComparison.Ordinal);

    private static void SetIncrementalTextWidgetAppearances(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary field,
        string value,
        int inheritedFlags,
        int? inheritedQuadding,
        int? inheritedMaxLength,
        PdfDictionary? inheritedDefaultResources,
        HashSet<int> changedObjectNumbers,
        HashSet<int> visited,
        ref int nextObjectNumber,
        ref int helveticaFontObjectNumber) {
        int fieldFlags = ReadFieldFlags(objects, field, inheritedFlags);
        int? fieldQuadding = ReadFieldQuadding(objects, field, inheritedQuadding);
        int? fieldMaxLength = ReadFieldMaxLength(objects, field, inheritedMaxLength);
        PdfDictionary? defaultResources = TryReadDefaultResources(objects, field) ?? inheritedDefaultResources;
        if (IsWidget(field) && TryReadRect(field, out double width, out double height)) {
            int appearanceObjectNumber = nextObjectNumber++;
            PdfFormFieldStyle style = ReadIncrementalWidgetAppearanceStyle(objects, field, fieldFlags, fieldQuadding, fieldMaxLength);
            objects[appearanceObjectNumber] = new PdfIndirectObject(
                appearanceObjectNumber,
                0,
                CreateIncrementalTextAppearanceStream(objects, defaultResources, value, width, height, style, ref nextObjectNumber, ref helveticaFontObjectNumber));

            var appearance = new PdfDictionary();
            appearance.Items["N"] = new PdfReference(appearanceObjectNumber, 0);
            field.Items["AP"] = appearance;
            changedObjectNumbers.Add(appearanceObjectNumber);
            if (helveticaFontObjectNumber > 0) {
                changedObjectNumbers.Add(helveticaFontObjectNumber);
            }
        }

        if (!field.Items.TryGetValue("Kids", out PdfObject? kidsObject) ||
            ResolveObject(objects, kidsObject) is not PdfArray kids) {
            return;
        }

        for (int i = 0; i < kids.Items.Count; i++) {
            PdfObject kidObject = kids.Items[i];
            int? kidObjectNumber = kidObject is PdfReference reference ? reference.ObjectNumber : null;
            if (kidObjectNumber.HasValue && !visited.Add(kidObjectNumber.Value)) {
                continue;
            }

            if (ResolveObject(objects, kidObject) is PdfDictionary kid) {
                SetIncrementalTextWidgetAppearances(objects, kid, value, fieldFlags, fieldQuadding, fieldMaxLength, defaultResources, changedObjectNumbers, visited, ref nextObjectNumber, ref helveticaFontObjectNumber);
                if (kidObjectNumber.HasValue) {
                    changedObjectNumbers.Add(kidObjectNumber.Value);
                }
            }
        }
    }

    private static void SetIncrementalWidgetAppearanceStates(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary field,
        string name,
        bool isRadioButtonGroup,
        HashSet<int> changedObjectNumbers,
        HashSet<int> visited,
        ref int nextObjectNumber) {
        if (IsWidget(field)) {
            string appearanceState = isRadioButtonGroup && !HasIncrementalButtonNormalAppearanceState(objects, field, name) ? "Off" : name;
            field.Items["AS"] = new PdfName(appearanceState);
            SetIncrementalButtonWidgetAppearances(objects, field, appearanceState, changedObjectNumbers, ref nextObjectNumber);
        }

        if (!field.Items.TryGetValue("Kids", out PdfObject? kidsObject) ||
            ResolveObject(objects, kidsObject) is not PdfArray kids) {
            return;
        }

        for (int i = 0; i < kids.Items.Count; i++) {
            PdfObject kidObject = kids.Items[i];
            int? kidObjectNumber = kidObject is PdfReference reference ? reference.ObjectNumber : null;
            if (kidObjectNumber.HasValue && !visited.Add(kidObjectNumber.Value)) {
                continue;
            }

            if (ResolveObject(objects, kidObject) is PdfDictionary kid) {
                SetIncrementalWidgetAppearanceStates(objects, kid, name, isRadioButtonGroup, changedObjectNumbers, visited, ref nextObjectNumber);
                if (kidObjectNumber.HasValue) {
                    changedObjectNumbers.Add(kidObjectNumber.Value);
                }
            }
        }
    }

    private static void SetIncrementalButtonWidgetAppearances(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary widget,
        string selectedName,
        HashSet<int> changedObjectNumbers,
        ref int nextObjectNumber) {
        if (!TryReadRect(widget, out double width, out double height)) {
            return;
        }

        var normalAppearances = new PdfDictionary();
        int offAppearanceObjectNumber = nextObjectNumber++;
        objects[offAppearanceObjectNumber] = new PdfIndirectObject(offAppearanceObjectNumber, 0, CreateIncrementalButtonAppearanceStream(width, height, selected: false, ReadIncrementalWidgetAppearanceStyle(objects, widget)));
        normalAppearances.Items["Off"] = new PdfReference(offAppearanceObjectNumber, 0);
        changedObjectNumbers.Add(offAppearanceObjectNumber);

        if (!string.Equals(selectedName, "Off", StringComparison.Ordinal)) {
            int selectedAppearanceObjectNumber = nextObjectNumber++;
            objects[selectedAppearanceObjectNumber] = new PdfIndirectObject(selectedAppearanceObjectNumber, 0, CreateIncrementalButtonAppearanceStream(width, height, selected: true, ReadIncrementalWidgetAppearanceStyle(objects, widget)));
            normalAppearances.Items[selectedName] = new PdfReference(selectedAppearanceObjectNumber, 0);
            changedObjectNumbers.Add(selectedAppearanceObjectNumber);
        }

        var appearance = new PdfDictionary();
        appearance.Items["N"] = normalAppearances;
        widget.Items["AP"] = appearance;
    }

    private static PdfStream CreateIncrementalTextAppearanceStream(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary? inheritedDefaultResources,
        string value,
        double width,
        double height,
        PdfFormFieldStyle style,
        ref int nextObjectNumber,
        ref int helveticaFontObjectNumber) {
        double fontSize = Math.Max(6D, Math.Min(12D, height - 4D));
        int fontObjectNumber = EnsureIncrementalHelveticaFont(objects, inheritedDefaultResources, ref nextObjectNumber, ref helveticaFontObjectNumber);
        string content = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(width, height, value, fontSize, style);
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Form");
        dictionary.Items["BBox"] = CreateNumberArray(0D, 0D, width, height);
        dictionary.Items["Resources"] = CreateIncrementalAppearanceResources(fontObjectNumber);
        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static PdfStream CreateIncrementalButtonAppearanceStream(double width, double height, bool selected, PdfFormFieldStyle? style = null) {
        string content = PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceContent(width, height, selected, style);
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Form");
        dictionary.Items["BBox"] = CreateNumberArray(0D, 0D, width, height);
        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static int EnsureIncrementalHelveticaFont(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary? inheritedDefaultResources,
        ref int nextObjectNumber,
        ref int helveticaFontObjectNumber) {
        if (TryFindFontResource(objects, inheritedDefaultResources, "Helv", out int existingFontObjectNumber)) {
            return existingFontObjectNumber;
        }

        if (helveticaFontObjectNumber == 0) {
            helveticaFontObjectNumber = nextObjectNumber++;
            objects[helveticaFontObjectNumber] = new PdfIndirectObject(helveticaFontObjectNumber, 0, PdfStandardFontDictionaryBuilder.BuildStandardType1FontDictionary(PdfStandardFont.Helvetica));
        }

        return helveticaFontObjectNumber;
    }

    private static PdfDictionary CreateIncrementalAppearanceResources(int helveticaFontObjectNumber) {
        var fonts = new PdfDictionary();
        fonts.Items["Helv"] = new PdfReference(helveticaFontObjectNumber, 0);
        var resources = new PdfDictionary();
        resources.Items["Font"] = fonts;
        return resources;
    }

    private static bool TryFindFontResource(Dictionary<int, PdfIndirectObject> objects, PdfDictionary? resources, string name, out int objectNumber) {
        objectNumber = 0;
        if (resources is null ||
            ResolveDictionary(objects, resources.Items.TryGetValue("Font", out PdfObject? fontsObject) ? fontsObject : null) is not PdfDictionary fonts ||
            !fonts.Items.TryGetValue(name, out PdfObject? fontObject) ||
            fontObject is not PdfReference fontReference) {
            return false;
        }

        objectNumber = fontReference.ObjectNumber;
        return true;
    }

    private static PdfFormFieldStyle ReadIncrementalWidgetAppearanceStyle(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, int fieldFlags = 0, int? inheritedQuadding = null, int? inheritedMaxLength = null) {
        var style = new PdfFormFieldStyle {
            IsMultiline = (fieldFlags & IncrementalMultilineFlag) != 0,
            IsPassword = (fieldFlags & IncrementalPasswordFlag) != 0,
            IsComb = (fieldFlags & IncrementalCombFlag) != 0
        };

        if (TryReadMaxLength(objects, widget, out int maxLength)) {
            style.MaxLength = maxLength;
        } else if (inheritedMaxLength.HasValue) {
            style.MaxLength = inheritedMaxLength.Value;
        }

        if (ResolveDictionary(objects, widget.Items.TryGetValue("MK", out PdfObject? mkObject) ? mkObject : null) is PdfDictionary mk) {
            if (TryReadColor(objects, mk, "BG", out PdfColor backgroundColor)) {
                style.BackgroundColor = backgroundColor;
            }

            if (TryReadColor(objects, mk, "BC", out PdfColor borderColor)) {
                style.BorderColor = borderColor;
            }
        }

        if (TryReadDefaultAppearanceTextColor(objects, widget, out PdfColor textColor)) {
            style.TextColor = textColor;
        }

        if (TryReadWidgetTextAlignment(objects, widget, inheritedQuadding, out PdfFormFieldTextAlignment textAlignment)) {
            style.TextAlignment = textAlignment;
        }

        return style;
    }

    private static bool HasIncrementalButtonNormalAppearanceState(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, string stateName) {
        if (string.IsNullOrEmpty(stateName)) {
            return false;
        }

        return TryGetIncrementalNormalAppearanceObject(objects, widget, out PdfObject? normalAppearance) &&
            normalAppearance is PdfDictionary appearanceStates &&
            appearanceStates.Items.ContainsKey(stateName);
    }

    private static bool TryGetIncrementalNormalAppearanceObject(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, out PdfObject? normalAppearance) {
        normalAppearance = null;
        if (ResolveDictionary(objects, widget.Items.TryGetValue("AP", out PdfObject? appearanceObject) ? appearanceObject : null) is not PdfDictionary appearance ||
            !appearance.Items.TryGetValue("N", out PdfObject? normalAppearanceObject)) {
            return false;
        }

        normalAppearance = ResolveObject(objects, normalAppearanceObject);
        return true;
    }

    private static bool TryReadRect(PdfDictionary dictionary, out double width, out double height) {
        width = 0D;
        height = 0D;
        if (dictionary.Items.TryGetValue("Rect", out PdfObject? rectObject) &&
            rectObject is PdfArray rect &&
            rect.Items.Count >= 4 &&
            rect.Items[0] is PdfNumber x1 &&
            rect.Items[1] is PdfNumber y1 &&
            rect.Items[2] is PdfNumber x2 &&
            rect.Items[3] is PdfNumber y2) {
            width = Math.Abs(x2.Value - x1.Value);
            height = Math.Abs(y2.Value - y1.Value);
            return width > 0D && height > 0D;
        }

        return false;
    }

    private static PdfArray CreateNumberArray(params double[] values) {
        var array = new PdfArray();
        for (int i = 0; i < values.Length; i++) {
            array.Items.Add(new PdfNumber(values[i]));
        }

        return array;
    }

    private static bool IsWidget(PdfDictionary dictionary) {
        return dictionary.Get<PdfName>("Subtype")?.Name == "Widget" ||
            dictionary.Items.ContainsKey("FT");
    }

    private static PdfDictionary? TryReadDefaultResources(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary) {
        return ResolveDictionary(objects, dictionary.Items.TryGetValue("DR", out PdfObject? resourcesObject) ? resourcesObject : null);
    }

    private static int? ReadFieldQuadding(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, int? inheritedQuadding) {
        return TryReadQuadding(objects, field, out int quadding) ? quadding : inheritedQuadding;
    }

    private static int? ReadFieldMaxLength(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, int? inheritedMaxLength) {
        return TryReadMaxLength(objects, field, out int maxLength) ? maxLength : inheritedMaxLength;
    }

    private static bool TryReadMaxLength(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, out int maxLength) {
        maxLength = 0;
        if (!field.Items.TryGetValue("MaxLen", out PdfObject? maxLengthObject) ||
            ResolveObject(objects, maxLengthObject) is not PdfNumber maxLengthNumber ||
            maxLengthNumber.Value < 1 ||
            maxLengthNumber.Value > int.MaxValue ||
            Math.Truncate(maxLengthNumber.Value) != maxLengthNumber.Value) {
            return false;
        }

        maxLength = (int)maxLengthNumber.Value;
        return true;
    }

    private static bool TryReadQuadding(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, out int quadding) {
        quadding = 0;
        if (!field.Items.TryGetValue("Q", out PdfObject? quaddingObject) ||
            ResolveObject(objects, quaddingObject) is not PdfNumber quaddingNumber ||
            Math.Truncate(quaddingNumber.Value) != quaddingNumber.Value) {
            return false;
        }

        quadding = (int)quaddingNumber.Value;
        return quadding >= 0 && quadding <= 2;
    }

    private static bool TryReadWidgetTextAlignment(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, int? inheritedQuadding, out PdfFormFieldTextAlignment textAlignment) {
        textAlignment = PdfFormFieldTextAlignment.Unknown;
        int effectiveQuadding;
        if (TryReadQuadding(objects, widget, out int quadding)) {
            effectiveQuadding = quadding;
        } else if (inheritedQuadding.HasValue) {
            effectiveQuadding = inheritedQuadding.Value;
        } else {
            return false;
        }

        switch (effectiveQuadding) {
            case 0:
                textAlignment = PdfFormFieldTextAlignment.Left;
                return true;
            case 1:
                textAlignment = PdfFormFieldTextAlignment.Center;
                return true;
            case 2:
                textAlignment = PdfFormFieldTextAlignment.Right;
                return true;
            default:
                return false;
        }
    }

    private static bool TryReadColor(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key, out PdfColor color) {
        color = PdfColor.Black;
        if (!dictionary.Items.TryGetValue(key, out PdfObject? colorObject) ||
            ResolveObject(objects, colorObject) is not PdfArray colorArray ||
            colorArray.Items.Count < 3 ||
            ResolveObject(objects, colorArray.Items[0]) is not PdfNumber red ||
            ResolveObject(objects, colorArray.Items[1]) is not PdfNumber green ||
            ResolveObject(objects, colorArray.Items[2]) is not PdfNumber blue) {
            return false;
        }

        color = new PdfColor(ClampColor(red.Value), ClampColor(green.Value), ClampColor(blue.Value));
        return true;
    }

    private static bool TryReadDefaultAppearanceTextColor(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, out PdfColor color) {
        color = PdfColor.Black;
        string? defaultAppearance = TryReadText(objects, widget, "DA");
        if (string.IsNullOrWhiteSpace(defaultAppearance)) {
            return false;
        }

        string[] tokens = defaultAppearance!.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        for (int i = 0; i + 3 < tokens.Length; i++) {
            if (string.Equals(tokens[i + 3], "rg", StringComparison.Ordinal) &&
                double.TryParse(tokens[i], NumberStyles.Float, CultureInfo.InvariantCulture, out double red) &&
                double.TryParse(tokens[i + 1], NumberStyles.Float, CultureInfo.InvariantCulture, out double green) &&
                double.TryParse(tokens[i + 2], NumberStyles.Float, CultureInfo.InvariantCulture, out double blue)) {
                color = new PdfColor(ClampColor(red), ClampColor(green), ClampColor(blue));
                return true;
            }
        }

        return false;
    }

    private static double ClampColor(double value) {
        if (value < 0D) {
            return 0D;
        }

        return value > 1D ? 1D : value;
    }

    private static byte[] AppendIncrementalObjects(
        byte[] pdf,
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        string trailerRaw,
        HashSet<int> changedObjectNumbers) {
        if (!security.RootObjectNumber.HasValue) {
            throw new InvalidOperationException("PDF root catalog reference is required for an incremental update.");
        }

        if (!security.LastStartXrefOffset.HasValue) {
            throw new InvalidOperationException("PDF startxref offset is required for an incremental update.");
        }

        var identityMap = objects.Keys.ToDictionary(static objectNumber => objectNumber, static objectNumber => objectNumber);
        var context = new PdfPageExtractor.SerializationContext(identityMap, pagesObjectId: 0, new Dictionary<int, Dictionary<string, PdfObject>>(), objects, preserveReferenceGenerations: true);
        int[] objectNumbers = changedObjectNumbers.OrderBy(static objectNumber => objectNumber).ToArray();
        var serialized = new List<(int ObjectNumber, int Generation, byte[] Bytes)>(objectNumbers.Length);
        foreach (int objectNumber in objectNumbers) {
            if (!objects.TryGetValue(objectNumber, out PdfIndirectObject? indirect)) {
                throw new InvalidOperationException("PDF object " + objectNumber.ToString(CultureInfo.InvariantCulture) + " was changed but could not be found.");
            }

            serialized.Add((objectNumber, indirect.Generation, PdfObjectBytes.WrapIndirectObject(objectNumber, indirect.Generation, PdfPageExtractor.SerializeObject(indirect.Value, context))));
        }

        using var output = new MemoryStream(pdf.Length + serialized.Sum(static item => item.Bytes.Length) + (serialized.Count * 32) + 256);
        output.Write(pdf, 0, pdf.Length);
        if (pdf.Length == 0 || (pdf[pdf.Length - 1] != (byte)'\n' && pdf[pdf.Length - 1] != (byte)'\r')) {
            output.WriteByte((byte)'\n');
        }

        var offsets = new Dictionary<int, long>();
        foreach (var item in serialized) {
            offsets[item.ObjectNumber] = output.Position;
            output.Write(item.Bytes, 0, item.Bytes.Length);
        }

        long xrefOffset = output.Position;
        int size = Math.Max(objects.Keys.Max(), objectNumbers.Max()) + 1;

        using var writer = new StreamWriter(output, Encoding.ASCII, 1024, leaveOpen: true) { NewLine = "\n" };
        writer.WriteLine("xref");
        foreach (int objectNumber in objectNumbers) {
            int generation = serialized.First(item => item.ObjectNumber == objectNumber).Generation;
            writer.WriteLine(objectNumber.ToString(CultureInfo.InvariantCulture) + " 1");
            writer.WriteLine(offsets[objectNumber].ToString("0000000000", CultureInfo.InvariantCulture) + " " + generation.ToString("00000", CultureInfo.InvariantCulture) + " n ");
        }

        writer.WriteLine("trailer");
        writer.WriteLine("<< /Size " + size.ToString(CultureInfo.InvariantCulture) +
            " /Root " + BuildExistingTrailerReference(objects, security.RootObjectNumber.Value) +
            (security.InfoObjectNumber.HasValue ? " /Info " + BuildExistingTrailerReference(objects, security.InfoObjectNumber.Value) : string.Empty) +
            " /Prev " + security.LastStartXrefOffset.Value.ToString(CultureInfo.InvariantCulture) +
            ReadTrailerIdEntry(trailerRaw) +
            " >>");
        writer.WriteLine("startxref");
        writer.WriteLine(xrefOffset.ToString(CultureInfo.InvariantCulture));
        writer.WriteLine("%%EOF");
        writer.Flush();

        return output.ToArray();
    }

    private static PdfObject? ResolveObject(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) =>
        PdfObjectLookup.Resolve(objects, value);

    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> objects, PdfObject? value) =>
        ResolveObject(objects, value) as PdfDictionary;

    private static string? TryReadName(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) =>
        dictionary.Items.TryGetValue(key, out PdfObject? value) &&
        ResolveObject(objects, value) is PdfName name &&
        !string.IsNullOrEmpty(name.Name)
            ? name.Name
            : null;

    private static string? TryReadText(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, string key) =>
        dictionary.Items.TryGetValue(key, out PdfObject? value) &&
        ResolveObject(objects, value) is PdfStringObj text
            ? text.Value
            : null;

    private static int ReadFieldFlags(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, int inheritedFlags) {
        if (!field.Items.TryGetValue("Ff", out PdfObject? flagsObject) ||
            ResolveObject(objects, flagsObject) is not PdfNumber flagsNumber) {
            return inheritedFlags;
        }

        return (int)flagsNumber.Value;
    }

    private static string? CombineFieldName(string? parentName, string? partialName) {
        if (string.IsNullOrEmpty(partialName)) {
            return parentName;
        }

        return string.IsNullOrEmpty(parentName) ? partialName : parentName + "." + partialName;
    }
}
