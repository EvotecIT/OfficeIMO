using OfficeIMO.Drawing.Internal;
using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfIncrementalUpdater {
    private const string IncrementalDefaultAppearanceFontName = "Helv";
    private const int IncrementalRadioButtonFlag = 1 << 15;
    private const int IncrementalMultilineFlag = 1 << 12;
    private const int IncrementalPasswordFlag = 1 << 13;
    private const int IncrementalEditableChoiceFlag = 1 << 18;
    private const int IncrementalMultiSelectChoiceFlag = 1 << 21;
    private const int IncrementalCombFlag = 1 << 24;

    private readonly struct IncrementalChoiceFillValue {
        public IncrementalChoiceFillValue(string exportValue, string displayValue) {
            ExportValue = exportValue;
            DisplayValue = displayValue;
        }

        public string ExportValue { get; }
        public string DisplayValue { get; }
    }

    private readonly struct IncrementalPreparedFieldValue {
        private IncrementalPreparedFieldValue(string[] storedValues, string appearanceValue, bool forceMultilineAppearance) {
            StoredValues = storedValues;
            AppearanceValue = appearanceValue;
            ForceMultilineAppearance = forceMultilineAppearance;
        }

        public string[] StoredValues { get; }
        public string FirstStoredValue => StoredValues[0];
        public string AppearanceValue { get; }
        public bool IsMultiple => StoredValues.Length > 1;
        public bool ForceMultilineAppearance { get; }

        public static IncrementalPreparedFieldValue Scalar(string storedValue, string appearanceValue) =>
            new IncrementalPreparedFieldValue(new[] { storedValue }, appearanceValue, forceMultilineAppearance: false);

        public static IncrementalPreparedFieldValue Multiple(string[] storedValues, string appearanceValue) =>
            new IncrementalPreparedFieldValue(storedValues, appearanceValue, forceMultilineAppearance: true);
    }

    /// <summary>
    /// Appends a simple AcroForm field-value revision to a PDF byte array without rewriting the existing bytes.
    /// </summary>
    public static byte[] UpdateFormFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues, bool keepNeedAppearances = true) {
        return UpdateFormFields(pdf, ToIncrementalFormFieldValues(fieldValues), new PdfIncrementalFormFieldUpdateOptions {
            KeepNeedAppearances = keepNeedAppearances,
            GenerateAppearanceStreams = !keepNeedAppearances
        });
    }

    /// <summary>
    /// Appends a simple AcroForm field-value revision to a PDF byte array without rewriting the existing bytes.
    /// </summary>
    public static byte[] UpdateFormFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues, PdfIncrementalFormFieldUpdateOptions? options) {
        return UpdateFormFields(pdf, ToIncrementalFormFieldValues(fieldValues), options);
    }

    /// <summary>Appends a form-field revision using optional password and parsing settings.</summary>
    public static byte[] UpdateFormFields(byte[] pdf, IReadOnlyDictionary<string, string> fieldValues, PdfIncrementalFormFieldUpdateOptions? options, PdfReadOptions? readOptions) {
        return UpdateFormFields(pdf, ToIncrementalFormFieldValues(fieldValues), options, readOptions);
    }

    /// <summary>
    /// Appends a simple AcroForm field-value revision to a PDF byte array without rewriting the existing bytes.
    /// </summary>
    public static byte[] UpdateFormFields(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, bool keepNeedAppearances = true) {
        return UpdateFormFields(pdf, fieldValues, new PdfIncrementalFormFieldUpdateOptions {
            KeepNeedAppearances = keepNeedAppearances,
            GenerateAppearanceStreams = !keepNeedAppearances
        });
    }

    /// <summary>
    /// Appends a simple AcroForm field-value revision to a PDF byte array without rewriting the existing bytes.
    /// </summary>
    public static byte[] UpdateFormFields(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfIncrementalFormFieldUpdateOptions? options) {
        return UpdateFormFields(pdf, fieldValues, options, readOptions: null);
    }

    /// <summary>Appends a form-field revision using optional password and parsing settings.</summary>
    public static byte[] UpdateFormFields(byte[] pdf, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfIncrementalFormFieldUpdateOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        ValidateFieldValues(fieldValues);
        PdfIncrementalFormFieldUpdateOptions effectiveOptions = options ?? new PdfIncrementalFormFieldUpdateOptions();
        _ = PdfMutationPlanner.RequireAppendOnly(
            pdf,
            PdfMutationOperation.FillFormFields,
            readOptions,
            fieldNames: fieldValues.Keys);

        PdfDocumentSecurityInfo security = PdfSyntax.ReadDocumentSecurityInfo(pdf, readOptions);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf, readOptions);
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
        PdfDictionary? acroFormDefaultResources = TryReadDefaultResources(objects, acroForm);
        string? acroFormDefaultAppearance = TryReadText(objects, acroForm, "DA");
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
                acroFormDefaultResources,
                acroFormDefaultAppearance,
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
        if (acroFormObject is PdfReference acroFormReferenceForNeedAppearances) {
            changedObjectNumbers.Add(acroFormReferenceForNeedAppearances.ObjectNumber);
        } else {
            changedObjectNumbers.Add(security.RootObjectNumber.Value);
        }

        if (changedObjectNumbers.Count == 0) {
            throw new ArgumentException("No supported AcroForm fields were updated.", nameof(fieldValues));
        }

        PdfStandardSecurityHandler? encryptionHandler = null;
        if (security.HasEncryption &&
            !PdfSyntax.TryCreateDecryptor(objects, trailerRaw, readOptions, out encryptionHandler)) {
            throw new PdfUnsupportedEncryptionException("PDF encryption context could not be created for the incremental form update.");
        }

        return AppendIncrementalObjects(pdf, objects, security, trailerRaw, changedObjectNumbers, encryptionHandler);
    }

    /// <summary>Appends a simple AcroForm field-value revision to a readable PDF stream.</summary>
    public static byte[] UpdateFormFields(Stream input, IReadOnlyDictionary<string, string> fieldValues, bool keepNeedAppearances = true) {
        return UpdateFormFields(input, ToIncrementalFormFieldValues(fieldValues), new PdfIncrementalFormFieldUpdateOptions {
            KeepNeedAppearances = keepNeedAppearances,
            GenerateAppearanceStreams = !keepNeedAppearances
        });
    }

    /// <summary>Appends a simple AcroForm field-value revision to a readable PDF stream.</summary>
    public static byte[] UpdateFormFields(Stream input, IReadOnlyDictionary<string, string> fieldValues, PdfIncrementalFormFieldUpdateOptions? options) {
        return UpdateFormFields(input, ToIncrementalFormFieldValues(fieldValues), options);
    }

    /// <summary>Appends a form-field revision from a readable stream using optional password and parsing settings.</summary>
    public static byte[] UpdateFormFields(Stream input, IReadOnlyDictionary<string, string> fieldValues, PdfIncrementalFormFieldUpdateOptions? options, PdfReadOptions? readOptions) {
        return UpdateFormFields(input, ToIncrementalFormFieldValues(fieldValues), options, readOptions);
    }

    /// <summary>Appends a simple AcroForm field-value revision to a readable PDF stream.</summary>
    public static byte[] UpdateFormFields(Stream input, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, bool keepNeedAppearances = true) {
        return UpdateFormFields(input, fieldValues, new PdfIncrementalFormFieldUpdateOptions {
            KeepNeedAppearances = keepNeedAppearances,
            GenerateAppearanceStreams = !keepNeedAppearances
        });
    }

    /// <summary>Appends a simple AcroForm field-value revision to a readable PDF stream.</summary>
    public static byte[] UpdateFormFields(Stream input, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfIncrementalFormFieldUpdateOptions? options) {
        return UpdateFormFields(input, fieldValues, options, readOptions: null);
    }

    /// <summary>Appends a form-field revision from a readable stream using optional password and parsing settings.</summary>
    public static byte[] UpdateFormFields(Stream input, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfIncrementalFormFieldUpdateOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNull(input, nameof(input));
        if (!input.CanRead) {
            throw new ArgumentException("Stream must be readable.", nameof(input));
        }

        using var buffer = new MemoryStream();
        input.CopyTo(buffer);
        return UpdateFormFields(buffer.ToArray(), fieldValues, options, readOptions);
    }

    /// <summary>Appends a simple AcroForm field-value revision to a PDF file and writes the result to <paramref name="outputPath"/>.</summary>
    public static void UpdateFormFields(string inputPath, string outputPath, IReadOnlyDictionary<string, string> fieldValues, bool keepNeedAppearances = true) {
        UpdateFormFields(inputPath, outputPath, ToIncrementalFormFieldValues(fieldValues), new PdfIncrementalFormFieldUpdateOptions {
            KeepNeedAppearances = keepNeedAppearances,
            GenerateAppearanceStreams = !keepNeedAppearances
        });
    }

    /// <summary>Appends a simple AcroForm field-value revision to a PDF file and writes the result to <paramref name="outputPath"/>.</summary>
    public static void UpdateFormFields(string inputPath, string outputPath, IReadOnlyDictionary<string, string> fieldValues, PdfIncrementalFormFieldUpdateOptions? options) {
        UpdateFormFields(inputPath, outputPath, ToIncrementalFormFieldValues(fieldValues), options);
    }

    /// <summary>Appends a form-field revision to a PDF file using optional password and parsing settings.</summary>
    public static void UpdateFormFields(string inputPath, string outputPath, IReadOnlyDictionary<string, string> fieldValues, PdfIncrementalFormFieldUpdateOptions? options, PdfReadOptions? readOptions) {
        UpdateFormFields(inputPath, outputPath, ToIncrementalFormFieldValues(fieldValues), options, readOptions);
    }

    /// <summary>Appends a simple AcroForm field-value revision to a PDF file and writes the result to <paramref name="outputPath"/>.</summary>
    public static void UpdateFormFields(string inputPath, string outputPath, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, bool keepNeedAppearances = true) {
        UpdateFormFields(inputPath, outputPath, fieldValues, new PdfIncrementalFormFieldUpdateOptions {
            KeepNeedAppearances = keepNeedAppearances,
            GenerateAppearanceStreams = !keepNeedAppearances
        });
    }

    /// <summary>Appends a simple AcroForm field-value revision to a PDF file and writes the result to <paramref name="outputPath"/>.</summary>
    public static void UpdateFormFields(string inputPath, string outputPath, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfIncrementalFormFieldUpdateOptions? options) {
        UpdateFormFields(inputPath, outputPath, fieldValues, options, readOptions: null);
    }

    /// <summary>Appends a form-field revision to a PDF file using optional password and parsing settings.</summary>
    public static void UpdateFormFields(string inputPath, string outputPath, IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues, PdfIncrementalFormFieldUpdateOptions? options, PdfReadOptions? readOptions) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNullOrWhiteSpace(outputPath, nameof(outputPath));
        OfficeFileCommit.WriteAllBytes(outputPath, UpdateFormFields(File.ReadAllBytes(inputPath), fieldValues, options, readOptions));
    }

    private static Dictionary<string, PdfFormFieldValue> ToIncrementalFormFieldValues(IReadOnlyDictionary<string, string> fieldValues) {
        Guard.NotNull(fieldValues, nameof(fieldValues));

        var converted = new Dictionary<string, PdfFormFieldValue>(fieldValues.Count, StringComparer.Ordinal);
        foreach (KeyValuePair<string, string> entry in fieldValues) {
            converted[entry.Key] = PdfFormFieldValue.From(entry.Value ?? string.Empty);
        }

        return converted;
    }

    private static void ValidateFieldValues(IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues) {
        Guard.NotNull(fieldValues, nameof(fieldValues));
        if (fieldValues.Count == 0) {
            throw new ArgumentException("At least one form field value must be provided.", nameof(fieldValues));
        }

        foreach (KeyValuePair<string, PdfFormFieldValue> entry in fieldValues) {
            if (string.IsNullOrWhiteSpace(entry.Key)) {
                throw new ArgumentException("Form field names cannot be empty.", nameof(fieldValues));
            }

            if (entry.Value is null) {
                throw new ArgumentException("Field values cannot be null. Use an empty string to clear a text value.", nameof(fieldValues));
            }

            if (entry.Value.Values.Count == 0) {
                throw new ArgumentException("Field values must contain at least one entry. Use an empty string to clear a text value.", nameof(fieldValues));
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
        string? inheritedDefaultAppearance,
        IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues,
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
        string? defaultAppearance = TryReadText(objects, field, "DA") ?? inheritedDefaultAppearance;

        if (fullName is not null && remaining.Contains(fullName) && fieldValues.TryGetValue(fullName, out PdfFormFieldValue? value)) {
            IncrementalPreparedFieldValue preparedValue = PrepareIncrementalFieldValue(objects, field, fieldType, fieldFlags, value);
            SetIncrementalFieldValue(objects, field, fieldType, fieldFlags, preparedValue);
            if (objectNumber.HasValue) {
                changedObjectNumbers.Add(objectNumber.Value);
            } else if (containingObjectNumber.HasValue) {
                changedObjectNumbers.Add(containingObjectNumber.Value);
            }

            if (string.Equals(fieldType, "Btn", StringComparison.Ordinal)) {
                bool isRadioButtonGroup = (fieldFlags & IncrementalRadioButtonFlag) != 0;
                string name = IsOffButtonValue(preparedValue.FirstStoredValue) ? "Off" : preparedValue.FirstStoredValue;
                SetIncrementalWidgetAppearanceStates(objects, field, name, isRadioButtonGroup, options.GenerateAppearanceStreams, changedObjectNumbers, new HashSet<int>(), ref nextObjectNumber);
            } else if (options.GenerateAppearanceStreams) {
                SetIncrementalTextWidgetAppearances(objects, field, preparedValue.AppearanceValue, fieldFlags, fieldQuadding, fieldMaxLength, defaultResources, defaultAppearance, preparedValue.ForceMultilineAppearance, changedObjectNumbers, new HashSet<int>(), ref nextObjectNumber, ref helveticaFontObjectNumber);
            }

            remaining.Remove(fullName);
        }

        if (!field.Items.TryGetValue("Kids", out PdfObject? kidsObject) ||
            ResolveObject(objects, kidsObject) is not PdfArray kids) {
            return;
        }

        int? kidsContainerObjectNumber = kidsObject is PdfReference kidsReference
            ? kidsReference.ObjectNumber
            : objectNumber ?? containingObjectNumber;
        for (int i = 0; i < kids.Items.Count; i++) {
            UpdateFormField(objects, kids.Items[i], kidsContainerObjectNumber, fullName, fieldType, fieldFlags, fieldQuadding, fieldMaxLength, defaultResources, defaultAppearance, fieldValues, remaining, changedObjectNumbers, options, visited, ref nextObjectNumber, ref helveticaFontObjectNumber);
        }
    }

    private static void SetIncrementalFieldValue(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, string? fieldType, int fieldFlags, IncrementalPreparedFieldValue value) {
        if (string.Equals(fieldType, "Btn", StringComparison.Ordinal)) {
            string name = IsOffButtonValue(value.FirstStoredValue) ? "Off" : value.FirstStoredValue;
            bool isRadioButtonGroup = (fieldFlags & IncrementalRadioButtonFlag) != 0;
            if (isRadioButtonGroup && !string.Equals(name, "Off", StringComparison.Ordinal)) {
                HashSet<string> availableStates = CollectIncrementalButtonNormalAppearanceStates(objects, field, new HashSet<int>());
                if (!availableStates.Contains(name)) {
                    throw new ArgumentException($"PDF radio button field cannot be filled with value '{name}' because it is not one of the available appearance states.", nameof(value));
                }
            }

            field.Items["V"] = new PdfName(name);
            field.Items["AS"] = new PdfName(name);
            return;
        }

        if (string.Equals(fieldType, "Ch", StringComparison.Ordinal) && value.IsMultiple) {
            field.Items["V"] = CreateIncrementalStringArray(value.StoredValues);
            return;
        }

        field.Items["V"] = new PdfStringObj(value.FirstStoredValue, useTextStringEncoding: true);
    }

    private static IncrementalPreparedFieldValue PrepareIncrementalFieldValue(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, string? fieldType, int fieldFlags, PdfFormFieldValue value) {
        IReadOnlyList<string> values = value.Values;
        string firstValue = values[0];
        if (string.Equals(fieldType, "Btn", StringComparison.Ordinal)) {
            if (values.Count > 1) {
                throw new ArgumentException("PDF button field cannot be filled with multiple values.", nameof(value));
            }

            string buttonValue = PrepareIncrementalButtonFieldValue(objects, field, fieldFlags, firstValue);
            return IncrementalPreparedFieldValue.Scalar(buttonValue, buttonValue);
        }

        if (!string.Equals(fieldType, "Ch", StringComparison.Ordinal)) {
            if (values.Count > 1) {
                throw new ArgumentException("PDF text field cannot be filled with multiple values.", nameof(value));
            }

            return IncrementalPreparedFieldValue.Scalar(firstValue, firstValue);
        }

        bool isMultiSelectChoice = (fieldFlags & IncrementalMultiSelectChoiceFlag) != 0;
        if (values.Count > 1 && !isMultiSelectChoice) {
            throw new ArgumentException("PDF scalar choice field cannot be filled with multiple values.", nameof(value));
        }

        IReadOnlyList<IncrementalChoiceFillValue> choiceValues = ResolveIncrementalChoiceFillValues(objects, field, (fieldFlags & IncrementalEditableChoiceFlag) != 0, values);
        if (isMultiSelectChoice) {
            return IncrementalPreparedFieldValue.Multiple(
                choiceValues.Select(item => item.ExportValue).ToArray(),
                string.Join("\n", choiceValues.Select(item => item.DisplayValue)));
        }

        IncrementalChoiceFillValue choiceValue = choiceValues[0];
        return IncrementalPreparedFieldValue.Scalar(choiceValue.ExportValue, choiceValue.DisplayValue);
    }

    private static IReadOnlyList<IncrementalChoiceFillValue> ResolveIncrementalChoiceFillValues(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, bool isEditableChoice, IReadOnlyList<string> values) {
        if (!field.Items.TryGetValue("Opt", out PdfObject? optionsObject) ||
            ResolveObject(objects, optionsObject) is not PdfArray options ||
            options.Items.Count == 0) {
            return values.Select(static item => new IncrementalChoiceFillValue(item, item)).ToArray();
        }

        var resolved = new List<IncrementalChoiceFillValue>(values.Count);
        for (int valueIndex = 0; valueIndex < values.Count; valueIndex++) {
            resolved.Add(ResolveIncrementalChoiceFillValue(objects, options, isEditableChoice, values[valueIndex]));
        }

        return resolved;
    }

    private static IncrementalChoiceFillValue ResolveIncrementalChoiceFillValue(Dictionary<int, PdfIndirectObject> objects, PdfArray options, bool isEditableChoice, string value) {
        for (int i = 0; i < options.Items.Count; i++) {
            PdfObject? optionObject = ResolveObject(objects, options.Items[i]);
            if (optionObject is PdfArray pair &&
                pair.Items.Count >= 2 &&
                TryReadOptionText(objects, pair.Items[0], out string? exportValue) &&
                exportValue is not null &&
                TryReadOptionText(objects, pair.Items[1], out string? displayValue) &&
                displayValue is not null) {
                if (string.Equals(value, exportValue, StringComparison.Ordinal) ||
                    string.Equals(value, displayValue, StringComparison.Ordinal)) {
                    return new IncrementalChoiceFillValue(exportValue, displayValue);
                }

                continue;
            }

            if (optionObject is not null &&
                TryReadOptionText(objects, optionObject, out string? optionValue) &&
                optionValue is not null &&
                string.Equals(value, optionValue, StringComparison.Ordinal)) {
                return new IncrementalChoiceFillValue(optionValue, optionValue);
            }
        }

        if (isEditableChoice) {
            return new IncrementalChoiceFillValue(value, value);
        }

        throw new ArgumentException($"PDF choice field cannot be filled with value '{value}' because it is not one of the allowed options.", nameof(value));
    }

    private static string PrepareIncrementalButtonFieldValue(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, int fieldFlags, string value) {
        if (IsOffButtonValue(value)) {
            return "Off";
        }

        HashSet<string> availableStates = CollectIncrementalButtonNormalAppearanceStates(objects, field, new HashSet<int>());
        bool isRadioButtonGroup = (fieldFlags & IncrementalRadioButtonFlag) != 0;
        if (availableStates.Contains(value)) {
            return value;
        }

        if (!isRadioButtonGroup &&
            IsTruthyButtonValue(value) &&
            availableStates.Count == 1) {
            return availableStates.Single();
        }

        string fieldKind = isRadioButtonGroup ? "radio button" : "checkbox";
        throw new ArgumentException($"PDF {fieldKind} field cannot be filled with value '{value}' because it is not one of the available appearance states.", nameof(value));
    }

    private static bool TryReadOptionText(Dictionary<int, PdfIndirectObject> objects, PdfObject value, out string? text) {
        text = null;
        switch (ResolveObject(objects, value)) {
            case PdfStringObj stringObj:
                text = stringObj.Value;
                return true;
            case PdfName name:
                text = name.Name;
                return true;
            default:
                return false;
        }
    }

    private static bool IsOffButtonValue(string value) =>
        string.IsNullOrWhiteSpace(value) ||
        string.Equals(value, "false", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "off", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "0", StringComparison.Ordinal);

    private static bool IsTruthyButtonValue(string value) =>
        string.Equals(value, "true", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "yes", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "on", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(value, "1", StringComparison.Ordinal);

    private static void SetIncrementalTextWidgetAppearances(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary field,
        string value,
        int inheritedFlags,
        int? inheritedQuadding,
        int? inheritedMaxLength,
        PdfDictionary? inheritedDefaultResources,
        string? inheritedDefaultAppearance,
        bool forceMultilineAppearance,
        HashSet<int> changedObjectNumbers,
        HashSet<int> visited,
        ref int nextObjectNumber,
        ref int helveticaFontObjectNumber) {
        int fieldFlags = ReadFieldFlags(objects, field, inheritedFlags);
        int? fieldQuadding = ReadFieldQuadding(objects, field, inheritedQuadding);
        int? fieldMaxLength = ReadFieldMaxLength(objects, field, inheritedMaxLength);
        PdfDictionary? defaultResources = TryReadDefaultResources(objects, field) ?? inheritedDefaultResources;
        string? defaultAppearance = TryReadText(objects, field, "DA") ?? inheritedDefaultAppearance;
        if (IsWidget(field) && TryReadRect(field, out double width, out double height)) {
            PdfDictionary? widgetAppearanceResources = TryReadIncrementalNormalAppearanceResources(objects, field);
            PdfDictionary? widgetPageResources = TryReadIncrementalWidgetPageResources(objects, field);
            int appearanceObjectNumber = nextObjectNumber++;
            PdfFormFieldStyle style = ReadIncrementalWidgetAppearanceStyle(objects, field, fieldFlags, fieldQuadding, fieldMaxLength, defaultAppearance);
            if (forceMultilineAppearance) {
                style.IsMultiline = true;
            }

            double fontSize = ReadIncrementalWidgetAppearanceFontSize(defaultAppearance, height);
            objects[appearanceObjectNumber] = new PdfIndirectObject(
                appearanceObjectNumber,
                0,
                CreateIncrementalTextAppearanceStream(objects, defaultResources, widgetAppearanceResources, widgetPageResources, defaultAppearance, value, width, height, style, fontSize, ref nextObjectNumber, ref helveticaFontObjectNumber));

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
                SetIncrementalTextWidgetAppearances(objects, kid, value, fieldFlags, fieldQuadding, fieldMaxLength, defaultResources, defaultAppearance, forceMultilineAppearance, changedObjectNumbers, visited, ref nextObjectNumber, ref helveticaFontObjectNumber);
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
        bool generateAppearanceStreams,
        HashSet<int> changedObjectNumbers,
        HashSet<int> visited,
        ref int nextObjectNumber) {
        if (IsWidget(field)) {
            string? widgetOnState = ReadIncrementalWidgetOnAppearanceState(objects, field);
            string appearanceState = isRadioButtonGroup && !string.Equals(widgetOnState, name, StringComparison.Ordinal) ? "Off" : name;
            field.Items["AS"] = new PdfName(appearanceState);
            if (generateAppearanceStreams) {
                string? preservedOnState = widgetOnState ?? (!string.Equals(name, "Off", StringComparison.Ordinal) ? name : null);
                SetIncrementalButtonWidgetAppearances(objects, field, appearanceState, preservedOnState, isRadioButtonGroup, changedObjectNumbers, ref nextObjectNumber);
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
                SetIncrementalWidgetAppearanceStates(objects, kid, name, isRadioButtonGroup, generateAppearanceStreams, changedObjectNumbers, visited, ref nextObjectNumber);
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
        string? preservedOnState,
        bool isRadioButton,
        HashSet<int> changedObjectNumbers,
        ref int nextObjectNumber) {
        if (!TryReadRect(widget, out double width, out double height)) {
            return;
        }

        var normalAppearances = new PdfDictionary();
        int offAppearanceObjectNumber = nextObjectNumber++;
        objects[offAppearanceObjectNumber] = new PdfIndirectObject(offAppearanceObjectNumber, 0, CreateIncrementalButtonAppearanceStream(width, height, selected: false, isRadioButton, ReadIncrementalWidgetAppearanceStyle(objects, widget)));
        normalAppearances.Items["Off"] = new PdfReference(offAppearanceObjectNumber, 0);
        changedObjectNumbers.Add(offAppearanceObjectNumber);

        string? onState = string.Equals(selectedName, "Off", StringComparison.Ordinal) ? preservedOnState : selectedName;
        if (!string.IsNullOrEmpty(onState)) {
            int onAppearanceObjectNumber = nextObjectNumber++;
            objects[onAppearanceObjectNumber] = new PdfIndirectObject(onAppearanceObjectNumber, 0, CreateIncrementalButtonAppearanceStream(width, height, selected: true, isRadioButton, ReadIncrementalWidgetAppearanceStyle(objects, widget)));
            normalAppearances.Items[onState!] = new PdfReference(onAppearanceObjectNumber, 0);
            changedObjectNumbers.Add(onAppearanceObjectNumber);
        }

        var appearance = new PdfDictionary();
        appearance.Items["N"] = normalAppearances;
        widget.Items["AP"] = appearance;
    }

    private static PdfStream CreateIncrementalTextAppearanceStream(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary? inheritedDefaultResources,
        PdfDictionary? widgetAppearanceResources,
        PdfDictionary? widgetPageResources,
        string? defaultAppearance,
        string value,
        double width,
        double height,
        PdfFormFieldStyle style,
        double fontSize,
        ref int nextObjectNumber,
        ref int helveticaFontObjectNumber) {
        string fontResourceName = IncrementalDefaultAppearanceFontName;
        PdfDictionary resources;
        if (!TryCreateIncrementalDefaultAppearanceFontResources(objects, defaultAppearance, out fontResourceName, out resources, inheritedDefaultResources, widgetAppearanceResources, widgetPageResources)) {
            PdfReference fontReference = EnsureIncrementalHelveticaFont(objects, ref nextObjectNumber, ref helveticaFontObjectNumber, inheritedDefaultResources, widgetAppearanceResources, widgetPageResources);
            resources = CreateIncrementalAppearanceResources(IncrementalDefaultAppearanceFontName, fontReference);
        }

        string content = PdfAcroFormDictionaryBuilder.BuildTextFieldAppearanceContent(width, height, value, fontSize, style, fontResourceName: fontResourceName);
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Form");
        dictionary.Items["BBox"] = CreateNumberArray(0D, 0D, width, height);
        dictionary.Items["Resources"] = resources;
        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static PdfStream CreateIncrementalButtonAppearanceStream(double width, double height, bool selected, bool isRadioButton, PdfFormFieldStyle? style = null) {
        string content = isRadioButton
            ? PdfAcroFormDictionaryBuilder.BuildRadioButtonAppearanceContent(width, height, selected, style)
            : PdfAcroFormDictionaryBuilder.BuildCheckBoxAppearanceContent(width, height, selected, style);
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Form");
        dictionary.Items["BBox"] = CreateNumberArray(0D, 0D, width, height);
        return new PdfStream(dictionary, PdfEncoding.Latin1GetBytes(content));
    }

    private static PdfReference EnsureIncrementalHelveticaFont(
        Dictionary<int, PdfIndirectObject> objects,
        ref int nextObjectNumber,
        ref int helveticaFontObjectNumber,
        params PdfDictionary?[] candidateResources) {
        if (TryFindFontResource(objects, IncrementalDefaultAppearanceFontName, out PdfReference existingFontReference, candidateResources)) {
            return existingFontReference;
        }

        if (helveticaFontObjectNumber == 0) {
            helveticaFontObjectNumber = nextObjectNumber++;
            objects[helveticaFontObjectNumber] = new PdfIndirectObject(helveticaFontObjectNumber, 0, PdfStandardFontDictionaryBuilder.BuildStandardType1FontDictionary(PdfStandardFont.Helvetica));
        }

        return new PdfReference(helveticaFontObjectNumber, 0);
    }

    private static PdfDictionary CreateIncrementalAppearanceResources(string fontResourceName, PdfObject fontResource) {
        var fonts = new PdfDictionary();
        fonts.Items[fontResourceName] = fontResource;
        var resources = new PdfDictionary();
        resources.Items["Font"] = fonts;
        return resources;
    }

    private static PdfArray CreateIncrementalStringArray(IEnumerable<string> values) {
        var array = new PdfArray();
        foreach (string value in values) {
            array.Items.Add(new PdfStringObj(value, useTextStringEncoding: true));
        }

        return array;
    }

    private static bool TryCreateIncrementalDefaultAppearanceFontResources(Dictionary<int, PdfIndirectObject> objects, string? defaultAppearance, out string fontResourceName, out PdfDictionary resources, params PdfDictionary?[] candidateResources) {
        fontResourceName = IncrementalDefaultAppearanceFontName;
        resources = null!;
        if (!PdfDefaultAppearanceParser.TryReadFontResourceName(defaultAppearance, out string defaultAppearanceFontName) ||
            !TryFindFontResourceObject(objects, defaultAppearanceFontName, out PdfObject? fontObject, candidateResources)) {
            return false;
        }

        fontResourceName = defaultAppearanceFontName;
        resources = CreateIncrementalAppearanceResources(defaultAppearanceFontName, fontObject!);
        return true;
    }

    private static bool TryFindFontResource(Dictionary<int, PdfIndirectObject> objects, string name, out PdfReference reference, params PdfDictionary?[] candidateResources) {
        reference = null!;
        if (!TryFindFontResourceObject(objects, name, out PdfObject? fontObject, candidateResources) ||
            fontObject is not PdfReference fontReference) {
            return false;
        }

        reference = fontReference;
        return true;
    }

    private static bool TryFindFontResourceObject(Dictionary<int, PdfIndirectObject> objects, string name, out PdfObject? fontObject, params PdfDictionary?[] candidateResources) {
        fontObject = null;
        var seen = new List<PdfDictionary>();
        foreach (PdfDictionary? candidateResourcesEntry in candidateResources) {
            if (candidateResourcesEntry is null || seen.Any(item => ReferenceEquals(item, candidateResourcesEntry))) {
                continue;
            }

            seen.Add(candidateResourcesEntry);
            if (ResolveDictionary(objects, candidateResourcesEntry.Items.TryGetValue("Font", out PdfObject? fontsObject) ? fontsObject : null) is PdfDictionary fonts &&
                fonts.Items.TryGetValue(name, out fontObject)) {
                return true;
            }
        }

        return false;
    }

    private static PdfDictionary? TryReadIncrementalNormalAppearanceResources(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget) {
        if (!TryGetIncrementalNormalAppearanceObject(objects, widget, out PdfObject? normalAppearance) ||
            ResolveObject(objects, normalAppearance) is not PdfStream normalAppearanceStream ||
            !normalAppearanceStream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject)) {
            return null;
        }

        return ResolveDictionary(objects, resourcesObject);
    }

    private static PdfDictionary? TryReadIncrementalWidgetPageResources(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget) {
        return ResolveDictionary(objects, widget.Items.TryGetValue("P", out PdfObject? pageObject) ? pageObject : null) is PdfDictionary page
            ? TryReadIncrementalInheritedPageResources(objects, page)
            : null;
    }

    private static PdfDictionary? TryReadIncrementalInheritedPageResources(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page) {
        PdfDictionary? current = page;
        int guard = 0;
        while (current is not null && guard++ < 100) {
            if (current.Items.TryGetValue("Resources", out PdfObject? resourcesObject) &&
                ResolveDictionary(objects, resourcesObject) is PdfDictionary resources) {
                return resources;
            }

            current = current.Items.TryGetValue("Parent", out PdfObject? parentObject) &&
                parentObject is PdfReference parentReference &&
                PdfObjectLookup.TryGet(objects, parentReference, out PdfIndirectObject? parentIndirect) &&
                parentIndirect.Value is PdfDictionary parent
                    ? parent
                    : null;
        }

        return null;
    }

    private static PdfFormFieldStyle ReadIncrementalWidgetAppearanceStyle(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, int fieldFlags = 0, int? inheritedQuadding = null, int? inheritedMaxLength = null, string? inheritedDefaultAppearance = null) {
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

        if (TryReadWidgetBorderWidth(objects, widget, out double borderWidth)) {
            style.BorderWidth = borderWidth;
        }

        if (TryReadWidgetBorderStyle(objects, widget, out PdfFormFieldBorderStyle borderStyle)) {
            style.BorderStyle = borderStyle;
        }

        if (TryReadWidgetBorderDashPattern(objects, widget, out IReadOnlyList<double>? borderDashPattern)) {
            style.BorderDashPattern = borderDashPattern;
        }

        if (TryReadDefaultAppearanceTextColor(objects, widget, inheritedDefaultAppearance, out PdfColor textColor)) {
            style.TextColor = textColor;
        }

        if (TryReadWidgetTextAlignment(objects, widget, inheritedQuadding, out PdfFormFieldTextAlignment textAlignment)) {
            style.TextAlignment = textAlignment;
        }

        return style;
    }

    private static HashSet<string> CollectIncrementalButtonNormalAppearanceStates(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, HashSet<int> visited) {
        var states = new HashSet<string>(StringComparer.Ordinal);
        CollectIncrementalButtonNormalAppearanceStates(objects, field, states, visited);
        states.Remove("Off");
        return states;
    }

    private static void CollectIncrementalButtonNormalAppearanceStates(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, HashSet<string> states, HashSet<int> visited) {
        if (IsWidget(field) &&
            TryGetIncrementalNormalAppearanceObject(objects, field, out PdfObject? normalAppearance) &&
            normalAppearance is PdfDictionary appearanceStates) {
            foreach (string stateName in appearanceStates.Items.Keys) {
                states.Add(stateName);
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
                CollectIncrementalButtonNormalAppearanceStates(objects, kid, states, visited);
            }
        }
    }

    private static string? ReadIncrementalWidgetOnAppearanceState(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget) {
        if (!TryGetIncrementalNormalAppearanceObject(objects, widget, out PdfObject? normalAppearance) ||
            normalAppearance is not PdfDictionary appearanceStates) {
            return null;
        }

        foreach (string stateName in appearanceStates.Items.Keys) {
            if (!string.Equals(stateName, "Off", StringComparison.Ordinal)) {
                return stateName;
            }
        }

        return null;
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

    private static bool TryReadWidgetBorderWidth(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, out double borderWidth) {
        borderWidth = 0D;
        if (ResolveDictionary(objects, widget.Items.TryGetValue("BS", out PdfObject? borderStyleObject) ? borderStyleObject : null) is PdfDictionary borderStyle &&
            borderStyle.Items.TryGetValue("W", out PdfObject? borderStyleWidthObject) &&
            TryReadNonNegativeFiniteNumber(objects, borderStyleWidthObject, out borderWidth)) {
            return true;
        }

        if (widget.Items.TryGetValue("Border", out PdfObject? borderObject) &&
            ResolveObject(objects, borderObject) is PdfArray border &&
            border.Items.Count >= 3 &&
            TryReadNonNegativeFiniteNumber(objects, border.Items[2], out borderWidth)) {
            return true;
        }

        return false;
    }

    private static bool TryReadWidgetBorderStyle(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, out PdfFormFieldBorderStyle borderStyle) {
        borderStyle = PdfFormFieldBorderStyle.Solid;
        if (ResolveDictionary(objects, widget.Items.TryGetValue("BS", out PdfObject? borderStyleObject) ? borderStyleObject : null) is not PdfDictionary borderStyleDictionary ||
            borderStyleDictionary.Get<PdfName>("S") is not PdfName styleName) {
            return false;
        }

        switch (styleName.Name) {
            case "D":
                borderStyle = PdfFormFieldBorderStyle.Dashed;
                return true;
            case "U":
                borderStyle = PdfFormFieldBorderStyle.Underline;
                return true;
            case "B":
                borderStyle = PdfFormFieldBorderStyle.Beveled;
                return true;
            case "I":
                borderStyle = PdfFormFieldBorderStyle.Inset;
                return true;
            case "S":
                borderStyle = PdfFormFieldBorderStyle.Solid;
                return true;
            default:
                return false;
        }
    }

    private static bool TryReadWidgetBorderDashPattern(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, out IReadOnlyList<double>? borderDashPattern) {
        borderDashPattern = null;
        if (ResolveDictionary(objects, widget.Items.TryGetValue("BS", out PdfObject? borderStyleObject) ? borderStyleObject : null) is not PdfDictionary borderStyle ||
            borderStyle.Get<PdfName>("S")?.Name != "D") {
            return false;
        }

        if (!borderStyle.Items.TryGetValue("D", out PdfObject? dashObject)) {
            borderDashPattern = new[] { 3D };
            return true;
        }

        return TryReadDashPattern(objects, dashObject, out borderDashPattern);
    }

    private static bool TryReadDashPattern(Dictionary<int, PdfIndirectObject> objects, PdfObject dashObject, out IReadOnlyList<double>? dashPattern) {
        dashPattern = null;
        if (ResolveObject(objects, dashObject) is not PdfArray dashArray || dashArray.Items.Count == 0) {
            return false;
        }

        var values = new double[dashArray.Items.Count];
        bool hasPositiveSegment = false;
        for (int i = 0; i < dashArray.Items.Count; i++) {
            if (!TryReadNonNegativeFiniteNumber(objects, dashArray.Items[i], out double segment)) {
                return false;
            }

            if (segment > 0D) {
                hasPositiveSegment = true;
            }

            values[i] = segment;
        }

        if (!hasPositiveSegment) {
            return false;
        }

        dashPattern = values;
        return true;
    }

    private static bool TryReadNonNegativeFiniteNumber(Dictionary<int, PdfIndirectObject> objects, PdfObject numberObject, out double value) {
        value = 0D;
        if (ResolveObject(objects, numberObject) is not PdfNumber number ||
            number.Value < 0D ||
            double.IsNaN(number.Value) ||
            double.IsInfinity(number.Value)) {
            return false;
        }

        value = number.Value;
        return true;
    }

    private static bool TryReadDefaultAppearanceTextColor(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, string? inheritedDefaultAppearance, out PdfColor color) {
        return PdfDefaultAppearanceParser.TryReadTextColor(TryReadText(objects, widget, "DA") ?? inheritedDefaultAppearance, out color);
    }

    private static double ReadIncrementalWidgetAppearanceFontSize(string? defaultAppearance, double height) {
        return PdfDefaultAppearanceParser.TryReadFontSize(defaultAppearance, out double fontSize)
            ? fontSize
            : Math.Max(6D, Math.Min(12D, height - 4D));
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
        HashSet<int> changedObjectNumbers,
        PdfStandardSecurityHandler? encryptionHandler) {
        return PdfIncrementalObjectWriter.Append(
            pdf,
            objects,
            security,
            trailerRaw,
            changedObjectNumbers,
            encryptionHandler: encryptionHandler);
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
