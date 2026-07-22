namespace OfficeIMO.Pdf;

public sealed partial class PdfReadDocument {
    private IReadOnlyList<PdfFormField> ExtractFormFields() {
        PdfDictionary? acroForm = GetAcroFormDictionary();
        if (acroForm is null ||
            !acroForm.Items.TryGetValue("Fields", out var fieldsObject) ||
            ResolveArray(fieldsObject) is not PdfArray fields) {
            return Array.Empty<PdfFormField>();
        }

        var result = new List<PdfFormField>();
        var visited = new HashSet<int>();
        var widgetPageNumbers = BuildWidgetPageNumberLookup();
        PdfFormFieldInheritedState inherited = PdfFormFieldInheritedState.FromAcroForm(_acroFormDefaultAppearance, _acroFormQuadding);
        if (fields.Items.Count > _options.Limits.MaxFormFields) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.FormFields, _options.Limits.MaxFormFields, fields.Items.Count);
        }

        for (int i = 0; i < fields.Items.Count; i++) {
            ReadFormField(fields.Items[i], null, inherited, result, visited, widgetPageNumbers, depth: 1);
        }

        return result.Count == 0 ? Array.Empty<PdfFormField>() : result.AsReadOnly();
    }

    private string? ExtractAcroFormText(string key) {
        PdfDictionary? acroForm = GetAcroFormDictionary();
        if (acroForm is null ||
            !acroForm.Items.TryGetValue(key, out var value) ||
            ResolveObject(value) is not PdfStringObj text ||
            string.IsNullOrEmpty(text.Value)) {
            return null;
        }

        return text.Value;
    }

    private bool? ExtractAcroFormBoolean(string key) {
        PdfDictionary? acroForm = GetAcroFormDictionary();
        if (acroForm is null ||
            !acroForm.Items.TryGetValue(key, out var value) ||
            ResolveObject(value) is not PdfBoolean boolean) {
            return null;
        }

        return boolean.Value;
    }

    private int? ExtractAcroFormInteger(string key) {
        PdfDictionary? acroForm = GetAcroFormDictionary();
        if (acroForm is null ||
            !acroForm.Items.TryGetValue(key, out var value) ||
            ResolveObject(value) is not PdfNumber number ||
            number.Value < int.MinValue ||
            number.Value > int.MaxValue ||
            Math.Truncate(number.Value) != number.Value) {
            return null;
        }

        return (int)number.Value;
    }

    private PdfAcroFormXfaInfo? ExtractAcroFormXfaInfo() {
        PdfDictionary? acroForm = GetAcroFormDictionary();
        if (acroForm is null ||
            !acroForm.Items.TryGetValue("XFA", out var xfaObject)) {
            return null;
        }

        int? objectNumber = xfaObject is PdfReference reference ? reference.ObjectNumber : null;
        PdfObject? resolved = ResolveObject(xfaObject);
        if (resolved is null || resolved is PdfNull) {
            return null;
        }

        return BuildAcroFormXfaInfo(resolved, objectNumber);
    }

    private PdfAcroFormXfaInfo BuildAcroFormXfaInfo(PdfObject xfaObject, int? objectNumber) {
        if (xfaObject is PdfArray array) {
            var packetNames = new List<string>();
            int streamCount = 0;
            int stringCount = 0;
            int dictionaryCount = 0;
            int totalPayloadBytes = 0;

            for (int i = 0; i < array.Items.Count; i++) {
                PdfObject? item = ResolveObject(array.Items[i]);
                if (item is PdfStringObj packetName &&
                    i + 1 < array.Items.Count) {
                    packetNames.Add(packetName.Value);
                    AddXfaPayloadStats(ResolveObject(array.Items[i + 1]), ref streamCount, ref stringCount, ref dictionaryCount, ref totalPayloadBytes);
                    i++;
                    continue;
                }

                AddXfaPayloadStats(item, ref streamCount, ref stringCount, ref dictionaryCount, ref totalPayloadBytes);
            }

            return new PdfAcroFormXfaInfo(
                "array",
                objectNumber,
                packetNames.Count,
                packetNames.AsReadOnly(),
                streamCount,
                stringCount,
                dictionaryCount,
                totalPayloadBytes,
                ContainsXfaPacket(packetNames, "template"),
                ContainsXfaPacket(packetNames, "datasets"));
        }

        int directStreamCount = 0;
        int directStringCount = 0;
        int directDictionaryCount = 0;
        int directPayloadBytes = 0;
        AddXfaPayloadStats(xfaObject, ref directStreamCount, ref directStringCount, ref directDictionaryCount, ref directPayloadBytes);

        return new PdfAcroFormXfaInfo(
            GetXfaObjectKind(xfaObject),
            objectNumber,
            0,
            Array.Empty<string>(),
            directStreamCount,
            directStringCount,
            directDictionaryCount,
            directPayloadBytes,
            false,
            false);
    }

    private static void AddXfaPayloadStats(PdfObject? value, ref int streamCount, ref int stringCount, ref int dictionaryCount, ref int totalPayloadBytes) {
        if (value is PdfStream stream) {
            streamCount++;
            totalPayloadBytes += stream.Data.Length;
        } else if (value is PdfStringObj text) {
            stringCount++;
            totalPayloadBytes += text.RawBytes.Length;
        } else if (value is PdfDictionary) {
            dictionaryCount++;
        }
    }

    private static bool ContainsXfaPacket(List<string> packetNames, string packetName) {
        for (int i = 0; i < packetNames.Count; i++) {
            if (string.Equals(packetNames[i], packetName, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private static string GetXfaObjectKind(PdfObject value) {
        if (value is PdfStream) return "stream";
        if (value is PdfStringObj) return "string";
        if (value is PdfDictionary) return "dictionary";
        if (value is PdfArray) return "array";
        if (value is PdfName) return "name";
        if (value is PdfNumber) return "number";
        if (value is PdfBoolean) return "boolean";
        if (value is PdfNull) return "null";
        return "unknown";
    }

    private PdfDictionary? GetAcroFormDictionary() {
        PdfDictionary? catalog = FindCatalog();
        if (catalog is null ||
            !catalog.Items.TryGetValue("AcroForm", out var acroFormObject) ||
            ResolveObject(acroFormObject) is not PdfDictionary acroForm) {
            return null;
        }

        return acroForm;
    }

    private Dictionary<int, int> BuildWidgetPageNumberLookup() {
        var widgetPageNumbers = new Dictionary<int, int>();
        for (int i = 0; i < Pages.Count; i++) {
            IReadOnlyList<int> annotationObjectNumbers = Pages[i].GetAnnotationObjectNumbers("Widget");
            for (int j = 0; j < annotationObjectNumbers.Count; j++) {
                if (!widgetPageNumbers.ContainsKey(annotationObjectNumbers[j])) {
                    widgetPageNumbers.Add(annotationObjectNumbers[j], i + 1);
                }
            }
        }

        return widgetPageNumbers;
    }

    private void ReadFormField(PdfObject fieldObject, string? parentName, PdfFormFieldInheritedState inherited, List<PdfFormField> result, HashSet<int> visited, IReadOnlyDictionary<int, int> widgetPageNumbers, int depth) {
        if (depth > _options.Limits.MaxFormFieldDepth) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.FormFieldDepth, _options.Limits.MaxFormFieldDepth, depth);
        }

        PdfObject? resolved = ResolveObject(fieldObject);
        if (resolved is not PdfDictionary field) {
            return;
        }

        int? objectNumber = null;
        if (fieldObject is PdfReference reference) {
            objectNumber = reference.ObjectNumber;
            if (!visited.Add(reference.ObjectNumber)) {
                return;
            }
        } else {
            int foundObjectNumber = FindExactObjectNumberFor(field);
            if (foundObjectNumber > 0) {
                objectNumber = foundObjectNumber;
                if (!visited.Add(foundObjectNumber)) {
                    return;
                }
            }
        }

        if (visited.Count > _options.Limits.MaxFormFields) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.FormFields, _options.Limits.MaxFormFields, visited.Count);
        }

        string? partialName = TryReadText(field, "T");
        string? fullName = CombineFieldName(parentName, partialName);
        string? fieldType = TryReadName(field, "FT") ?? inherited.FieldType;
        string? value = field.Items.ContainsKey("V") ? TryReadSimpleFieldValue(field, "V") : inherited.Value;
        IReadOnlyList<string> values = field.Items.ContainsKey("V") ? ReadSimpleFieldValues(field, "V") : inherited.Values;
        string? defaultValue = field.Items.ContainsKey("DV") ? TryReadSimpleFieldValue(field, "DV") : inherited.DefaultValue;
        IReadOnlyList<string> defaultValues = field.Items.ContainsKey("DV") ? ReadSimpleFieldValues(field, "DV") : inherited.DefaultValues;
        string? alternateName = TryReadText(field, "TU");
        string? mappingName = TryReadText(field, "TM");
        int? flags = field.Items.ContainsKey("Ff") ? TryReadInteger(field, "Ff") : inherited.Flags;
        int? maxLength = field.Items.ContainsKey("MaxLen") ? TryReadPositiveInteger(field, "MaxLen") : inherited.MaxLength;
        string? defaultAppearance = field.Items.ContainsKey("DA") ? TryReadText(field, "DA") : inherited.DefaultAppearance;
        int? quadding = field.Items.ContainsKey("Q") ? TryReadInteger(field, "Q") : inherited.Quadding;
        IReadOnlyList<PdfFormFieldOption> options = field.Items.ContainsKey("Opt") ? ReadFormFieldOptions(field) : inherited.Options;
        bool isWidget = IsWidget(field);
        var widgets = new List<PdfFormWidget>();
        if (TryReadFormWidget(field, fullName, objectNumber, widgetPageNumbers, out PdfFormWidget? widget) && widget is not null) {
            widgets.Add(widget);
        }

        PdfArray? kids = field.Items.TryGetValue("Kids", out var kidsObject) ? ResolveArray(kidsObject) : null;
        bool hasReadableFieldState = fieldType != null || value != null || defaultValue != null || flags.HasValue;
        var fieldKids = new List<PdfObject>();
        if (kids is not null) {
            for (int i = 0; i < kids.Items.Count; i++) {
                PdfObject kidObject = kids.Items[i];
                PdfDictionary? kid = ResolveObject(kidObject) as PdfDictionary;
                if (kid is not null && IsWidget(kid) && !HasOwnFieldName(kid)) {
                    int? kidObjectNumber = TryGetObjectNumber(kidObject, kid);
                    if (TryReadFormWidget(kid, fullName, kidObjectNumber, widgetPageNumbers, out PdfFormWidget? kidWidget) && kidWidget is not null) {
                        widgets.Add(kidWidget);
                    }

                    continue;
                }

                fieldKids.Add(kidObject);
            }
        }

        bool hasTerminalShape = isWidget || fieldKids.Count == 0;
        if (hasTerminalShape && (fullName != null || hasReadableFieldState || defaultValues.Count > 0 || alternateName != null || mappingName != null || maxLength.HasValue || defaultAppearance != null || quadding.HasValue || options.Count > 0)) {
            if (result.Count >= _options.Limits.MaxFormFields) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.FormFields, _options.Limits.MaxFormFields, result.Count + 1L);
            }

            result.Add(new PdfFormField(
                objectNumber: objectNumber,
                name: fullName,
                partialName: partialName,
                fieldType: fieldType,
                value: value,
                alternateName: alternateName,
                mappingName: mappingName,
                flags: flags,
                maxLength: maxLength,
                values: values.Count == 0 ? null : values,
                defaultValue: defaultValue,
                defaultValues: defaultValues.Count == 0 ? null : defaultValues,
                defaultAppearance: defaultAppearance,
                quadding: quadding,
                options: options.Count == 0 ? null : options,
                widgets: widgets.Count == 0 ? null : widgets.AsReadOnly()));
        }

        if (fieldKids.Count == 0) {
            return;
        }

        var childInherited = new PdfFormFieldInheritedState(fieldType, value, values, defaultValue, defaultValues, flags, maxLength, defaultAppearance, quadding, options);
        for (int i = 0; i < fieldKids.Count; i++) {
            ReadFormField(fieldKids[i], fullName, childInherited, result, visited, widgetPageNumbers, depth + 1);
        }
    }

    private sealed class PdfFormFieldInheritedState {
        internal static readonly PdfFormFieldInheritedState Empty = new PdfFormFieldInheritedState(null, null, Array.Empty<string>(), null, Array.Empty<string>(), null, null, null, null, Array.Empty<PdfFormFieldOption>());

        internal static PdfFormFieldInheritedState FromAcroForm(string? defaultAppearance, int? quadding) {
            return string.IsNullOrEmpty(defaultAppearance) && !quadding.HasValue
                ? Empty
                : new PdfFormFieldInheritedState(null, null, Array.Empty<string>(), null, Array.Empty<string>(), null, null, defaultAppearance, quadding, Array.Empty<PdfFormFieldOption>());
        }

        internal PdfFormFieldInheritedState(string? fieldType, string? value, IReadOnlyList<string> values, string? defaultValue, IReadOnlyList<string> defaultValues, int? flags, int? maxLength, string? defaultAppearance, int? quadding, IReadOnlyList<PdfFormFieldOption> options) {
            FieldType = fieldType;
            Value = value;
            Values = values;
            DefaultValue = defaultValue;
            DefaultValues = defaultValues;
            Flags = flags;
            MaxLength = maxLength;
            DefaultAppearance = defaultAppearance;
            Quadding = quadding;
            Options = options;
        }

        internal string? FieldType { get; }

        internal string? Value { get; }

        internal IReadOnlyList<string> Values { get; }

        internal string? DefaultValue { get; }

        internal IReadOnlyList<string> DefaultValues { get; }

        internal int? Flags { get; }

        internal int? MaxLength { get; }

        internal string? DefaultAppearance { get; }

        internal int? Quadding { get; }

        internal IReadOnlyList<PdfFormFieldOption> Options { get; }
    }

    private int? TryGetObjectNumber(PdfObject sourceObject, PdfDictionary resolvedDictionary) {
        if (sourceObject is PdfReference reference) {
            return reference.ObjectNumber;
        }

        int foundObjectNumber = FindExactObjectNumberFor(resolvedDictionary);
        return foundObjectNumber > 0 ? foundObjectNumber : null;
    }

    private bool TryReadFormWidget(PdfDictionary dictionary, string? fieldName, int? objectNumber, IReadOnlyDictionary<int, int> widgetPageNumbers, out PdfFormWidget? widget) {
        widget = null;
        if (!IsWidget(dictionary) ||
            !TryReadRectangle(dictionary.Items.TryGetValue("Rect", out var rectObject) ? rectObject : null, out var rect)) {
            return false;
        }

        int? pageNumber = null;
        if (objectNumber.HasValue && widgetPageNumbers.TryGetValue(objectNumber.Value, out int foundPageNumber)) {
            pageNumber = foundPageNumber;
        }

        widget = new PdfFormWidget(
            objectNumber,
            fieldName,
            pageNumber,
            rect.X1,
            rect.Y1,
            rect.X2,
            rect.Y2,
            TryReadName(dictionary, "AS"),
            TryReadInteger(dictionary, "F"),
            ReadWidgetNormalAppearanceStates(dictionary));
        return true;
    }

    private IReadOnlyList<string> ReadWidgetNormalAppearanceStates(PdfDictionary dictionary) {
        if (!dictionary.Items.TryGetValue("AP", out var appearancesObject) ||
            ResolveObject(appearancesObject) is not PdfDictionary appearances ||
            !appearances.Items.TryGetValue("N", out var normalObject) ||
            ResolveObject(normalObject) is not PdfDictionary normalAppearances ||
            normalAppearances.Items.Count == 0) {
            return Array.Empty<string>();
        }

        if (normalAppearances.Items.Count > _options.Limits.MaxFormFieldAppearanceStates) {
            throw PdfReadLimitException.Create(
                PdfReadLimitKind.FormAppearanceStates,
                _options.Limits.MaxFormFieldAppearanceStates,
                normalAppearances.Items.Count);
        }

        var states = new List<string>();
        foreach (string state in normalAppearances.Items.Keys) {
            if (!string.IsNullOrEmpty(state)) {
                states.Add(state);
            }
        }

        if (states.Count == 0) {
            return Array.Empty<string>();
        }

        states.Sort(StringComparer.Ordinal);
        return states.AsReadOnly();
    }

    private static bool IsWidget(PdfDictionary dictionary) {
        return dictionary.Items.TryGetValue("Subtype", out var subtype) &&
            subtype is PdfName name &&
            string.Equals(name.Name, "Widget", StringComparison.Ordinal);
    }

    private bool HasOwnFieldName(PdfDictionary dictionary) {
        return TryReadText(dictionary, "T") is not null;
    }

    private bool TryReadRectangle(PdfObject? obj, out (double X1, double Y1, double X2, double Y2) rect) {
        rect = default;
        var array = ResolveArray(obj);
        if (array is null || array.Items.Count < 4) {
            return false;
        }

        if (ResolveObject(array.Items[0]) is not PdfNumber x1 ||
            ResolveObject(array.Items[1]) is not PdfNumber y1 ||
            ResolveObject(array.Items[2]) is not PdfNumber x2 ||
            ResolveObject(array.Items[3]) is not PdfNumber y2) {
            return false;
        }

        double left = Math.Min(x1.Value, x2.Value);
        double right = Math.Max(x1.Value, x2.Value);
        double bottom = Math.Min(y1.Value, y2.Value);
        double top = Math.Max(y1.Value, y2.Value);
        if (double.IsNaN(left) || double.IsInfinity(left) ||
            double.IsNaN(right) || double.IsInfinity(right) ||
            double.IsNaN(bottom) || double.IsInfinity(bottom) ||
            double.IsNaN(top) || double.IsInfinity(top) ||
            right <= left ||
            top <= bottom) {
            return false;
        }

        rect = (left, bottom, right, top);
        return true;
    }
}
