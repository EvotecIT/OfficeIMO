namespace OfficeIMO.Pdf;

internal static partial class PdfIncrementalUpdater {
    private readonly struct IncrementalChoiceFillValue {
        public IncrementalChoiceFillValue(string exportValue, string displayValue, int? optionIndex) {
            ExportValue = exportValue;
            DisplayValue = displayValue;
            OptionIndex = optionIndex;
        }

        public string ExportValue { get; }
        public string DisplayValue { get; }
        public int? OptionIndex { get; }
    }

    private readonly struct IncrementalPreparedFieldValue {
        private IncrementalPreparedFieldValue(string[] storedValues, string appearanceValue, bool forceMultilineAppearance, int[] selectedOptionIndices) {
            StoredValues = storedValues;
            AppearanceValue = appearanceValue;
            ForceMultilineAppearance = forceMultilineAppearance;
            SelectedOptionIndices = selectedOptionIndices;
        }

        public string[] StoredValues { get; }
        public string FirstStoredValue => StoredValues[0];
        public string AppearanceValue { get; }
        public bool ForceMultilineAppearance { get; }
        public int[] SelectedOptionIndices { get; }

        public static IncrementalPreparedFieldValue Scalar(string storedValue, string appearanceValue) =>
            new IncrementalPreparedFieldValue(new[] { storedValue }, appearanceValue, forceMultilineAppearance: false, Array.Empty<int>());

        public static IncrementalPreparedFieldValue Choice(string[] storedValues, string appearanceValue, bool forceMultilineAppearance, int[] selectedOptionIndices) =>
            new IncrementalPreparedFieldValue(storedValues, appearanceValue, forceMultilineAppearance, selectedOptionIndices);
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

        if (string.Equals(fieldType, "Ch", StringComparison.Ordinal)) {
            bool isMultiSelectChoice = (fieldFlags & IncrementalMultiSelectChoiceFlag) != 0;
            field.Items["V"] = isMultiSelectChoice
                ? CreateIncrementalStringArray(value.StoredValues)
                : new PdfStringObj(value.FirstStoredValue, useTextStringEncoding: true);
            SetIncrementalChoiceSelectionIndices(field, fieldFlags, value.SelectedOptionIndices);
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
            return IncrementalPreparedFieldValue.Choice(
                choiceValues.Select(item => item.ExportValue).ToArray(),
                string.Join("\n", choiceValues.Select(item => item.DisplayValue)),
                forceMultilineAppearance: true,
                choiceValues.All(item => item.OptionIndex.HasValue)
                    ? choiceValues.Select(item => item.OptionIndex!.Value).ToArray()
                    : Array.Empty<int>());
        }

        IncrementalChoiceFillValue choiceValue = choiceValues[0];
        return IncrementalPreparedFieldValue.Choice(
            new[] { choiceValue.ExportValue },
            choiceValue.DisplayValue,
            forceMultilineAppearance: false,
            choiceValue.OptionIndex.HasValue ? new[] { choiceValue.OptionIndex.Value } : Array.Empty<int>());
    }

    private static IReadOnlyList<IncrementalChoiceFillValue> ResolveIncrementalChoiceFillValues(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, bool isEditableChoice, IReadOnlyList<string> values) {
        if (!field.Items.TryGetValue("Opt", out PdfObject? optionsObject) ||
            ResolveObject(objects, optionsObject) is not PdfArray options ||
            options.Items.Count == 0) {
            return values.Select(static item => new IncrementalChoiceFillValue(item, item, null)).ToArray();
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
                    return new IncrementalChoiceFillValue(exportValue, displayValue, i);
                }

                continue;
            }

            if (optionObject is not null &&
                TryReadOptionText(objects, optionObject, out string? optionValue) &&
                optionValue is not null &&
                string.Equals(value, optionValue, StringComparison.Ordinal)) {
                return new IncrementalChoiceFillValue(optionValue, optionValue, i);
            }
        }

        if (isEditableChoice) {
            return new IncrementalChoiceFillValue(value, value, null);
        }

        throw new ArgumentException($"PDF choice field cannot be filled with value '{value}' because it is not one of the allowed options.", nameof(value));
    }

    private static void SetIncrementalChoiceSelectionIndices(PdfDictionary field, int fieldFlags, int[] selectedIndices) {
        if ((fieldFlags & IncrementalComboChoiceFlag) != 0 || selectedIndices.Length == 0) {
            field.Items.Remove("I");
            field.Items.Remove("TI");
            return;
        }

        var indices = new PdfArray();
        for (int i = 0; i < selectedIndices.Length; i++) {
            indices.Items.Add(new PdfNumber(selectedIndices[i]));
        }

        field.Items["I"] = indices;
        field.Items["TI"] = new PdfNumber(selectedIndices[0]);
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

        if (!isRadioButtonGroup && IsTruthyButtonValue(value) && availableStates.Count == 1) {
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
}
