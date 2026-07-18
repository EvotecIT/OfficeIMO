namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    private static void FillField(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject fieldObject,
        string? parentName,
        string? inheritedFieldType,
        int inheritedFlags,
        int? inheritedQuadding,
        int? inheritedMaxLength,
        PdfDictionary? inheritedDefaultResources,
        string? inheritedDefaultAppearance,
        PdfArray? inheritedChoiceOptions,
        IReadOnlyDictionary<string, PdfFormFieldValue> fieldValues,
        PdfFormFillerOptions? options,
        HashSet<string> remaining,
        HashSet<int> visited,
        ref int nextObjectNumber) {
        if (fieldObject is PdfReference reference && !visited.Add(reference.ObjectNumber)) {
            return;
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
        PdfArray? choiceOptions = TryReadChoiceOptions(objects, field) ?? inheritedChoiceOptions;
        if (fullName is not null && remaining.Contains(fullName) && fieldValues.TryGetValue(fullName, out PdfFormFieldValue? value)) {
            SetFieldValue(objects, field, fullName, fieldType, fieldFlags, fieldQuadding, fieldMaxLength, defaultResources, defaultAppearance, choiceOptions, value, options, ref nextObjectNumber);
            remaining.Remove(fullName);
        }

        if (!field.Items.TryGetValue("Kids", out var kidsObject) ||
            ResolveObject(objects, kidsObject) is not PdfArray kids) {
            return;
        }

        for (int i = 0; i < kids.Items.Count; i++) {
            FillField(objects, kids.Items[i], fullName, fieldType, fieldFlags, fieldQuadding, fieldMaxLength, defaultResources, defaultAppearance, choiceOptions, fieldValues, options, remaining, visited, ref nextObjectNumber);
        }
    }

    private static void SetFieldValue(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, string fieldName, string? fieldType, int fieldFlags, int? inheritedQuadding, int? inheritedMaxLength, PdfDictionary? inheritedDefaultResources, string? inheritedDefaultAppearance, PdfArray? choiceOptions, PdfFormFieldValue value, PdfFormFillerOptions? options, ref int nextObjectNumber) {
        IReadOnlyList<string> values = value.Values;
        string firstValue = values[0];
        if (string.Equals(fieldType, "Btn", StringComparison.Ordinal)) {
            string name = string.IsNullOrEmpty(firstValue) ? "Off" : firstValue;
            bool isRadioButtonGroup = (fieldFlags & RadioButtonFlag) != 0;
            if (isRadioButtonGroup) {
                if (values.Count > 1) {
                    throw new ArgumentException("PDF radio button field cannot be filled with multiple values.", nameof(value));
                }

                if (!string.Equals(name, "Off", StringComparison.Ordinal)) {
                    HashSet<string> availableStates = CollectButtonNormalAppearanceStates(objects, field, new HashSet<int>());
                    if (!availableStates.Contains(name)) {
                        throw new ArgumentException($"PDF radio button field cannot be filled with value '{name}' because it is not one of the available appearance states.", nameof(value));
                    }
                }
            }

            field.Items["V"] = new PdfName(name);
            field.Items["AS"] = new PdfName(name);
            SetWidgetAppearanceStates(objects, field, name, isRadioButtonGroup, new HashSet<int>(), ref nextObjectNumber);
            return;
        }

        if (string.Equals(fieldType, "Ch", StringComparison.Ordinal)) {
            bool isMultiSelectChoice = (fieldFlags & MultiSelectChoiceFlag) != 0;
            if (values.Count > 1 && !isMultiSelectChoice) {
                throw new ArgumentException("PDF scalar choice field cannot be filled with multiple values.", nameof(value));
            }

            IReadOnlyList<ChoiceFillValue> choiceValues = ResolveChoiceFillValues(objects, choiceOptions, (fieldFlags & EditableChoiceFlag) != 0, values);
            if (isMultiSelectChoice) {
                field.Items["V"] = CreateStringArray(choiceValues.Select(item => item.ExportValue));
                SetTextWidgetAppearances(objects, field, string.Join("\n", choiceValues.Select(item => item.DisplayValue)), fieldName, fieldFlags, inheritedQuadding, inheritedMaxLength, inheritedDefaultResources, inheritedDefaultAppearance, true, options, new HashSet<int>(), ref nextObjectNumber);
                return;
            }

            ChoiceFillValue choiceValue = choiceValues[0];
            field.Items["V"] = new PdfStringObj(choiceValue.ExportValue, useTextStringEncoding: true);
            SetTextWidgetAppearances(objects, field, choiceValue.DisplayValue, fieldName, fieldFlags, inheritedQuadding, inheritedMaxLength, inheritedDefaultResources, inheritedDefaultAppearance, false, options, new HashSet<int>(), ref nextObjectNumber);
            return;
        }

        field.Items["V"] = new PdfStringObj(firstValue, useTextStringEncoding: true);
        SetTextWidgetAppearances(objects, field, firstValue, fieldName, fieldFlags, inheritedQuadding, inheritedMaxLength, inheritedDefaultResources, inheritedDefaultAppearance, false, options, new HashSet<int>(), ref nextObjectNumber);
    }

    private static int ReadFieldFlags(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field, int inheritedFlags) {
        if (!field.Items.TryGetValue("Ff", out PdfObject? flagsObject) ||
            ResolveObject(objects, flagsObject) is not PdfNumber flagsNumber) {
            return inheritedFlags;
        }

        return (int)flagsNumber.Value;
    }

    private static PdfArray? TryReadChoiceOptions(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field) {
        if (!field.Items.TryGetValue("Opt", out PdfObject? optionsObject)) {
            return null;
        }

        return ResolveObject(objects, optionsObject) as PdfArray;
    }
}
