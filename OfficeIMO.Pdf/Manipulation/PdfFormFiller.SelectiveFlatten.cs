namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    private static void ValidateFlattenFieldNames(IReadOnlyCollection<string> fieldNames) {
        Guard.NotNull(fieldNames, nameof(fieldNames));
        if (fieldNames.Count == 0) throw new ArgumentException("At least one form field name is required.", nameof(fieldNames));
        foreach (string name in fieldNames) Guard.NotNullOrWhiteSpace(name, nameof(fieldNames));
    }

    private static void CollectSelectedFlattenFields(
        Dictionary<int, PdfIndirectObject> objects,
        PdfArray owner,
        HashSet<string> requested,
        HashSet<string> matched,
        string? inheritedFieldType,
        int inheritedFlags,
        int? inheritedMaxLength,
        PdfDictionary? inheritedDefaultResources,
        string? inheritedDefaultAppearance,
        string? inheritedDisplayValue,
        IReadOnlyList<PdfFreeTextRichTextRun>? inheritedRichAppearanceRuns,
        string? inheritedName,
        PdfArray? inheritedChoiceOptions,
        PdfFormFillerOptions? options,
        Dictionary<int, FlattenWidgetState> widgets,
        HashSet<int> removableObjects,
        ref int nextObjectNumber) {
        for (int i = owner.Items.Count - 1; i >= 0; i--) {
            PdfObject fieldObject = owner.Items[i];
            if (ResolveObject(objects, fieldObject) is not PdfDictionary field) continue;

            string? partialName = TryReadText(objects, field, "T");
            string? fullName = CombineFieldName(inheritedName, partialName);
            string? fieldType = TryReadName(objects, field, "FT") ?? inheritedFieldType;
            int fieldFlags = ReadFieldFlags(objects, field, inheritedFlags);
            int? fieldMaxLength = ReadFieldMaxLength(objects, field, inheritedMaxLength);
            PdfDictionary? defaultResources = TryReadDefaultResources(objects, field) ?? inheritedDefaultResources;
            string? defaultAppearance = TryReadText(objects, field, "DA") ?? inheritedDefaultAppearance;
            PdfArray? choiceOptions = TryReadChoiceOptions(objects, field) ?? inheritedChoiceOptions;
            IReadOnlyList<string>? values = TryReadSimpleValues(objects, field, "V");
            string? richValue = TryReadText(objects, field, "RV");
            string? richPlainText = PdfFreeTextStyleParser.ExtractPlainText(richValue);
            IReadOnlyList<PdfFreeTextRichTextRun>? richRuns = PdfFreeTextStyleParser.ExtractRichTextRuns(richValue) ?? inheritedRichAppearanceRuns;
            string? displayValue = values is { Count: > 0 } ? values[0] : richPlainText ?? inheritedDisplayValue;

            if (fullName is not null && fullName.Length > 0 && requested.Contains(fullName)) {
                CollectFlattenWidgets(objects, fieldObject, inheritedFieldType, inheritedFlags, inheritedMaxLength, inheritedDefaultResources, inheritedDefaultAppearance, inheritedDisplayValue, inheritedRichAppearanceRuns, inheritedName, inheritedChoiceOptions, options, widgets, removableObjects, new HashSet<int>(), ref nextObjectNumber);
                owner.Items.RemoveAt(i);
                matched.Add(fullName!);
                continue;
            }

            if (!field.Items.TryGetValue("Kids", out PdfObject? kidsObject) || ResolveObject(objects, kidsObject) is not PdfArray kids) continue;
            CollectSelectedFlattenFields(objects, kids, requested, matched, fieldType, fieldFlags, fieldMaxLength, defaultResources, defaultAppearance, displayValue, richRuns, fullName, choiceOptions, options, widgets, removableObjects, ref nextObjectNumber);
            if (kids.Items.Count == 0 && !IsWidget(field)) {
                owner.Items.RemoveAt(i);
                if (fieldObject is PdfReference reference) removableObjects.Add(reference.ObjectNumber);
            }
        }
    }

    private static void FilterCalculationOrder(Dictionary<int, PdfIndirectObject> objects, PdfDictionary acroForm, HashSet<int> removedObjects) {
        if (!acroForm.Items.TryGetValue("CO", out PdfObject? orderObject) || ResolveObject(objects, orderObject) is not PdfArray order) return;
        for (int i = order.Items.Count - 1; i >= 0; i--) {
            if (order.Items[i] is PdfReference reference && removedObjects.Contains(reference.ObjectNumber)) order.Items.RemoveAt(i);
        }
        if (order.Items.Count == 0) acroForm.Items.Remove("CO");
    }
}
