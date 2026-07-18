namespace OfficeIMO.Pdf;

internal static partial class PdfMerger {
    private static byte[] ApplyFormPolicy(
        byte[] merged,
        IReadOnlyList<ImportedSource> sources,
        int primarySourceIndex,
        PdfMergeStructureMode mode,
        PdfMergeCollisionMode collisionMode,
        List<PdfMergeDecision> decisions) {
        int totalCount = sources.Sum(static source => source.Document.FormFields.Count);
        int incomingCount = sources.Where((source, index) => index != primarySourceIndex).Sum(static source => source.Document.FormFields.Count);
        if (totalCount == 0) {
            decisions.Add(new PdfMergeDecision("Forms", mode, "No AcroForm fields were present."));
            return merged;
        }
        if (mode == PdfMergeStructureMode.RejectIncoming && incomingCount > 0) {
            throw new InvalidOperationException("PDF merge policy rejected " + incomingCount + " incoming AcroForm field(s).");
        }
        if (mode == PdfMergeStructureMode.Combine) ValidateCombinableForms(sources);

        var renamed = new List<string>();
        int dropped = 0;
        IReadOnlyList<string> expectedNames;
        byte[] output = RewriteForms(merged, sources, primarySourceIndex, mode, collisionMode, renamed, ref dropped, out expectedNames);
        ValidateFormReadback(output, expectedNames);
        string action;
        int imported = 0;
        switch (mode) {
            case PdfMergeStructureMode.KeepPrimary: action = "Rebuilt the primary AcroForm and removed incoming widgets."; dropped = incomingCount; break;
            case PdfMergeStructureMode.RejectIncoming: action = "Rebuilt the primary AcroForm; no incoming fields were present."; break;
            case PdfMergeStructureMode.Drop: action = "Removed the AcroForm and all widget annotations."; break;
            case PdfMergeStructureMode.Combine: action = "Combined AcroForm field roots, widgets, values, and appearances."; imported = incomingCount - dropped; break;
            default: throw new ArgumentOutOfRangeException(nameof(mode));
        }
        decisions.Add(new PdfMergeDecision("Forms", mode, action, imported, dropped, renamed.AsReadOnly()));
        return output;
    }

    private static void ValidateCombinableForms(IReadOnlyList<ImportedSource> sources) {
        for (int sourceIndex = 0; sourceIndex < sources.Count; sourceIndex++) {
            PdfReadDocument document = sources[sourceIndex].Document;
            if (document.AcroFormXfa != null) throw new NotSupportedException("Combining XFA forms is not supported; flatten or remove XFA before merging.");
            if (document.FormFields.Any(static field => field.Kind == PdfFormFieldKind.Signature)) {
                throw new NotSupportedException("Combining signature fields is not supported because a full-rewrite merge cannot preserve their signature validity.");
            }
        }
    }

    private static int[] CollectAcroFormFieldRoots(
        Dictionary<int, PdfIndirectObject> objects,
        PdfReadDocument document,
        PdfPageExtractor.ObjectCollector collector) {
        if (!document.Security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(document.Security.RootObjectNumber.Value, out PdfIndirectObject? catalogObject) ||
            catalogObject.Value is not PdfDictionary catalog ||
            !catalog.Items.TryGetValue("AcroForm", out PdfObject? acroFormObject) ||
            ResolveDictionary(objects, acroFormObject) is not PdfDictionary acroForm ||
            !acroForm.Items.TryGetValue("Fields", out PdfObject? fieldsObject) ||
            ResolveObject(objects, fieldsObject) is not PdfArray fields) {
            return Array.Empty<int>();
        }

        collector.CollectObjectGraph(acroFormObject);
        return fields.Items
            .OfType<PdfReference>()
            .Select(static reference => reference.ObjectNumber)
            .Distinct()
            .ToArray();
    }

    private static byte[] RewriteForms(
        byte[] merged,
        IReadOnlyList<ImportedSource> sources,
        int primarySourceIndex,
        PdfMergeStructureMode mode,
        PdfMergeCollisionMode collisionMode,
        List<string> renamed,
        ref int dropped,
        out IReadOnlyList<string> expectedNames) {
        PdfReadDocument document = PdfReadDocument.Open(merged);
        var names = new List<string>();
        int localDropped = dropped;
        byte[] output = PdfDocumentObjectGraphRewriter.Rewrite(merged, null, null, (objects, security) => {
            PdfDictionary catalog = RequireCatalog(objects, security);
            List<MergedFormRoot> roots = FindFormRoots(objects, sources);
            var selected = new List<MergedFormRoot>();
            var fieldNames = new HashSet<string>(StringComparer.Ordinal);
            var droppedWidgetObjectNumbers = new HashSet<int>();

            foreach (MergedFormRoot root in roots) {
                bool include = mode == PdfMergeStructureMode.Combine ||
                    ((mode == PdfMergeStructureMode.KeepPrimary || mode == PdfMergeStructureMode.RejectIncoming) && root.SourceIndex == primarySourceIndex);
                if (!include) { foreach (int widget in root.WidgetObjectNumbers) droppedWidgetObjectNumbers.Add(widget); continue; }
                string fieldName = root.FieldName;
                if (!fieldNames.Add(fieldName)) {
                    if (collisionMode == PdfMergeCollisionMode.Reject) throw new InvalidOperationException("PDF AcroForm field name collision: " + fieldName);
                    if (collisionMode == PdfMergeCollisionMode.KeepFirst) {
                        foreach (int widget in root.WidgetObjectNumbers) droppedWidgetObjectNumbers.Add(widget);
                        localDropped++;
                        continue;
                    }
                    string renamedField = GetUniqueFormFieldName(fieldName, root.SourceIndex, fieldNames);
                    renamed.Add("source " + root.SourceIndex + ": " + fieldName + " -> " + renamedField);
                    fieldName = renamedField;
                    fieldNames.Add(fieldName);
                }
                root.FieldName = fieldName;
                PrepareFormRoot(objects, root);
                selected.Add(root);
                names.Add(fieldName);
            }

            RemoveDroppedWidgets(objects, document, droppedWidgetObjectNumbers);
            catalog.Items.Remove("AcroForm");
            if (selected.Count > 0) catalog.Items["AcroForm"] = BuildMergedAcroForm(selected, sources[primarySourceIndex].Document);
            return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value) ? security.InfoObjectNumber : null;
        });
        dropped = localDropped;
        expectedNames = names.OrderBy(static name => name, StringComparer.Ordinal).ToArray();
        return output;
    }

    private static List<MergedFormRoot> FindFormRoots(
        Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyList<ImportedSource> sources) {
        var byRoot = new Dictionary<int, MergedFormRoot>();
        var visited = new HashSet<int>();
        for (int sourceIndex = 0; sourceIndex < sources.Count; sourceIndex++) {
            ImportedSource source = sources[sourceIndex];
            IReadOnlyDictionary<int, int> numberMap = source.OutputNumberMap ??
                throw new InvalidOperationException("PDF merge form mapping was not initialized.");
            PdfObject? defaultResources = CloneMappedAcroFormDefaultResources(source, numberMap);
            for (int rootIndex = 0; rootIndex < source.FormFieldRootObjectNumbers.Length; rootIndex++) {
                if (numberMap.TryGetValue(source.FormFieldRootObjectNumbers[rootIndex], out int mappedRoot)) {
                    CollectFormTerminals(
                        objects,
                        sourceIndex,
                        mappedRoot,
                        string.Empty,
                        source.Document.AcroFormDefaultAppearance,
                        source.Document.AcroFormQuadding,
                        defaultResources,
                        byRoot,
                        visited);
                }
            }
        }

        return byRoot.Values.OrderBy(static root => root.SourceIndex).ThenBy(static root => root.RootObjectNumber).ToList();
    }

    private static void CollectFormTerminals(
        Dictionary<int, PdfIndirectObject> objects,
        int sourceIndex,
        int fieldObjectNumber,
        string parentName,
        string? sourceDefaultAppearance,
        int? sourceQuadding,
        PdfObject? sourceDefaultResources,
        Dictionary<int, MergedFormRoot> byRoot,
        HashSet<int> visited) {
        if (!visited.Add(fieldObjectNumber) ||
            !objects.TryGetValue(fieldObjectNumber, out PdfIndirectObject? indirect) ||
            indirect.Value is not PdfDictionary field) {
            return;
        }

        string partialName = field.Items.TryGetValue("T", out PdfObject? nameObject) && ResolveObject(objects, nameObject) is PdfStringObj name
            ? name.Value
            : string.Empty;
        string fullName = string.IsNullOrEmpty(parentName)
            ? partialName
            : string.IsNullOrEmpty(partialName) ? parentName : parentName + "." + partialName;
        var fieldChildren = new List<PdfReference>();
        var widgets = new List<int>();
        if (field.Items.TryGetValue("Kids", out PdfObject? kidsObject) && ResolveObject(objects, kidsObject) is PdfArray kids) {
            foreach (PdfObject kidObject in kids.Items) {
                if (kidObject is not PdfReference kidReference || ResolveDictionary(objects, kidReference) is not PdfDictionary kid) continue;
                bool pureWidget = kid.Get<PdfName>("Subtype")?.Name == "Widget" &&
                    !kid.Items.ContainsKey("T") &&
                    !kid.Items.ContainsKey("FT") &&
                    !kid.Items.ContainsKey("Kids");
                if (pureWidget) widgets.Add(kidReference.ObjectNumber); else fieldChildren.Add(kidReference);
            }
        }

        if (fieldChildren.Count > 0) {
            for (int childIndex = 0; childIndex < fieldChildren.Count; childIndex++) {
                CollectFormTerminals(
                    objects,
                    sourceIndex,
                    fieldChildren[childIndex].ObjectNumber,
                    fullName,
                    sourceDefaultAppearance,
                    sourceQuadding,
                    sourceDefaultResources,
                    byRoot,
                    visited);
            }

            return;
        }

        if (field.Get<PdfName>("Subtype")?.Name == "Widget") widgets.Add(fieldObjectNumber);
        if (string.IsNullOrEmpty(fullName)) throw new NotSupportedException("Combining unnamed AcroForm fields is not supported.");
        var discovered = new MergedFormRoot(
            sourceIndex,
            fieldObjectNumber,
            fullName,
            field,
            widgets,
            sourceDefaultAppearance,
            sourceQuadding,
            sourceDefaultResources);
        if (!byRoot.TryGetValue(fieldObjectNumber, out MergedFormRoot? existing)) {
            byRoot[fieldObjectNumber] = discovered;
            return;
        }

        foreach (int widgetObjectNumber in widgets) {
            if (!existing.WidgetObjectNumbers.Contains(widgetObjectNumber)) existing.WidgetObjectNumbers.Add(widgetObjectNumber);
        }
    }

    private static string GetUniqueFormFieldName(string fieldName, int sourceIndex, HashSet<string> names) {
        int sequence = 1;
        while (true) {
            string candidate = fieldName + ".source" + (sourceIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) +
                (sequence == 1 ? string.Empty : "." + sequence.ToString(System.Globalization.CultureInfo.InvariantCulture));
            if (!names.Contains(candidate)) return candidate;
            sequence++;
        }
    }

    private static void PrepareFormRoot(Dictionary<int, PdfIndirectObject> objects, MergedFormRoot root) {
        MaterializeInheritedFormAttributes(objects, root.FieldDictionary);
        if (!root.FieldDictionary.Items.ContainsKey("DA") && !string.IsNullOrEmpty(root.SourceDefaultAppearance)) {
            root.FieldDictionary.Items["DA"] = new PdfStringObj(root.SourceDefaultAppearance!, true);
        }
        if (!root.FieldDictionary.Items.ContainsKey("Q") && root.SourceQuadding.HasValue) {
            root.FieldDictionary.Items["Q"] = new PdfNumber(root.SourceQuadding.Value);
        }
        if (!root.FieldDictionary.Items.ContainsKey("DR") && root.SourceDefaultResources != null) {
            root.FieldDictionary.Items["DR"] = root.SourceDefaultResources;
        }
        root.FieldDictionary.Items.Remove("Parent");
        root.FieldDictionary.Items["T"] = new PdfStringObj(root.FieldName, true);
    }

    private static PdfObject? CloneMappedAcroFormDefaultResources(ImportedSource source, IReadOnlyDictionary<int, int> numberMap) {
        if (!source.Document.Security.RootObjectNumber.HasValue ||
            !source.Objects.TryGetValue(source.Document.Security.RootObjectNumber.Value, out PdfIndirectObject? catalogObject) ||
            catalogObject.Value is not PdfDictionary catalog ||
            !catalog.Items.TryGetValue("AcroForm", out PdfObject? acroFormObject) ||
            ResolveDictionary(source.Objects, acroFormObject) is not PdfDictionary acroForm ||
            !acroForm.Items.TryGetValue("DR", out PdfObject? defaultResources)) {
            return null;
        }

        return CloneMappedFormObject(defaultResources, numberMap);
    }

    private static PdfObject CloneMappedFormObject(PdfObject value, IReadOnlyDictionary<int, int> numberMap) {
        if (value is PdfReference reference) {
            if (!numberMap.TryGetValue(reference.ObjectNumber, out int mapped)) {
                throw new InvalidOperationException("PDF merge form resource mapping was incomplete.");
            }
            return new PdfReference(mapped, 0);
        }
        if (value is PdfArray array) {
            var clone = new PdfArray();
            foreach (PdfObject item in array.Items) clone.Items.Add(CloneMappedFormObject(item, numberMap));
            return clone;
        }
        if (value is PdfDictionary dictionary) {
            var clone = new PdfDictionary();
            foreach (KeyValuePair<string, PdfObject> item in dictionary.Items) clone.Items[item.Key] = CloneMappedFormObject(item.Value, numberMap);
            return clone;
        }
        if (value is PdfStream stream) {
            return new PdfStream((PdfDictionary)CloneMappedFormObject(stream.Dictionary, numberMap), (byte[])stream.Data.Clone(), stream.DecodingFailed, stream.DecodingError);
        }
        if (value is PdfStringObj text) return new PdfStringObj(text.RawBytes, text.UseTextStringEncoding);
        if (value is PdfName name) return new PdfName(name.Name);
        if (value is PdfNumber number) return new PdfNumber(number.Value);
        if (value is PdfBoolean boolean) return new PdfBoolean(boolean.Value);
        return PdfNull.Instance;
    }

    private static void MaterializeInheritedFormAttributes(Dictionary<int, PdfIndirectObject> objects, PdfDictionary field) {
        string[] inheritedKeys = { "FT", "Ff", "V", "DV", "DA", "Q", "Opt", "MaxLen", "AA" };
        PdfDictionary current = field;
        var visited = new HashSet<int>();
        while (current.Items.TryGetValue("Parent", out PdfObject? parentObject) &&
            parentObject is PdfReference parentReference &&
            visited.Add(parentReference.ObjectNumber) &&
            ResolveDictionary(objects, parentReference) is PdfDictionary parent) {
            for (int keyIndex = 0; keyIndex < inheritedKeys.Length; keyIndex++) {
                string key = inheritedKeys[keyIndex];
                if (!field.Items.ContainsKey(key) && parent.Items.TryGetValue(key, out PdfObject? value)) field.Items[key] = value;
            }

            current = parent;
        }
    }

    private static void RemoveDroppedWidgets(Dictionary<int, PdfIndirectObject> objects, PdfReadDocument document, HashSet<int> droppedWidgets) {
        if (droppedWidgets.Count == 0) return;
        foreach (PdfReadPage page in document.Pages) {
            if (!objects.TryGetValue(page.ObjectNumber, out PdfIndirectObject? pageObject) || pageObject.Value is not PdfDictionary pageDictionary ||
                !pageDictionary.Items.TryGetValue("Annots", out PdfObject? annotationsObject) || ResolveObject(objects, annotationsObject) is not PdfArray annotations) continue;
            var kept = new PdfArray();
            foreach (PdfObject annotation in annotations.Items) {
                if (annotation is PdfReference reference && droppedWidgets.Contains(reference.ObjectNumber)) continue;
                kept.Items.Add(annotation);
            }
            if (kept.Items.Count == 0) pageDictionary.Items.Remove("Annots"); else pageDictionary.Items["Annots"] = kept;
        }
    }

    private static PdfDictionary BuildMergedAcroForm(IReadOnlyList<MergedFormRoot> roots, PdfReadDocument primary) {
        var fields = new PdfArray();
        foreach (MergedFormRoot root in roots) fields.Items.Add(new PdfReference(root.RootObjectNumber, 0));
        var acroForm = new PdfDictionary(); acroForm.Items["Fields"] = fields; acroForm.Items["NeedAppearances"] = new PdfBoolean(false);
        if (!string.IsNullOrEmpty(primary.AcroFormDefaultAppearance)) acroForm.Items["DA"] = new PdfStringObj(primary.AcroFormDefaultAppearance!, true);
        if (primary.AcroFormQuadding.HasValue) acroForm.Items["Q"] = new PdfNumber(primary.AcroFormQuadding.Value);
        return acroForm;
    }

    private static void ValidateFormReadback(byte[] output, IReadOnlyList<string> expectedNames) {
        string[] actual = PdfReadDocument.Open(output).FormFields.Select(static field => field.Name ?? string.Empty).OrderBy(static name => name, StringComparer.Ordinal).ToArray();
        if (!actual.SequenceEqual(expectedNames, StringComparer.Ordinal)) throw new InvalidOperationException("PDF AcroForm merge validation failed; the artifact was not returned.");
    }

    private sealed class MergedFormRoot {
        internal MergedFormRoot(int sourceIndex, int rootObjectNumber, string fieldName, PdfDictionary fieldDictionary, List<int> widgetObjectNumbers, string? sourceDefaultAppearance, int? sourceQuadding, PdfObject? sourceDefaultResources) { SourceIndex = sourceIndex; RootObjectNumber = rootObjectNumber; FieldName = fieldName; FieldDictionary = fieldDictionary; WidgetObjectNumbers = widgetObjectNumbers; SourceDefaultAppearance = sourceDefaultAppearance; SourceQuadding = sourceQuadding; SourceDefaultResources = sourceDefaultResources; }
        internal int SourceIndex { get; } internal int RootObjectNumber { get; } internal string FieldName { get; set; }
        internal PdfDictionary FieldDictionary { get; } internal List<int> WidgetObjectNumbers { get; }
        internal string? SourceDefaultAppearance { get; } internal int? SourceQuadding { get; } internal PdfObject? SourceDefaultResources { get; }
    }
}
