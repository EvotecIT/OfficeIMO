namespace OfficeIMO.Pdf;

public static partial class PdfMerger {
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

    private static byte[] RewriteForms(
        byte[] merged,
        IReadOnlyList<ImportedSource> sources,
        int primarySourceIndex,
        PdfMergeStructureMode mode,
        PdfMergeCollisionMode collisionMode,
        List<string> renamed,
        ref int dropped,
        out IReadOnlyList<string> expectedNames) {
        PdfReadDocument document = PdfReadDocument.Load(merged);
        int[] sourceByPage = BuildSourceByPage(sources);
        var names = new List<string>();
        int localDropped = dropped;
        byte[] output = PdfDocumentObjectGraphRewriter.Rewrite(merged, null, null, (objects, security) => {
            PdfDictionary catalog = RequireCatalog(objects, security);
            List<MergedFormRoot> roots = FindFormRoots(objects, document, sourceByPage);
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
                    RenameFormRoot(root, renamedField);
                    renamed.Add("source " + root.SourceIndex + ": " + fieldName + " -> " + renamedField);
                    fieldName = renamedField;
                    fieldNames.Add(fieldName);
                }
                root.FieldName = fieldName;
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

    private static int[] BuildSourceByPage(IReadOnlyList<ImportedSource> sources) {
        var result = new List<int>();
        for (int sourceIndex = 0; sourceIndex < sources.Count; sourceIndex++) {
            for (int page = 0; page < sources[sourceIndex].PageObjectNumbers.Length; page++) result.Add(sourceIndex);
        }
        return result.ToArray();
    }

    private static List<MergedFormRoot> FindFormRoots(Dictionary<int, PdfIndirectObject> objects, PdfReadDocument document, int[] sourceByPage) {
        var byRoot = new Dictionary<int, MergedFormRoot>();
        for (int pageIndex = 0; pageIndex < document.Pages.Count; pageIndex++) {
            PdfReadPage page = document.Pages[pageIndex];
            if (!objects.TryGetValue(page.ObjectNumber, out PdfIndirectObject? pageObject) || pageObject.Value is not PdfDictionary pageDictionary ||
                !pageDictionary.Items.TryGetValue("Annots", out PdfObject? annotationsObject) || ResolveObject(objects, annotationsObject) is not PdfArray annotations) continue;
            foreach (PdfObject annotationObject in annotations.Items) {
                if (annotationObject is not PdfReference widgetReference || ResolveDictionary(objects, widgetReference) is not PdfDictionary widget || widget.Get<PdfName>("Subtype")?.Name != "Widget") continue;
                MergedFormRoot discovered = FindFormRoot(objects, widgetReference, sourceByPage[pageIndex]);
                if (!byRoot.TryGetValue(discovered.RootObjectNumber, out MergedFormRoot? existing)) {
                    byRoot[discovered.RootObjectNumber] = discovered;
                    existing = discovered;
                } else if (!string.Equals(existing.FieldName, discovered.FieldName, StringComparison.Ordinal)) {
                    throw new NotSupportedException("Combining a hierarchical AcroForm root with multiple terminal field names is not supported yet.");
                }
                if (!existing.WidgetObjectNumbers.Contains(widgetReference.ObjectNumber)) existing.WidgetObjectNumbers.Add(widgetReference.ObjectNumber);
            }
        }
        return byRoot.Values.OrderBy(static root => root.SourceIndex).ThenBy(static root => root.RootObjectNumber).ToList();
    }

    private static MergedFormRoot FindFormRoot(Dictionary<int, PdfIndirectObject> objects, PdfReference widgetReference, int sourceIndex) {
        PdfReference current = widgetReference;
        PdfDictionary currentDictionary = ResolveDictionary(objects, current) ?? throw new InvalidOperationException("PDF widget field dictionary is not readable.");
        PdfDictionary? renameDictionary = null;
        var parts = new List<string>();
        while (true) {
            if (currentDictionary.Items.TryGetValue("T", out PdfObject? nameObject) && ResolveObject(objects, nameObject) is PdfStringObj name && !string.IsNullOrEmpty(name.Value)) {
                renameDictionary ??= currentDictionary;
                parts.Add(name.Value);
            }
            if (!currentDictionary.Items.TryGetValue("Parent", out PdfObject? parentObject) || parentObject is not PdfReference parentReference || ResolveDictionary(objects, parentReference) is not PdfDictionary parent) break;
            current = parentReference; currentDictionary = parent;
        }
        parts.Reverse();
        string fieldName = string.Join(".", parts);
        if (fieldName.Length == 0 || renameDictionary == null) throw new NotSupportedException("Combining unnamed AcroForm fields is not supported.");
        return new MergedFormRoot(sourceIndex, current.ObjectNumber, fieldName, renameDictionary, new List<int> { widgetReference.ObjectNumber });
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

    private static void RenameFormRoot(MergedFormRoot root, string fullName) {
        string oldFullName = root.FieldName;
        string oldPartial = oldFullName.Substring(oldFullName.LastIndexOf('.') + 1);
#pragma warning disable CA1845 // Span-based string.Concat is unavailable on every target framework.
        string newPartial = fullName.StartsWith(oldFullName, StringComparison.Ordinal)
            ? oldPartial + fullName.Substring(oldFullName.Length)
            : fullName;
#pragma warning restore CA1845
        root.RenameDictionary.Items["T"] = new PdfStringObj(newPartial, true);
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
        string[] actual = PdfReadDocument.Load(output).FormFields.Select(static field => field.Name ?? string.Empty).OrderBy(static name => name, StringComparer.Ordinal).ToArray();
        if (!actual.SequenceEqual(expectedNames, StringComparer.Ordinal)) throw new InvalidOperationException("PDF AcroForm merge validation failed; the artifact was not returned.");
    }

    private sealed class MergedFormRoot {
        internal MergedFormRoot(int sourceIndex, int rootObjectNumber, string fieldName, PdfDictionary renameDictionary, List<int> widgetObjectNumbers) { SourceIndex = sourceIndex; RootObjectNumber = rootObjectNumber; FieldName = fieldName; RenameDictionary = renameDictionary; WidgetObjectNumbers = widgetObjectNumbers; }
        internal int SourceIndex { get; } internal int RootObjectNumber { get; } internal string FieldName { get; set; }
        internal PdfDictionary RenameDictionary { get; } internal List<int> WidgetObjectNumbers { get; }
    }
}
