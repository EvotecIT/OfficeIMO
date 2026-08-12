namespace OfficeIMO.Pdf;

internal static class PdfProvenanceGraphEditor {
    internal static byte[] RemoveFileSpecifications(
        byte[] pdf,
        HashSet<int> fileSpecificationObjectNumbers,
        PdfReadOptions? readOptions) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(fileSpecificationObjectNumbers, nameof(fileSpecificationObjectNumbers));
        if (fileSpecificationObjectNumbers.Count == 0) return (byte[])pdf.Clone();
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyAttachments, readOptions);
        return PdfDocumentObjectGraphRewriter.Rewrite(pdf, readOptions, null, (objects, security) => {
            RemoveFromEmbeddedFilesNameTree(objects, security, fileSpecificationObjectNumbers);
            HashSet<PdfDictionary> removedDirectAnnotations = CollectFileAttachmentAnnotations(
                objects, fileSpecificationObjectNumbers, out HashSet<int> removedAnnotationObjectNumbers);
            var removedObjectNumbers = new HashSet<int>(fileSpecificationObjectNumbers);
            removedObjectNumbers.UnionWith(removedAnnotationObjectNumbers);
            var visited = new HashSet<PdfObject>();
            foreach (PdfIndirectObject item in objects.Values.ToArray()) {
                ScrubReferences(item.Value, removedObjectNumbers, removedDirectAnnotations, visited);
            }
            foreach (int objectNumber in removedObjectNumbers) objects.Remove(objectNumber);
            return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value)
                ? security.InfoObjectNumber
                : null;
        });
    }

    private static void RemoveFromEmbeddedFilesNameTree(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        HashSet<int> targets) {
        if (!security.RootObjectNumber.HasValue ||
            !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? root) ||
            root.Value is not PdfDictionary catalog ||
            PdfObjectLookup.Resolve(objects, catalog.Items.TryGetValue("Names", out PdfObject? namesValue) ? namesValue : null) is not PdfDictionary names ||
            !names.Items.TryGetValue("EmbeddedFiles", out PdfObject? embeddedFiles)) return;
        RemoveNameTreePairs(objects, embeddedFiles, targets, new HashSet<PdfObject>());
    }

    private static void RemoveNameTreePairs(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        HashSet<int> targets,
        HashSet<PdfObject> visited) {
        PdfObject? resolved = PdfObjectLookup.Resolve(objects, value);
        if (resolved == null || !visited.Add(resolved) || resolved is not PdfDictionary dictionary) return;
        if (PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Names", out PdfObject? namesValue) ? namesValue : null) is PdfArray names) {
            for (int index = names.Items.Count - 2; index >= 0; index -= 2) {
                if (index + 1 < names.Items.Count && names.Items[index + 1] is PdfReference reference && targets.Contains(reference.ObjectNumber)) {
                    names.Items.RemoveAt(index + 1);
                    names.Items.RemoveAt(index);
                }
            }
            if (names.Items.Count == 0) dictionary.Items.Remove("Names");
        }
        if (PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Kids", out PdfObject? kidsValue) ? kidsValue : null) is not PdfArray kids) return;
        foreach (PdfObject child in kids.Items) RemoveNameTreePairs(objects, child, targets, visited);
    }

    private static HashSet<PdfDictionary> CollectFileAttachmentAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<int> targets,
        out HashSet<int> indirectObjectNumbers) {
        var annotations = new HashSet<PdfDictionary>();
        indirectObjectNumbers = new HashSet<int>();
        var visited = new HashSet<PdfObject>();
        foreach (PdfIndirectObject item in objects.Values) {
            CollectFileAttachmentAnnotations(item.Value, targets, annotations, visited);
            if (item.Value is PdfDictionary dictionary && annotations.Contains(dictionary)) indirectObjectNumbers.Add(item.ObjectNumber);
        }
        return annotations;
    }

    private static void CollectFileAttachmentAnnotations(
        PdfObject value,
        HashSet<int> targets,
        HashSet<PdfDictionary> annotations,
        HashSet<PdfObject> visited) {
        if (!visited.Add(value)) return;
        PdfDictionary? dictionary = value is PdfStream stream ? stream.Dictionary : value as PdfDictionary;
        if (dictionary != null) {
            if (string.Equals(dictionary.Get<PdfName>("Subtype")?.Name, "FileAttachment", StringComparison.Ordinal) &&
                dictionary.Items.TryGetValue("FS", out PdfObject? fileSpecification) &&
                fileSpecification is PdfReference reference && targets.Contains(reference.ObjectNumber)) {
                annotations.Add(dictionary);
            }
            foreach (PdfObject child in dictionary.Items.Values) {
                if (child is not PdfReference) CollectFileAttachmentAnnotations(child, targets, annotations, visited);
            }
            return;
        }
        if (value is PdfArray array) {
            foreach (PdfObject child in array.Items) {
                if (child is not PdfReference) CollectFileAttachmentAnnotations(child, targets, annotations, visited);
            }
        }
    }

    private static void ScrubReferences(
        PdfObject value,
        HashSet<int> removedObjectNumbers,
        HashSet<PdfDictionary> removedDirectAnnotations,
        HashSet<PdfObject> visited) {
        if (!visited.Add(value)) return;
        PdfDictionary? dictionary = value is PdfStream stream ? stream.Dictionary : value as PdfDictionary;
        if (dictionary != null) {
            foreach (string key in dictionary.Items.Keys.ToArray()) {
                PdfObject child = dictionary.Items[key];
                if (IsRemoved(child, removedObjectNumbers, removedDirectAnnotations)) {
                    dictionary.Items.Remove(key);
                    continue;
                }
                if (child is not PdfReference) ScrubReferences(child, removedObjectNumbers, removedDirectAnnotations, visited);
                if (key == "AF" && child is PdfArray array && array.Items.Count == 0) {
                    dictionary.Items.Remove(key);
                }
            }
            return;
        }
        if (value is not PdfArray values) return;
        for (int index = values.Items.Count - 1; index >= 0; index--) {
            PdfObject child = values.Items[index];
            if (IsRemoved(child, removedObjectNumbers, removedDirectAnnotations)) values.Items.RemoveAt(index);
            else if (child is not PdfReference) ScrubReferences(child, removedObjectNumbers, removedDirectAnnotations, visited);
        }
    }

    private static bool IsRemoved(
        PdfObject value,
        HashSet<int> removedObjectNumbers,
        HashSet<PdfDictionary> removedDirectAnnotations) =>
        value is PdfReference reference && removedObjectNumbers.Contains(reference.ObjectNumber) ||
        value is PdfDictionary dictionary && removedDirectAnnotations.Contains(dictionary);
}
