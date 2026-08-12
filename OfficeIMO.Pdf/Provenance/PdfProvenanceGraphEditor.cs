namespace OfficeIMO.Pdf;

internal static class PdfProvenanceGraphEditor {
    internal static byte[] RemoveFileSpecifications(
        byte[] pdf,
        HashSet<int> fileSpecificationObjectNumbers,
        PdfReadOptions? readOptions,
        long maximumOutputBytes) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(fileSpecificationObjectNumbers, nameof(fileSpecificationObjectNumbers));
        if (fileSpecificationObjectNumbers.Count == 0) return (byte[])pdf.Clone();
        _ = PdfMutationPlanner.RequireFullRewrite(pdf, PdfMutationOperation.ModifyAttachments, readOptions);
        return PdfDocumentObjectGraphRewriter.Rewrite(pdf, readOptions, null, (objects, security) => {
            RemoveFromEmbeddedFilesNameTree(objects, security, fileSpecificationObjectNumbers);
            HashSet<PdfDictionary> removedDirectAnnotations = CollectFileAttachmentAnnotations(
                objects, security, fileSpecificationObjectNumbers, out HashSet<int> removedAnnotationObjectNumbers);
            var removedObjectNumbers = new HashSet<int>(fileSpecificationObjectNumbers);
            removedObjectNumbers.UnionWith(removedAnnotationObjectNumbers);
            var visited = new HashSet<PdfObject>();
            foreach (PdfIndirectObject item in objects.Values.ToArray()) {
                ScrubReferences(objects, item.Value, removedObjectNumbers, removedDirectAnnotations, visited);
            }
            visited.Clear();
            foreach (PdfIndirectObject item in objects.Values.ToArray()) {
                RemoveEmptyAssociatedFileReferences(objects, item.Value, visited);
            }
            foreach (int objectNumber in removedObjectNumbers) objects.Remove(objectNumber);
            return security.InfoObjectNumber.HasValue && objects.ContainsKey(security.InfoObjectNumber.Value)
                ? security.InfoObjectNumber
                : null;
        }, maximumOutputBytes);
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
        if (PdfObjectLookup.Resolve(objects, embeddedFiles) is not PdfDictionary rootTree) return;
        if (!PruneNameTree(objects, rootTree, targets, new HashSet<PdfObject>(), out _, out _)) {
            names.Items.Remove("EmbeddedFiles");
        }
    }

    private static bool PruneNameTree(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        HashSet<int> targets,
        HashSet<PdfObject> visited,
        out PdfObject? firstName,
        out PdfObject? lastName) {
        firstName = lastName = null;
        PdfObject? resolved = PdfObjectLookup.Resolve(objects, value);
        if (resolved == null || !visited.Add(resolved) || resolved is not PdfDictionary dictionary) return false;
        bool hadLimits = dictionary.Items.ContainsKey("Limits");
        if (PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Names", out PdfObject? namesValue) ? namesValue : null) is PdfArray names) {
            int completePairCount = names.Items.Count / 2;
            for (int pair = completePairCount - 1; pair >= 0; pair--) {
                int index = pair * 2;
                PdfObject fileSpecification = names.Items[index + 1];
                if (fileSpecification is not PdfReference reference ||
                    !targets.Contains(reference.ObjectNumber) ||
                    !PdfObjectLookup.TryGet(objects, reference, out _)) continue;
                names.Items.RemoveAt(index + 1);
                names.Items.RemoveAt(index);
            }
            if (names.Items.Count == 0) dictionary.Items.Remove("Names");
            else {
                completePairCount = names.Items.Count / 2;
                if (completePairCount == 0) {
                    dictionary.Items.Remove("Names");
                } else {
                    firstName = names.Items[0];
                    lastName = names.Items[(completePairCount - 1) * 2];
                }
            }
        }
        if (PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Kids", out PdfObject? kidsValue) ? kidsValue : null) is PdfArray kids) {
            for (int index = kids.Items.Count - 1; index >= 0; index--) {
                if (!PruneNameTree(objects, kids.Items[index], targets, visited, out PdfObject? childFirst, out PdfObject? childLast)) {
                    kids.Items.RemoveAt(index);
                    continue;
                }
                firstName = childFirst ?? firstName;
                if (lastName == null) lastName = childLast;
            }
            if (kids.Items.Count == 0) dictionary.Items.Remove("Kids");
        }
        if (firstName == null || lastName == null) {
            dictionary.Items.Remove("Limits");
            return false;
        }
        if (hadLimits) {
            var limits = new PdfArray();
            limits.Items.Add(firstName);
            limits.Items.Add(lastName);
            dictionary.Items["Limits"] = limits;
        }
        return true;
    }

    private static HashSet<PdfDictionary> CollectFileAttachmentAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        HashSet<int> targets,
        out HashSet<int> indirectObjectNumbers) {
        var annotations = new HashSet<PdfDictionary>();
        indirectObjectNumbers = new HashSet<int>();
        if (!security.RootObjectNumber.HasValue || !objects.TryGetValue(security.RootObjectNumber.Value, out PdfIndirectObject? root) ||
            root.Value is not PdfDictionary catalog || !catalog.Items.TryGetValue("Pages", out PdfObject? pages)) return annotations;
        var pending = new Stack<PdfObject>();
        var visited = new HashSet<PdfObject>();
        pending.Push(pages);
        while (pending.Count > 0) {
            PdfObject current = pending.Pop();
            PdfObject? resolved = PdfObjectLookup.Resolve(objects, current);
            if (resolved is not PdfDictionary dictionary || !visited.Add(resolved)) continue;
            string? type = dictionary.Get<PdfName>("Type")?.Name;
            if (type == "Pages") {
                if (PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Kids", out PdfObject? kidsValue) ? kidsValue : null) is PdfArray kids) {
                    foreach (PdfObject child in kids.Items) pending.Push(child);
                }
                continue;
            }
            if (type != "Page" || PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Annots", out PdfObject? annotsValue) ? annotsValue : null) is not PdfArray pageAnnotations) continue;
            foreach (PdfObject annotationValue in pageAnnotations.Items) {
                PdfObject? annotation = PdfObjectLookup.Resolve(objects, annotationValue);
                CollectFileAttachmentAnnotation(objects, annotation, targets, annotations);
                if (annotationValue is PdfReference reference && annotation is PdfDictionary && annotations.Contains((PdfDictionary)annotation)) {
                    indirectObjectNumbers.Add(reference.ObjectNumber);
                }
            }
        }
        CollectLinkedPopupAnnotations(objects, annotations, indirectObjectNumbers);
        return annotations;
    }

    private static void CollectLinkedPopupAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        HashSet<PdfDictionary> annotations,
        HashSet<int> indirectObjectNumbers) {
        foreach (PdfDictionary annotation in annotations.ToArray()) {
            if (!annotation.Items.TryGetValue("Popup", out PdfObject? popup)) continue;
            if (popup is PdfReference reference &&
                PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) &&
                indirect.Value is PdfDictionary popupDictionary) {
                annotations.Add(popupDictionary);
                indirectObjectNumbers.Add(reference.ObjectNumber);
            } else if (popup is PdfDictionary directPopup) {
                annotations.Add(directPopup);
            }
        }

        foreach (PdfIndirectObject item in objects.Values) {
            if (item.Value is not PdfDictionary dictionary ||
                !string.Equals(dictionary.Get<PdfName>("Subtype")?.Name, "Popup", StringComparison.Ordinal) ||
                !dictionary.Items.TryGetValue("Parent", out PdfObject? parent) ||
                parent is not PdfReference parentReference ||
                !PdfObjectLookup.TryGet(objects, parentReference, out PdfIndirectObject? parentIndirect) ||
                parentIndirect.Value is not PdfDictionary parentDictionary ||
                !annotations.Contains(parentDictionary)) continue;
            annotations.Add(dictionary);
            indirectObjectNumbers.Add(item.ObjectNumber);
        }
    }

    private static void CollectFileAttachmentAnnotation(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject? value,
        HashSet<int> targets,
        HashSet<PdfDictionary> annotations) {
        if (value == null) return;
        PdfDictionary? dictionary = value is PdfStream stream ? stream.Dictionary : value as PdfDictionary;
        if (dictionary != null) {
            if (string.Equals(dictionary.Get<PdfName>("Subtype")?.Name, "FileAttachment", StringComparison.Ordinal) &&
                dictionary.Items.TryGetValue("FS", out PdfObject? fileSpecification) &&
                fileSpecification is PdfReference reference && targets.Contains(reference.ObjectNumber) &&
                PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect) &&
                indirect.Value is PdfDictionary fileSpecificationDictionary &&
                (fileSpecificationDictionary.Get<PdfName>("Type")?.Name is null or "Filespec")) {
                annotations.Add(dictionary);
            }
        }
    }

    private static void ScrubReferences(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        HashSet<int> removedObjectNumbers,
        HashSet<PdfDictionary> removedDirectAnnotations,
        HashSet<PdfObject> visited) {
        if (!visited.Add(value)) return;
        PdfDictionary? dictionary = value is PdfStream stream ? stream.Dictionary : value as PdfDictionary;
        if (dictionary != null) {
            foreach (string key in dictionary.Items.Keys.ToArray()) {
                PdfObject child = dictionary.Items[key];
                if (IsRemoved(objects, child, removedObjectNumbers, removedDirectAnnotations)) {
                    dictionary.Items.Remove(key);
                    continue;
                }
                if (child is not PdfReference) ScrubReferences(objects, child, removedObjectNumbers, removedDirectAnnotations, visited);
                if (key == "AF" && child is PdfArray array && array.Items.Count == 0) {
                    dictionary.Items.Remove(key);
                }
            }
            return;
        }
        if (value is not PdfArray values) return;
        for (int index = values.Items.Count - 1; index >= 0; index--) {
            PdfObject child = values.Items[index];
            if (IsRemoved(objects, child, removedObjectNumbers, removedDirectAnnotations)) values.Items.RemoveAt(index);
            else if (child is not PdfReference) ScrubReferences(objects, child, removedObjectNumbers, removedDirectAnnotations, visited);
        }
    }

    private static bool IsRemoved(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        HashSet<int> removedObjectNumbers,
        HashSet<PdfDictionary> removedDirectAnnotations) =>
        value is PdfReference reference &&
        removedObjectNumbers.Contains(reference.ObjectNumber) &&
        PdfObjectLookup.TryGet(objects, reference, out _) ||
        value is PdfDictionary dictionary && removedDirectAnnotations.Contains(dictionary);

    private static void RemoveEmptyAssociatedFileReferences(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        HashSet<PdfObject> visited) {
        if (!visited.Add(value)) return;
        PdfDictionary? dictionary = value is PdfStream stream ? stream.Dictionary : value as PdfDictionary;
        if (dictionary != null) {
            foreach (string key in dictionary.Items.Keys.ToArray()) {
                PdfObject child = dictionary.Items[key];
                if (key == "AF" && PdfObjectLookup.Resolve(objects, child) is PdfArray associatedFiles && associatedFiles.Items.Count == 0) {
                    dictionary.Items.Remove(key);
                    continue;
                }
                if (child is not PdfReference) RemoveEmptyAssociatedFileReferences(objects, child, visited);
            }
            return;
        }
        if (value is not PdfArray array) return;
        foreach (PdfObject child in array.Items) {
            if (child is not PdfReference) RemoveEmptyAssociatedFileReferences(objects, child, visited);
        }
    }
}
