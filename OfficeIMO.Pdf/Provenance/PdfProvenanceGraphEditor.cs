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
        if (!PruneNameTree(objects, rootTree, targets)) {
            names.Items.Remove("EmbeddedFiles");
        }
    }

    private static bool PruneNameTree(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        HashSet<int> targets) {
        PdfObject? root = PdfObjectLookup.Resolve(objects, value);
        if (root is not PdfDictionary rootDictionary) return false;
        var visited = new HashSet<PdfObject>();
        var pending = new Stack<(PdfDictionary Dictionary, bool Expanded)>();
        var results = new Dictionary<PdfDictionary, NameTreeResult>();
        visited.Add(rootDictionary);
        pending.Push((rootDictionary, false));
        while (pending.Count > 0) {
            (PdfDictionary dictionary, bool expanded) = pending.Pop();
            if (!expanded) {
                pending.Push((dictionary, true));
                if (PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Kids", out PdfObject? kidsValue) ? kidsValue : null) is PdfArray kids) {
                    for (int index = kids.Items.Count - 1; index >= 0; index--) {
                        if (PdfObjectLookup.Resolve(objects, kids.Items[index]) is PdfDictionary child && visited.Add(child)) {
                            pending.Push((child, false));
                        }
                    }
                }
                continue;
            }
            results[dictionary] = PruneNameTreeDictionary(objects, dictionary, targets, results);
        }
        return results.TryGetValue(rootDictionary, out NameTreeResult result) && result.ShouldRetain;
    }

    private static NameTreeResult PruneNameTreeDictionary(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        HashSet<int> targets,
        Dictionary<PdfDictionary, NameTreeResult> results) {
        bool hadLimits = dictionary.Items.ContainsKey("Limits");
        bool changed = false;
        PdfObject? firstName = null;
        PdfObject? lastName = null;
        if (PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Names", out PdfObject? namesValue) ? namesValue : null) is PdfArray names) {
            int completePairCount = names.Items.Count / 2;
            for (int pair = completePairCount - 1; pair >= 0; pair--) {
                int index = pair * 2;
                if (!IsTargetReference(objects, names.Items[index], targets) &&
                    !IsTargetReference(objects, names.Items[index + 1], targets)) continue;
                names.Items.RemoveAt(index + 1);
                names.Items.RemoveAt(index);
                changed = true;
            }
            if (names.Items.Count % 2 != 0 && IsTargetReference(objects, names.Items[names.Items.Count - 1], targets)) {
                names.Items.RemoveAt(names.Items.Count - 1);
                changed = true;
            }
            completePairCount = names.Items.Count / 2;
            if (completePairCount == 0) {
                if (changed) dictionary.Items.Remove("Names");
            } else {
                firstName = names.Items[0];
                lastName = names.Items[(completePairCount - 1) * 2];
            }
        }
        if (PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Kids", out PdfObject? kidsValue) ? kidsValue : null) is PdfArray kids) {
            for (int index = kids.Items.Count - 1; index >= 0; index--) {
                if (PdfObjectLookup.Resolve(objects, kids.Items[index]) is PdfDictionary child &&
                    results.TryGetValue(child, out NameTreeResult childResult) &&
                    childResult.WasChanged) {
                    changed = true;
                    if (!childResult.ShouldRetain) kids.Items.RemoveAt(index);
                }
            }
            foreach (PdfObject childValue in kids.Items) {
                if (PdfObjectLookup.Resolve(objects, childValue) is not PdfDictionary child ||
                    !results.TryGetValue(child, out NameTreeResult childResult)) continue;
                firstName ??= childResult.FirstName;
                lastName = childResult.LastName ?? lastName;
            }
            if (kids.Items.Count == 0 && changed) dictionary.Items.Remove("Kids");
        }
        if (firstName == null || lastName == null) {
            bool hasUnrelatedContent = dictionary.Items.Keys.Any(key =>
                !string.Equals(key, "Limits", StringComparison.Ordinal));
            if (changed && !hasUnrelatedContent) dictionary.Items.Remove("Limits");
            return new NameTreeResult(null, null, changed, !changed || hasUnrelatedContent);
        }
        if (hadLimits && changed) {
            var limits = new PdfArray();
            limits.Items.Add(firstName);
            limits.Items.Add(lastName);
            dictionary.Items["Limits"] = limits;
        }
        return new NameTreeResult(firstName, lastName, changed, shouldRetain: true);
    }

    private static bool IsTargetReference(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        HashSet<int> targets) =>
        value is PdfReference reference &&
        targets.Contains(reference.ObjectNumber) &&
        PdfObjectLookup.TryGet(objects, reference, out _);

    private readonly struct NameTreeResult {
        internal NameTreeResult(PdfObject? firstName, PdfObject? lastName, bool wasChanged, bool shouldRetain) {
            FirstName = firstName;
            LastName = lastName;
            WasChanged = wasChanged;
            ShouldRetain = shouldRetain;
        }

        internal PdfObject? FirstName { get; }
        internal PdfObject? LastName { get; }
        internal bool WasChanged { get; }
        internal bool ShouldRetain { get; }
        internal bool HasEntries => FirstName != null && LastName != null;
    }

    private static HashSet<PdfDictionary> CollectFileAttachmentAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDocumentSecurityInfo security,
        HashSet<int> targets,
        out HashSet<int> indirectObjectNumbers) {
        var annotations = new HashSet<PdfDictionary>();
        var activeAnnotations = new List<(PdfObject Value, PdfDictionary Dictionary)>();
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
            PdfArray? kids = PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Kids", out PdfObject? kidsValue) ? kidsValue : null) as PdfArray;
            if (type == "Pages" || kids != null) {
                if (kids != null) {
                    foreach (PdfObject child in kids.Items) pending.Push(child);
                }
                continue;
            }
            if ((type != null && type != "Page") ||
                PdfObjectLookup.Resolve(objects, dictionary.Items.TryGetValue("Annots", out PdfObject? annotsValue) ? annotsValue : null) is not PdfArray pageAnnotations) continue;
            foreach (PdfObject annotationValue in pageAnnotations.Items) {
                PdfObject? annotation = PdfObjectLookup.Resolve(objects, annotationValue);
                if (annotation is PdfDictionary annotationDictionary) activeAnnotations.Add((annotationValue, annotationDictionary));
                CollectFileAttachmentAnnotation(objects, annotation, targets, annotations);
                if (annotationValue is PdfReference reference && annotation is PdfDictionary && annotations.Contains((PdfDictionary)annotation)) {
                    indirectObjectNumbers.Add(reference.ObjectNumber);
                }
            }
        }
        CollectDependentAnnotations(objects, activeAnnotations, annotations, indirectObjectNumbers);
        return annotations;
    }

    private static void CollectDependentAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyList<(PdfObject Value, PdfDictionary Dictionary)> activeAnnotations,
        HashSet<PdfDictionary> annotations,
        HashSet<int> indirectObjectNumbers) {
        var dependents = new Dictionary<PdfDictionary, List<(PdfDictionary Dictionary, int? ObjectNumber)>>();
        foreach ((PdfObject value, PdfDictionary dictionary) in activeAnnotations) {
            if (dictionary.Items.TryGetValue("IRT", out PdfObject? replyTo) &&
                PdfObjectLookup.Resolve(objects, replyTo) is PdfDictionary replyParent) {
                AddDependent(dependents, replyParent, dictionary, value is PdfReference replyReference ? replyReference.ObjectNumber : (int?)null);
            }
            if (!dictionary.Items.TryGetValue("Popup", out PdfObject? popup) ||
                PdfObjectLookup.Resolve(objects, popup) is not PdfDictionary popupDictionary ||
                !IsLinkedPopup(objects, popupDictionary, dictionary)) continue;
            AddDependent(dependents, dictionary, popupDictionary, popup is PdfReference popupReference ? popupReference.ObjectNumber : (int?)null);
        }
        foreach (PdfIndirectObject item in objects.Values) {
            if (item.Value is not PdfDictionary dictionary ||
                !string.Equals(dictionary.Get<PdfName>("Subtype")?.Name, "Popup", StringComparison.Ordinal) ||
                !dictionary.Items.TryGetValue("Parent", out PdfObject? parent) ||
                PdfObjectLookup.Resolve(objects, parent) is not PdfDictionary parentDictionary) continue;
            AddDependent(dependents, parentDictionary, dictionary, item.ObjectNumber);
        }

        var pending = new Queue<PdfDictionary>(annotations);
        while (pending.Count > 0) {
            PdfDictionary parent = pending.Dequeue();
            if (!dependents.TryGetValue(parent, out List<(PdfDictionary Dictionary, int? ObjectNumber)>? children)) continue;
            foreach ((PdfDictionary dictionary, int? objectNumber) in children) {
                if (!annotations.Add(dictionary)) continue;
                if (objectNumber.HasValue) indirectObjectNumbers.Add(objectNumber.Value);
                pending.Enqueue(dictionary);
            }
        }
    }

    private static void AddDependent(
        Dictionary<PdfDictionary, List<(PdfDictionary Dictionary, int? ObjectNumber)>> dependents,
        PdfDictionary parent,
        PdfDictionary child,
        int? objectNumber) {
        if (!dependents.TryGetValue(parent, out List<(PdfDictionary Dictionary, int? ObjectNumber)>? children)) {
            children = new List<(PdfDictionary Dictionary, int? ObjectNumber)>();
            dependents[parent] = children;
        }
        children.Add((child, objectNumber));
    }

    private static bool IsLinkedPopup(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary popup,
        PdfDictionary expectedParent) =>
        string.Equals(popup.Get<PdfName>("Subtype")?.Name, "Popup", StringComparison.Ordinal) &&
        popup.Items.TryGetValue("Parent", out PdfObject? parent) &&
        ReferenceEquals(PdfObjectLookup.Resolve(objects, parent), expectedParent);

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
                if ((key == "Names" || key == "Nums") && PdfObjectLookup.Resolve(objects, child) is PdfArray treePairs) {
                    RemoveTreePairs(objects, treePairs, removedObjectNumbers, removedDirectAnnotations);
                    ScrubReferences(objects, treePairs, removedObjectNumbers, removedDirectAnnotations, visited);
                    if (treePairs.Items.Count == 0) dictionary.Items.Remove(key);
                    continue;
                }
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

    private static void RemoveTreePairs(
        Dictionary<int, PdfIndirectObject> objects,
        PdfArray names,
        HashSet<int> removedObjectNumbers,
        HashSet<PdfDictionary> removedDirectAnnotations) {
        int completePairCount = names.Items.Count / 2;
        for (int pair = completePairCount - 1; pair >= 0; pair--) {
            int index = pair * 2;
            if (!IsRemoved(objects, names.Items[index], removedObjectNumbers, removedDirectAnnotations) &&
                !IsRemoved(objects, names.Items[index + 1], removedObjectNumbers, removedDirectAnnotations)) continue;
            names.Items.RemoveAt(index + 1);
            names.Items.RemoveAt(index);
        }
        if (names.Items.Count % 2 != 0 &&
            IsRemoved(objects, names.Items[names.Items.Count - 1], removedObjectNumbers, removedDirectAnnotations)) {
            names.Items.RemoveAt(names.Items.Count - 1);
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
