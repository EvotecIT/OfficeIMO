namespace OfficeIMO.Pdf;

internal static partial class PdfRedactionApplier {
    private static bool RemoveMatchedAnnotations(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary pageDictionary,
        IReadOnlyList<PdfRedactionMatch> matches,
        HashSet<int> formFieldObjectNumbers, System.Threading.CancellationToken cancellationToken) {
        PdfRedactionMatch[] annotationMatches = matches
            .Where(match => match.Kind == PdfRedactionMatchKind.Annotation)
            .ToArray();
        if (annotationMatches.Length == 0 ||
            !pageDictionary.Items.TryGetValue("Annots", out PdfObject? annotsObject) ||
            PdfObjectLookup.Resolve(objects, annotsObject) is not PdfArray annotations) {
            return false;
        }

        var annotationObjectNumbers = new HashSet<int>(annotationMatches
            .Where(match => match.ObjectNumber.HasValue)
            .Select(match => match.ObjectNumber!.Value));
        var popupObjectNumbers = new HashSet<int>();
        var removedObjects = new HashSet<int>();
        bool changed = false;
        for (int i = annotations.Items.Count - 1; i >= 0; i--) {
            PdfObject item = annotations.Items[i];
            PdfReference? reference = item as PdfReference;
            PdfDictionary? annotation = PdfObjectLookup.Resolve(objects, item) as PdfDictionary;
            if (annotation is null) {
                continue;
            }

            bool removeAnnotation =
                (reference is not null && annotationObjectNumbers.Contains(reference.ObjectNumber)) ||
                MatchesAnnotationRedaction(objects, annotation, annotationMatches);
            if (!removeAnnotation) {
                continue;
            }

            AddPopupObjectNumber(annotation, popupObjectNumbers);
            if (reference is not null) {
                annotationObjectNumbers.Add(reference.ObjectNumber);
                bool remainsOwnedByFormField =
                    string.Equals(TryReadName(objects, annotation, "Subtype"), "Widget", StringComparison.OrdinalIgnoreCase) &&
                    (formFieldObjectNumbers.Contains(reference.ObjectNumber) ||
                     annotation.Items.TryGetValue("Parent", out PdfObject? parentObject) &&
                     PdfObjectLookup.Resolve(objects, parentObject) is PdfDictionary);
                if (!remainsOwnedByFormField) {
                    removedObjects.Add(reference.ObjectNumber);
                }
            }

            annotations.Items.RemoveAt(i);
            changed = true;
        }

        RemovePopupReferences(objects, annotations, annotationObjectNumbers, popupObjectNumbers, removedObjects, ref changed);

        if (!changed) {
            return false;
        }

        // Read StructParent before deleting its annotation, then repair only known ownership edges.
        PdfStructureTreeAnnotationPruner.RemoveAnnotationReferences(objects, removedObjects, cancellationToken);
        RemoveDeletedAnnotationLinks(objects, removedObjects, cancellationToken);
        foreach (int number in removedObjects) objects.Remove(number);

        if (annotations.Items.Count == 0) {
            pageDictionary.Items.Remove("Annots");
        }

        return true;
    }

    internal static bool TryReadFieldRectangle(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget, out double x, out double y, out double width, out double height) =>
        TryReadRect(objects, widget, out x, out y, out width, out height);

    private static bool MatchesAnnotationRedaction(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary annotation,
        PdfRedactionMatch[] annotationMatches) {
        if (!TryReadRect(objects, annotation, out double x, out double y, out double width, out double height)) {
            return false;
        }

        string? subtype = TryReadName(objects, annotation, "Subtype");
        for (int i = 0; i < annotationMatches.Length; i++) {
            PdfRedactionMatch match = annotationMatches[i];
            if (!string.IsNullOrEmpty(match.Subtype) &&
                !string.Equals(match.Subtype, subtype, StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            if (AreClose(match.X, x) &&
                AreClose(match.Y, y) &&
                AreClose(match.Width, width) &&
                AreClose(match.Height, height)) {
                return true;
            }
        }

        return false;
    }

    private static void AddPopupObjectNumber(PdfDictionary annotation, HashSet<int> popupObjectNumbers) {
        if (annotation.Items.TryGetValue("Popup", out PdfObject? popupObject) &&
            popupObject is PdfReference popupReference) {
            popupObjectNumbers.Add(popupReference.ObjectNumber);
        }
    }

    private static void RemovePopupReferences(
        Dictionary<int, PdfIndirectObject> objects,
        PdfArray annotations,
        HashSet<int> removedAnnotationObjectNumbers,
        HashSet<int> popupObjectNumbers,
        HashSet<int> removedObjects,
        ref bool changed) {
        for (int i = annotations.Items.Count - 1; i >= 0; i--) {
            PdfObject item = annotations.Items[i];
            PdfReference? reference = item as PdfReference;
            PdfDictionary? annotation = PdfObjectLookup.Resolve(objects, item) as PdfDictionary;
            if (annotation is null) {
                continue;
            }

            bool removePopup =
                (reference is not null && popupObjectNumbers.Contains(reference.ObjectNumber)) ||
                IsPopupForRemovedAnnotation(annotation, removedAnnotationObjectNumbers);
            if (!removePopup) {
                continue;
            }

            if (reference is not null) {
                removedObjects.Add(reference.ObjectNumber);
            }

            annotations.Items.RemoveAt(i);
            changed = true;
        }

        for (int i = 0; i < annotations.Items.Count; i++) {
            if (PdfObjectLookup.Resolve(objects, annotations.Items[i]) is not PdfDictionary annotation ||
                !annotation.Items.TryGetValue("Popup", out PdfObject? popupObject) ||
                popupObject is not PdfReference popupReference ||
                (!removedAnnotationObjectNumbers.Contains(popupReference.ObjectNumber) &&
                !popupObjectNumbers.Contains(popupReference.ObjectNumber))) {
                continue;
            }

            annotation.Items.Remove("Popup");
            changed = true;
        }
    }

    private static bool IsPopupForRemovedAnnotation(PdfDictionary annotation, HashSet<int> removedAnnotationObjectNumbers) {
        if (!string.Equals(annotation.Get<PdfName>("Subtype")?.Name, "Popup", StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        return annotation.Items.TryGetValue("Parent", out PdfObject? parentObject) &&
            parentObject is PdfReference parentReference &&
            removedAnnotationObjectNumbers.Contains(parentReference.ObjectNumber);
    }

    /// <summary>Returns dictionaries established as annotations by their type or page ownership.</summary>
    internal static HashSet<PdfDictionary> GetAnnotationOwners(Dictionary<int, PdfIndirectObject> objects) {
        var owners = new HashSet<PdfDictionary>();
        foreach (PdfIndirectObject item in objects.Values) {
            if (item.Value is not PdfDictionary dictionary) continue;
            if (dictionary.Get<PdfName>("Type")?.Name == "Annot") owners.Add(dictionary);
            if (dictionary.Get<PdfName>("Type")?.Name != "Page" || !dictionary.Items.TryGetValue("Annots", out PdfObject? value) ||
                PdfObjectLookup.Resolve(objects, value) is not PdfArray annotations) continue;
            foreach (PdfObject annotation in annotations.Items)
                if (PdfObjectLookup.Resolve(objects, annotation) is PdfDictionary owner) owners.Add(owner);
        }
        return owners;
    }

    /// <summary>Detaches known annotation relationships; unexplained references remain for the rewrite proof to refuse.</summary>
    private static void RemoveDeletedAnnotationLinks(Dictionary<int, PdfIndirectObject> objects, HashSet<int> removed,
        System.Threading.CancellationToken cancellationToken) {
        foreach (PdfIndirectObject item in objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (removed.Contains(item.ObjectNumber) || item.Value is not PdfDictionary dictionary) continue;
            if (dictionary.Get<PdfName>("Type")?.Name == "Page" && dictionary.Items.TryGetValue("Annots", out PdfObject? value) &&
                PdfObjectLookup.Resolve(objects, value) is PdfArray annotations) {
                annotations.Items.RemoveAll(value => value is PdfReference reference && removed.Contains(reference.ObjectNumber));
                if (annotations.Items.Count == 0) dictionary.Items.Remove("Annots");
            }
        }
        foreach (PdfDictionary dictionary in GetAnnotationOwners(objects)) {
            if (dictionary.Items.TryGetValue("IRT", out PdfObject? reply) && reply is PdfReference target && removed.Contains(target.ObjectNumber)) {
                dictionary.Items.Remove("IRT");
                dictionary.Items.Remove("RT");
            }
            if (dictionary.Items.TryGetValue("Popup", out PdfObject? popup) && popup is PdfReference popupTarget && removed.Contains(popupTarget.ObjectNumber))
                dictionary.Items.Remove("Popup");
        }
    }
}
