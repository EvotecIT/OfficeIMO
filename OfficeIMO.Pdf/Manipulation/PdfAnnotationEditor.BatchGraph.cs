namespace OfficeIMO.Pdf;

internal static partial class PdfAnnotationEditor {
    private static int CopyBatchAnnotations(Dictionary<int, PdfIndirectObject> objects, PdfAnnotation[] selected,
        Dictionary<int, PdfPagePoint> offsets, HashSet<int> changed) {
        var copies = new Dictionary<int, int>();
        foreach (PdfAnnotation annotation in selected) {
            int original = annotation.ObjectNumber!.Value;
            int number = NextAnnotationObjectNumber(objects);
            var clone = CopyAnnotationDictionary((PdfDictionary)objects[original].Value);
            clone.Items["NM"] = new PdfStringObj(Guid.NewGuid().ToString("N"));
            // A copy is a new review item, not a reply or a recorded decision on the original.
            clone.Items.Remove("IRT"); clone.Items.Remove("RT"); clone.Items.Remove("State"); clone.Items.Remove("StateModel");
            clone.Items.Remove("StructParent"); clone.Items.Remove("Popup");
            objects[number] = new PdfIndirectObject(number, 0, clone);
            copies[original] = number;
            var offset = offsets[annotation.PageNumber!.Value];
            foreach (int generated in ApplyUpdates(objects, clone, CreateBatchMoveOptions(objects, annotation, offset.X, offset.Y))) changed.Add(generated);
            changed.Add(number);
        }
        List<int> pages = GetPageObjectNumbersInDocumentOrder(objects);
        int addedCount = 0;
        foreach (PdfAnnotation annotation in selected) {
            var offset = offsets[annotation.PageNumber!.Value];
            int original = annotation.ObjectNumber!.Value, number = copies[original];
            var clone = (PdfDictionary)objects[number].Value;
            if (annotation.Review is { IsGroup: true, InReplyToObjectNumber: int primary } && copies.TryGetValue(primary, out int copiedPrimary)) {
                clone.Items["IRT"] = new PdfReference(copiedPrimary, 0); clone.Items["RT"] = new PdfName("Group");
            }
            int pageNumber = pages[annotation.PageNumber!.Value - 1];
            var page = (PdfDictionary)objects[pageNumber].Value;
            clone.Items["P"] = new PdfReference(pageNumber, objects[pageNumber].Generation);
            changed.Add(AddAnnotationReference(objects, pageNumber, page, new PdfReference(number, 0)));
            addedCount++;
            var source = (PdfDictionary)objects[original].Value;
            if (source.Items.TryGetValue("Popup", out PdfObject? popupValue) && PdfObjectLookup.Resolve(objects, popupValue) is PdfDictionary popup) {
                int popupNumber = NextAnnotationObjectNumber(objects);
                var popupClone = CopyAnnotationDictionary(popup);
                popupClone.Items["NM"] = new PdfStringObj(Guid.NewGuid().ToString("N"));
                popupClone.Items["Parent"] = new PdfReference(number, 0);
                popupClone.Items["P"] = new PdfReference(pageNumber, objects[pageNumber].Generation);
                if (popup.Items.TryGetValue("Rect", out PdfObject? rectangle) && PdfObjectLookup.Resolve(objects, rectangle) is PdfArray array && array.Items.Count == 4 &&
                    array.Items.All(item => PdfObjectLookup.Resolve(objects, item) is PdfNumber)) {
                    var coordinates = array.Items.Select(item => ((PdfNumber)PdfObjectLookup.Resolve(objects, item)!).Value).ToArray();
                    popupClone.Items["Rect"] = CreateNumberArray(new[] { coordinates[0] + offset.X, coordinates[1] + offset.Y, coordinates[2] + offset.X, coordinates[3] + offset.Y });
                }
                objects[popupNumber] = new PdfIndirectObject(popupNumber, 0, popupClone);
                clone.Items["Popup"] = new PdfReference(popupNumber, 0);
                changed.Add(popupNumber); changed.Add(AddAnnotationReference(objects, pageNumber, page, new PdfReference(popupNumber, 0)));
                addedCount++;
            }
        }
        return addedCount;
    }

    private static PdfDictionary CopyAnnotationDictionary(PdfDictionary source) {
        var copy = new PdfDictionary();
        foreach (var entry in source.Items) copy.Items[entry.Key] = entry.Value;
        return copy;
    }

    private static void ArrangeBatchAnnotations(Dictionary<int, PdfIndirectObject> objects, PdfAnnotation[] selected,
        PdfAnnotationOrderChange change, HashSet<int> changed) {
        var numbers = new HashSet<int>(selected.Select(annotation => annotation.ObjectNumber!.Value));
        // Popups travel with their owner in the painting list.
        foreach (int number in numbers.ToArray()) {
            var dictionary = (PdfDictionary)objects[number].Value;
            if (dictionary.Items.TryGetValue("Popup", out PdfObject? popup) && popup is PdfReference reference) numbers.Add(reference.ObjectNumber);
        }
        foreach (var page in GetBatchAnnotationPages(objects)) {
            PdfObject[] before = page.Annotations.Items.ToArray();
            bool IsSelected(PdfObject item) => item is PdfReference reference && numbers.Contains(reference.ObjectNumber);
            if (change is PdfAnnotationOrderChange.BringToFront or PdfAnnotationOrderChange.SendToBack) {
                var chosen = before.Where(IsSelected); var others = before.Where(item => !IsSelected(item));
                page.Annotations.Items.Clear();
                page.Annotations.Items.AddRange(change == PdfAnnotationOrderChange.BringToFront ? others.Concat(chosen) : chosen.Concat(others));
            } else if (change == PdfAnnotationOrderChange.Raise) {
                for (int index = page.Annotations.Items.Count - 2; index >= 0; index--) {
                    if (IsSelected(page.Annotations.Items[index]) && !IsSelected(page.Annotations.Items[index + 1]))
                        (page.Annotations.Items[index], page.Annotations.Items[index + 1]) = (page.Annotations.Items[index + 1], page.Annotations.Items[index]);
                }
            } else {
                for (int index = 1; index < page.Annotations.Items.Count; index++) {
                    if (IsSelected(page.Annotations.Items[index]) && !IsSelected(page.Annotations.Items[index - 1]))
                        (page.Annotations.Items[index], page.Annotations.Items[index - 1]) = (page.Annotations.Items[index - 1], page.Annotations.Items[index]);
                }
            }
            if (!before.SequenceEqual(page.Annotations.Items)) changed.Add(page.Owner);
        }
    }

    private static HashSet<int> RemoveBatchAnnotations(Dictionary<int, PdfIndirectObject> objects, PdfAnnotation[] selected, HashSet<int> changed) {
        var numbers = new HashSet<int>(selected.Select(annotation => annotation.ObjectNumber!.Value));
        foreach (int number in numbers.ToArray()) {
            var dictionary = (PdfDictionary)objects[number].Value;
            if (dictionary.Items.TryGetValue("Popup", out PdfObject? popup) && popup is PdfReference reference) numbers.Add(reference.ObjectNumber);
        }
        foreach (var page in GetBatchAnnotationPages(objects)) {
            bool removed = page.Annotations.Items.RemoveAll(item => item is PdfReference reference && numbers.Contains(reference.ObjectNumber)) > 0;
            if (removed) changed.Add(page.Owner);
        }
        DetachRemovedAnnotationRelationships(objects, numbers, changed);
        return numbers;
    }

    // Keep surviving comments usable when a parent or group primary is removed.
    private static void DetachRemovedAnnotationRelationships(Dictionary<int, PdfIndirectObject> objects, HashSet<int> numbers, HashSet<int> changed) {
        foreach (var page in GetBatchAnnotationPages(objects)) {
            foreach (PdfObject item in page.Annotations.Items) {
                if (PdfObjectLookup.Resolve(objects, item) is not PdfDictionary annotation ||
                    !annotation.Items.TryGetValue("IRT", out PdfObject? parent) || parent is not PdfReference parentReference || !numbers.Contains(parentReference.ObjectNumber)) continue;
                annotation.Items.Remove("IRT"); annotation.Items.Remove("RT");
                changed.Add(item is PdfReference annotationReference ? annotationReference.ObjectNumber : page.Owner);
            }
        }
    }

    private static IEnumerable<(int Owner, PdfArray Annotations)> GetBatchAnnotationPages(Dictionary<int, PdfIndirectObject> objects) {
        foreach (int number in GetPageObjectNumbersInDocumentOrder(objects)) {
            var page = (PdfDictionary)objects[number].Value;
            if (page.Items.TryGetValue("Annots", out PdfObject? value) && PdfObjectLookup.Resolve(objects, value) is PdfArray annotations)
                yield return (value is PdfReference reference ? reference.ObjectNumber : number, annotations);
        }
    }
}
