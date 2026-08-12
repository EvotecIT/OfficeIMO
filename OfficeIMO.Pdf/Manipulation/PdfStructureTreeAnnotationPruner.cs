namespace OfficeIMO.Pdf;

/// <summary>
/// Removes tagged-structure references that point at annotations being deleted by a full rewrite.
/// </summary>
internal static class PdfStructureTreeAnnotationPruner {
    internal static void RemoveAnnotationReferences(
        Dictionary<int, PdfIndirectObject> objects,
        IEnumerable<int> annotationObjectNumbers) {
        var annotations = new HashSet<int>(annotationObjectNumbers);
        if (annotations.Count == 0) {
            return;
        }

        var structParentIndexes = new HashSet<int>();
        foreach (int annotationObjectNumber in annotations) {
            if (objects.TryGetValue(annotationObjectNumber, out PdfIndirectObject? annotation) &&
                annotation.Value is PdfDictionary annotationDictionary &&
                annotationDictionary.Get<PdfNumber>("StructParent") is PdfNumber structParent &&
                structParent.Value >= 0D &&
                structParent.Value <= int.MaxValue &&
                Math.Floor(structParent.Value) == structParent.Value) {
                structParentIndexes.Add((int)structParent.Value);
            }
        }

        var removedStructElements = new HashSet<int>();
        foreach (KeyValuePair<int, PdfIndirectObject> entry in objects.ToArray()) {
            if (entry.Value.Value is not PdfDictionary dictionary || !IsStructureElement(dictionary)) {
                continue;
            }

            if (RemoveObjectReferenceKids(dictionary, annotations) && !HasStructureKids(dictionary)) {
                removedStructElements.Add(entry.Key);
            }
        }

        bool changed;
        do {
            changed = false;
            foreach (KeyValuePair<int, PdfIndirectObject> entry in objects.ToArray()) {
                if (removedStructElements.Contains(entry.Key) ||
                    entry.Value.Value is not PdfDictionary dictionary ||
                    !IsStructureElement(dictionary)) {
                    continue;
                }

                if (RemoveIndirectStructureKids(dictionary, removedStructElements) && !HasStructureKids(dictionary)) {
                    removedStructElements.Add(entry.Key);
                    changed = true;
                }
            }
        } while (changed);

        foreach (PdfIndirectObject indirect in objects.Values) {
            if (indirect.Value is not PdfDictionary dictionary ||
                dictionary.Get<PdfName>("Type")?.Name != "StructTreeRoot") {
                continue;
            }

            RemoveIndirectStructureKids(dictionary, removedStructElements);
            if (dictionary.Items.TryGetValue("ParentTree", out PdfObject? parentTree)) {
                PruneParentTree(objects, parentTree, structParentIndexes, removedStructElements, new HashSet<int>());
            }
        }

        foreach (int objectNumber in removedStructElements) {
            objects.Remove(objectNumber);
        }
    }

    private static bool IsStructureElement(PdfDictionary dictionary) =>
        dictionary.Get<PdfName>("Type")?.Name == "StructElem";

    private static bool HasStructureKids(PdfDictionary dictionary) {
        if (!dictionary.Items.TryGetValue("K", out PdfObject? kids)) {
            return false;
        }

        return kids is not PdfArray array || array.Items.Count > 0;
    }

    private static bool RemoveObjectReferenceKids(PdfDictionary dictionary, HashSet<int> annotations) {
        if (!dictionary.Items.TryGetValue("K", out PdfObject? kids)) {
            return false;
        }

        if (IsObjectReferenceTo(kids, annotations)) {
            dictionary.Items.Remove("K");
            return true;
        }

        if (kids is not PdfArray array) {
            return false;
        }

        bool changed = false;
        for (int i = array.Items.Count - 1; i >= 0; i--) {
            if (!IsObjectReferenceTo(array.Items[i], annotations)) {
                continue;
            }

            array.Items.RemoveAt(i);
            changed = true;
        }

        if (array.Items.Count == 0) {
            dictionary.Items.Remove("K");
        }

        return changed;
    }

    private static bool IsObjectReferenceTo(PdfObject value, HashSet<int> annotations) =>
        value is PdfDictionary dictionary &&
        dictionary.Get<PdfName>("Type")?.Name == "OBJR" &&
        dictionary.Items.TryGetValue("Obj", out PdfObject? objectValue) &&
        objectValue is PdfReference reference &&
        annotations.Contains(reference.ObjectNumber);

    private static bool RemoveIndirectStructureKids(PdfDictionary dictionary, HashSet<int> removedStructElements) {
        if (removedStructElements.Count == 0 || !dictionary.Items.TryGetValue("K", out PdfObject? kids)) {
            return false;
        }

        if (kids is PdfReference reference && removedStructElements.Contains(reference.ObjectNumber)) {
            dictionary.Items.Remove("K");
            return true;
        }

        if (kids is not PdfArray array) {
            return false;
        }

        bool changed = false;
        for (int i = array.Items.Count - 1; i >= 0; i--) {
            if (array.Items[i] is PdfReference childReference && removedStructElements.Contains(childReference.ObjectNumber)) {
                array.Items.RemoveAt(i);
                changed = true;
            }
        }

        if (array.Items.Count == 0) {
            dictionary.Items.Remove("K");
        }

        return changed;
    }

    private static void PruneParentTree(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        HashSet<int> structParentIndexes,
        HashSet<int> removedStructElements,
        HashSet<int> visited) {
        if (value is PdfReference reference) {
            if (!visited.Add(reference.ObjectNumber) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) {
                return;
            }

            PruneParentTree(objects, indirect.Value, structParentIndexes, removedStructElements, visited);
            return;
        }

        if (value is not PdfDictionary dictionary) {
            return;
        }

        if (dictionary.Items.TryGetValue("Nums", out PdfObject? numsObject) && numsObject is PdfArray nums) {
            for (int i = nums.Items.Count - 2; i >= 0; i -= 2) {
                bool removePair = nums.Items[i] is PdfNumber key &&
                    key.Value >= 0D && key.Value <= int.MaxValue &&
                    Math.Floor(key.Value) == key.Value &&
                    structParentIndexes.Contains((int)key.Value);
                if (!removePair && nums.Items[i + 1] is PdfReference structElementReference) {
                    removePair = removedStructElements.Contains(structElementReference.ObjectNumber);
                }

                if (removePair) {
                    nums.Items.RemoveAt(i + 1);
                    nums.Items.RemoveAt(i);
                }
            }
        }

        if (dictionary.Items.TryGetValue("Kids", out PdfObject? kidsObject) && kidsObject is PdfArray kids) {
            foreach (PdfObject kid in kids.Items) {
                PruneParentTree(objects, kid, structParentIndexes, removedStructElements, visited);
            }
        }
    }
}
