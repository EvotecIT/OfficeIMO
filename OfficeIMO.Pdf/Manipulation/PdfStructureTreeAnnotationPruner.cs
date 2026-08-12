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
        var removedObjectReferences = new HashSet<int>();
        foreach (KeyValuePair<int, PdfIndirectObject> entry in objects.ToArray()) {
            if (entry.Value.Value is not PdfDictionary dictionary || !IsStructureElement(dictionary)) {
                continue;
            }

            if (RemoveObjectReferenceKids(objects, dictionary, annotations, removedObjectReferences) && !HasStructureKids(objects, dictionary)) {
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

                if (RemoveIndirectStructureKids(objects, dictionary, removedStructElements) && !HasStructureKids(objects, dictionary)) {
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

            RemoveIndirectStructureKids(objects, dictionary, removedStructElements);
            if (dictionary.Items.TryGetValue("ParentTree", out PdfObject? parentTree)) {
                if (!PruneParentTree(objects, parentTree, structParentIndexes, removedStructElements, new HashSet<int>(), out _, out _)) {
                    dictionary.Items.Remove("ParentTree");
                }
            }
            if (dictionary.Items.TryGetValue("IDTree", out PdfObject? idTree) &&
                !PruneIdTree(objects, idTree, removedStructElements, new HashSet<int>(), out _, out _)) {
                dictionary.Items.Remove("IDTree");
            }
        }

        foreach (int objectNumber in removedStructElements) {
            objects.Remove(objectNumber);
        }
        foreach (int objectNumber in removedObjectReferences) {
            objects.Remove(objectNumber);
        }
    }

    private static bool IsStructureElement(PdfDictionary dictionary) =>
        dictionary.Get<PdfName>("Type")?.Name == "StructElem";

    private static bool HasStructureKids(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary) {
        if (!dictionary.Items.TryGetValue("K", out PdfObject? kids)) {
            return false;
        }

        return PdfObjectLookup.Resolve(objects, kids) is not PdfArray array || array.Items.Count > 0;
    }

    private static bool RemoveObjectReferenceKids(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        HashSet<int> annotations,
        HashSet<int> removedObjectReferences) {
        if (!dictionary.Items.TryGetValue("K", out PdfObject? kids)) {
            return false;
        }

        if (IsObjectReferenceTo(objects, kids, annotations, removedObjectReferences)) {
            dictionary.Items.Remove("K");
            return true;
        }

        if (PdfObjectLookup.Resolve(objects, kids) is not PdfArray array) {
            return false;
        }

        bool changed = false;
        for (int i = array.Items.Count - 1; i >= 0; i--) {
            if (!IsObjectReferenceTo(objects, array.Items[i], annotations, removedObjectReferences)) {
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

    private static bool IsObjectReferenceTo(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        HashSet<int> annotations,
        HashSet<int> removedObjectReferences) {
        PdfDictionary? dictionary = value as PdfDictionary;
        if (value is PdfReference objectReference &&
            objects.TryGetValue(objectReference.ObjectNumber, out PdfIndirectObject? indirect) &&
            indirect.Value is PdfDictionary indirectDictionary) {
            dictionary = indirectDictionary;
        }

        bool matches = dictionary is not null &&
        dictionary.Get<PdfName>("Type")?.Name == "OBJR" &&
        dictionary.Items.TryGetValue("Obj", out PdfObject? objectValue) &&
        objectValue is PdfReference reference &&
        annotations.Contains(reference.ObjectNumber);
        if (matches && value is PdfReference matchedReference) {
            removedObjectReferences.Add(matchedReference.ObjectNumber);
        }
        return matches;
    }

    private static bool RemoveIndirectStructureKids(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        HashSet<int> removedStructElements) {
        if (removedStructElements.Count == 0 || !dictionary.Items.TryGetValue("K", out PdfObject? kids)) {
            return false;
        }

        if (kids is PdfReference reference && removedStructElements.Contains(reference.ObjectNumber)) {
            dictionary.Items.Remove("K");
            return true;
        }

        if (PdfObjectLookup.Resolve(objects, kids) is not PdfArray array) {
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

    private static bool PruneParentTree(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        HashSet<int> structParentIndexes,
        HashSet<int> removedStructElements,
        HashSet<int> visited,
        out PdfNumber? firstKey,
        out PdfNumber? lastKey) {
        firstKey = null;
        lastKey = null;
        if (value is PdfReference reference) {
            if (!visited.Add(reference.ObjectNumber) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) {
                return true;
            }

            return PruneParentTree(objects, indirect.Value, structParentIndexes, removedStructElements, visited, out firstKey, out lastKey);
        }

        if (value is not PdfDictionary dictionary) {
            return true;
        }

        bool hasEntries = false;
        if (dictionary.Items.TryGetValue("Nums", out PdfObject? numsObject) &&
            PdfObjectLookup.Resolve(objects, numsObject) is PdfArray nums) {
            if (nums.Items.Count % 2 != 0) return true;
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
            if (nums.Items.Count == 0) {
                dictionary.Items.Remove("Nums");
            } else if (nums.Items[0] is PdfNumber first && nums.Items[nums.Items.Count - 2] is PdfNumber last) {
                hasEntries = true;
                firstKey = first;
                lastKey = last;
            } else {
                hasEntries = true;
            }
        } else if (dictionary.Items.ContainsKey("Nums")) {
            hasEntries = true;
        }

        if (dictionary.Items.TryGetValue("Kids", out PdfObject? kidsObject)) {
            if (PdfObjectLookup.Resolve(objects, kidsObject) is PdfArray kids) {
                for (int index = kids.Items.Count - 1; index >= 0; index--) {
                    if (PruneParentTree(objects, kids.Items[index], structParentIndexes, removedStructElements, visited, out PdfNumber? childFirst, out PdfNumber? childLast)) {
                        hasEntries = true;
                        if (childFirst != null && (firstKey == null || childFirst.Value < firstKey.Value)) firstKey = childFirst;
                        if (childLast != null && (lastKey == null || childLast.Value > lastKey.Value)) lastKey = childLast;
                    } else {
                        kids.Items.RemoveAt(index);
                    }
                }
                if (kids.Items.Count == 0) dictionary.Items.Remove("Kids");
            } else {
                hasEntries = true;
            }
        }

        if (!hasEntries) {
            dictionary.Items.Remove("Limits");
            return false;
        }
        if (firstKey != null && lastKey != null) {
            var limits = new PdfArray();
            limits.Items.Add(new PdfNumber(firstKey.Value));
            limits.Items.Add(new PdfNumber(lastKey.Value));
            dictionary.Items["Limits"] = limits;
        }
        return true;
    }

    private static bool PruneIdTree(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        HashSet<int> removedStructElements,
        HashSet<int> visited,
        out PdfStringObj? firstKey,
        out PdfStringObj? lastKey) {
        firstKey = null;
        lastKey = null;
        if (value is PdfReference reference) {
            if (!visited.Add(reference.ObjectNumber)) return true;
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) return true;
            return PruneIdTree(objects, indirect.Value, removedStructElements, visited, out firstKey, out lastKey);
        }
        if (value is not PdfDictionary dictionary) return true;

        bool hasEntries = false;
        if (dictionary.Items.TryGetValue("Names", out PdfObject? namesObject)) {
            if (PdfObjectLookup.Resolve(objects, namesObject) is PdfArray names && names.Items.Count % 2 == 0) {
                for (int index = names.Items.Count - 2; index >= 0; index -= 2) {
                    if (names.Items[index + 1] is not PdfReference target ||
                        !removedStructElements.Contains(target.ObjectNumber)) {
                        continue;
                    }
                    names.Items.RemoveAt(index + 1);
                    names.Items.RemoveAt(index);
                }
                if (names.Items.Count == 0) {
                    dictionary.Items.Remove("Names");
                } else {
                    hasEntries = true;
                    firstKey = names.Items[0] as PdfStringObj;
                    lastKey = names.Items[names.Items.Count - 2] as PdfStringObj;
                }
            } else {
                // Preserve malformed or unresolved name arrays rather than deleting unrelated structure.
                hasEntries = true;
            }
        }

        if (dictionary.Items.TryGetValue("Kids", out PdfObject? kidsObject)) {
            if (PdfObjectLookup.Resolve(objects, kidsObject) is PdfArray kids) {
                var emptyKids = new List<int>();
                for (int index = 0; index < kids.Items.Count; index++) {
                    PdfObject kid = kids.Items[index];
                    if (PruneIdTree(objects, kid, removedStructElements, visited, out PdfStringObj? childFirst, out PdfStringObj? childLast)) {
                        hasEntries = true;
                        firstKey ??= childFirst;
                        if (childLast != null) lastKey = childLast;
                    } else {
                        emptyKids.Add(index);
                    }
                }
                for (int index = emptyKids.Count - 1; index >= 0; index--) kids.Items.RemoveAt(emptyKids[index]);
                if (kids.Items.Count == 0) dictionary.Items.Remove("Kids");
            } else {
                hasEntries = true;
            }
        }

        if (!hasEntries) {
            dictionary.Items.Remove("Limits");
            return false;
        }
        if (firstKey != null && lastKey != null) {
            var limits = new PdfArray();
            limits.Items.Add(firstKey);
            limits.Items.Add(lastKey);
            dictionary.Items["Limits"] = limits;
        }
        return true;
    }
}
