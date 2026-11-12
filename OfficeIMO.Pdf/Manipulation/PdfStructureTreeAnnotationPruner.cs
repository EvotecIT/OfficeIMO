namespace OfficeIMO.Pdf;

/// <summary>What <see cref="PdfStructureTreeAnnotationPruner"/> removed, so a preservation proof can expect exactly that and nothing more.</summary>
internal sealed class PdfStructureTreePruning {
    /// <summary>Structure elements removed because only references to deleted annotations were left in them.</summary>
    internal HashSet<int> RemovedStructureElements { get; } = new();

    /// <summary>Object references removed from each remaining or removed structure element, by element object number.</summary>
    internal Dictionary<int, int> RemovedObjectReferencesByElement { get; } = new();

    /// <summary>Parent tree keys removed with their entries.</summary>
    internal HashSet<int> RemovedParentTreeKeys { get; } = new();

    /// <summary>
    /// The structure the rewritten file should have: <paramref name="original"/> without the removed elements, their child links,
    /// the removed object references and the removed parent tree entries.
    /// </summary>
    internal PdfTaggedContentInfo Apply(PdfTaggedContentInfo original) {
        var elements = new List<PdfStructureElementInfo>(original.StructureElements.Count);
        foreach (PdfStructureElementInfo element in original.StructureElements) {
            if (RemovedStructureElements.Contains(element.ObjectNumber)) continue;
            RemovedObjectReferencesByElement.TryGetValue(element.ObjectNumber, out int removedReferences);
            elements.Add(new PdfStructureElementInfo(element.ObjectNumber, element.StructureType, element.ParentObjectNumber,
                element.PageObjectNumber, element.Language, element.AlternateText,
                element.ChildElementObjectNumbers.Where(child => !RemovedStructureElements.Contains(child)).ToArray(),
                element.MarkedContentReferences, element.ObjectReferenceCount - removedReferences));
        }
        return new PdfTaggedContentInfo(original.Marked, original.Suspects, original.UserProperties, original.StructTreeRootObjectNumber,
            original.ParentTreeObjectNumber, original.ParentTreeNextKey, original.RoleMap,
            original.RootElementObjectNumbers.Where(root => !RemovedStructureElements.Contains(root)).ToArray(),
            original.ParentTreeStructParentIndexes.Where(key => !RemovedParentTreeKeys.Contains(key)).ToArray(),
            elements);
    }
}

/// <summary>
/// Removes tagged-structure references that point at annotations being deleted by a full rewrite.
/// </summary>
internal static class PdfStructureTreeAnnotationPruner {
    internal static PdfStructureTreePruning RemoveAnnotationReferences(
        Dictionary<int, PdfIndirectObject> objects,
        IEnumerable<int> annotationObjectNumbers, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var pruning = new PdfStructureTreePruning();
        var annotations = new HashSet<int>(annotationObjectNumbers);
        if (annotations.Count == 0) {
            return pruning;
        }

        var structParentIndexes = new HashSet<int>();
        foreach (int annotationObjectNumber in annotations) {
            cancellationToken.ThrowIfCancellationRequested();
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
            cancellationToken.ThrowIfCancellationRequested();
            if (entry.Value.Value is not PdfDictionary dictionary || !IsStructureElement(dictionary)) {
                continue;
            }

            int removedReferences = RemoveObjectReferenceKids(objects, dictionary, annotations, removedObjectReferences, cancellationToken);
            if (removedReferences == 0) continue;
            pruning.RemovedObjectReferencesByElement[entry.Key] = removedReferences;
            if (!HasStructureKids(objects, dictionary, cancellationToken)) {
                removedStructElements.Add(entry.Key);
            }
        }

        bool changed;
        do {
            changed = false;
            foreach (KeyValuePair<int, PdfIndirectObject> entry in objects.ToArray()) {
                cancellationToken.ThrowIfCancellationRequested();
                if (removedStructElements.Contains(entry.Key) ||
                    entry.Value.Value is not PdfDictionary dictionary ||
                    !IsStructureElement(dictionary)) {
                    continue;
                }

                if (RemoveIndirectStructureKids(objects, dictionary, removedStructElements, cancellationToken) && !HasStructureKids(objects, dictionary, cancellationToken)) {
                    removedStructElements.Add(entry.Key);
                    changed = true;
                }
            }
        } while (changed);

        foreach (PdfIndirectObject indirect in objects.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (indirect.Value is not PdfDictionary dictionary ||
                dictionary.Get<PdfName>("Type")?.Name != "StructTreeRoot") {
                continue;
            }

            RemoveIndirectStructureKids(objects, dictionary, removedStructElements, cancellationToken);
            if (dictionary.Items.TryGetValue("ParentTree", out PdfObject? parentTree)) {
                if (!PruneParentTree(objects, parentTree, structParentIndexes, removedStructElements, pruning.RemovedParentTreeKeys, new HashSet<int>(), out _, out _, cancellationToken)) {
                    dictionary.Items.Remove("ParentTree");
                }
            }
            if (dictionary.Items.TryGetValue("IDTree", out PdfObject? idTree) &&
                !PruneIdTree(objects, idTree, removedStructElements, new HashSet<int>(), out _, out _, cancellationToken)) {
                dictionary.Items.Remove("IDTree");
            }
        }

        foreach (int objectNumber in removedStructElements) {
            cancellationToken.ThrowIfCancellationRequested();
            objects.Remove(objectNumber);
        }
        foreach (int objectNumber in removedObjectReferences) {
            cancellationToken.ThrowIfCancellationRequested();
            objects.Remove(objectNumber);
        }
        pruning.RemovedStructureElements.UnionWith(removedStructElements);
        return pruning;
    }

    private static bool IsStructureElement(PdfDictionary dictionary) =>
        dictionary.Get<PdfName>("Type")?.Name == "StructElem";

    private static bool HasStructureKids(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!dictionary.Items.TryGetValue("K", out PdfObject? kids)) {
            return false;
        }

        return PdfObjectLookup.Resolve(objects, kids) is not PdfArray array || array.Items.Count > 0;
    }

    /// <summary>Removes the element's object references to deleted annotations and returns how many it removed.</summary>
    private static int RemoveObjectReferenceKids(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary dictionary,
        HashSet<int> annotations,
        HashSet<int> removedObjectReferences, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!dictionary.Items.TryGetValue("K", out PdfObject? kids)) {
            return 0;
        }

        if (IsObjectReferenceTo(objects, kids, annotations, removedObjectReferences, cancellationToken)) {
            dictionary.Items.Remove("K");
            return 1;
        }

        if (PdfObjectLookup.Resolve(objects, kids) is not PdfArray array) {
            return 0;
        }

        int removed = 0;
        for (int i = array.Items.Count - 1; i >= 0; i--) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!IsObjectReferenceTo(objects, array.Items[i], annotations, removedObjectReferences, cancellationToken)) {
                continue;
            }

            array.Items.RemoveAt(i);
            removed++;
        }

        if (array.Items.Count == 0) {
            dictionary.Items.Remove("K");
        }

        return removed;
    }

    private static bool IsObjectReferenceTo(
        Dictionary<int, PdfIndirectObject> objects,
        PdfObject value,
        HashSet<int> annotations,
        HashSet<int> removedObjectReferences, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
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
        HashSet<int> removedStructElements, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
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
            cancellationToken.ThrowIfCancellationRequested();
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
        HashSet<int> removedKeys,
        HashSet<int> visited,
        out PdfNumber? firstKey,
        out PdfNumber? lastKey, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        firstKey = null;
        lastKey = null;
        if (value is PdfReference reference) {
            if (!visited.Add(reference.ObjectNumber) ||
                !PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) {
                return true;
            }

            return PruneParentTree(objects, indirect.Value, structParentIndexes, removedStructElements, removedKeys, visited, out firstKey, out lastKey, cancellationToken);
        }

        if (value is not PdfDictionary dictionary) {
            return true;
        }

        bool hasEntries = false;
        if (dictionary.Items.TryGetValue("Nums", out PdfObject? numsObject) &&
            PdfObjectLookup.Resolve(objects, numsObject) is PdfArray nums) {
            if (nums.Items.Count % 2 != 0) return true;
            for (int i = nums.Items.Count - 2; i >= 0; i -= 2) {
                cancellationToken.ThrowIfCancellationRequested();
                bool removePair = nums.Items[i] is PdfNumber key &&
                    key.Value >= 0D && key.Value <= int.MaxValue &&
                    Math.Floor(key.Value) == key.Value &&
                    structParentIndexes.Contains((int)key.Value);
                if (!removePair && nums.Items[i + 1] is PdfReference structElementReference) {
                    removePair = removedStructElements.Contains(structElementReference.ObjectNumber);
                }

                if (removePair) {
                    if (nums.Items[i] is PdfNumber removedKey && removedKey.Value >= 0D && removedKey.Value <= int.MaxValue &&
                        Math.Floor(removedKey.Value) == removedKey.Value) {
                        removedKeys.Add((int)removedKey.Value);
                    }
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
                    cancellationToken.ThrowIfCancellationRequested();
                    if (PruneParentTree(objects, kids.Items[index], structParentIndexes, removedStructElements, removedKeys, visited, out PdfNumber? childFirst, out PdfNumber? childLast, cancellationToken)) {
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
        out PdfStringObj? lastKey, System.Threading.CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        firstKey = null;
        lastKey = null;
        if (value is PdfReference reference) {
            if (!visited.Add(reference.ObjectNumber)) return true;
            if (!PdfObjectLookup.TryGet(objects, reference, out PdfIndirectObject? indirect)) return true;
            return PruneIdTree(objects, indirect.Value, removedStructElements, visited, out firstKey, out lastKey, cancellationToken);
        }
        if (value is not PdfDictionary dictionary) return true;

        bool hasEntries = false;
        if (dictionary.Items.TryGetValue("Names", out PdfObject? namesObject)) {
            if (PdfObjectLookup.Resolve(objects, namesObject) is PdfArray names && names.Items.Count % 2 == 0) {
                for (int index = names.Items.Count - 2; index >= 0; index -= 2) {
                    cancellationToken.ThrowIfCancellationRequested();
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
                    cancellationToken.ThrowIfCancellationRequested();
                    PdfObject kid = kids.Items[index];
                    if (PruneIdTree(objects, kid, removedStructElements, visited, out PdfStringObj? childFirst, out PdfStringObj? childLast, cancellationToken)) {
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
