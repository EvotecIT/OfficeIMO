using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    private static bool IsSimpleCatalogDictionary(PdfDictionary dictionary,
        CancellationToken cancellationToken = default) {
        foreach (var value in dictionary.Items.Values) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!IsSimpleCatalogValue(value, cancellationToken)) {
                return false;
            }
        }
    
        return true;
    }
    
    private static bool IsSimpleCatalogValue(PdfObject value,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (value) {
            case PdfNumber:
            case PdfBoolean:
            case PdfName:
            case PdfStringObj:
            case PdfNull:
                return true;
            case PdfArray array:
                foreach (var item in array.Items) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!IsSimpleCatalogValue(item, cancellationToken)) {
                        return false;
                    }
                }
    
                return true;
            default:
                return false;
        }
    }
    
    private static bool IsXmpMetadataStream(PdfStream stream) {
        return stream.Dictionary.Get<PdfName>("Type")?.Name == "Metadata" &&
            stream.Dictionary.Get<PdfName>("Subtype")?.Name == "XML";
    }
    
    private static bool IsSupportedCatalogMetadataGraph(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject value,
        HashSet<int> visitedReferences,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (value) {
            case PdfNumber:
            case PdfBoolean:
            case PdfName:
            case PdfStringObj:
            case PdfNull:
                return true;
            case PdfReference reference:
                if (!visitedReferences.Add(reference.ObjectNumber)) {
                    return true;
                }
    
                if (!PdfObjectLookup.TryGet(sourceObjects, reference, out var indirect)) {
                    return false;
                }
    
                return !IsPageDictionary(indirect.Value) &&
                    IsSupportedCatalogMetadataGraph(sourceObjects, indirect.Value, visitedReferences, cancellationToken);
            case PdfArray array:
                foreach (var item in array.Items) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!IsSupportedCatalogMetadataGraph(sourceObjects, item, visitedReferences, cancellationToken)) {
                        return false;
                    }
                }
    
                return true;
            case PdfDictionary dictionary:
                if (IsPageDictionary(dictionary)) {
                    return false;
                }
    
                foreach (var item in dictionary.Items.Values) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!IsSupportedCatalogMetadataGraph(sourceObjects, item, visitedReferences, cancellationToken)) {
                        return false;
                    }
                }
    
                return true;
            case PdfStream stream:
                if (IsPageDictionary(stream.Dictionary)) {
                    return false;
                }
    
                foreach (var item in stream.Dictionary.Items.Values) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!IsSupportedCatalogMetadataGraph(sourceObjects, item, visitedReferences, cancellationToken)) {
                        return false;
                    }
                }
    
                return true;
            default:
                return false;
        }
    }
    
    private static bool IsSupportedOutlineGraph(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject value,
        HashSet<int> visitedReferences,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (value) {
            case PdfNumber:
            case PdfBoolean:
            case PdfName:
            case PdfStringObj:
            case PdfNull:
                return true;
            case PdfReference reference:
                if (!visitedReferences.Add(reference.ObjectNumber)) {
                    return true;
                }
    
                if (!PdfObjectLookup.TryGet(sourceObjects, reference, out var indirect)) {
                    return false;
                }
    
                return IsPageDictionary(indirect.Value) ||
                    IsSupportedOutlineGraph(sourceObjects, indirect.Value, visitedReferences, cancellationToken);
            case PdfArray array:
                foreach (var item in array.Items) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!IsSupportedOutlineGraph(sourceObjects, item, visitedReferences, cancellationToken)) {
                        return false;
                    }
                }
    
                return true;
            case PdfDictionary dictionary:
                if (IsPageDictionary(dictionary)) {
                    return true;
                }
    
                if (dictionary.Items.ContainsKey("AA")) {
                    return false;
                }
    
                if (dictionary.Items.TryGetValue("A", out var action) &&
                    !IsSupportedOutlineAction(sourceObjects, action, cancellationToken)) {
                    return false;
                }
    
                foreach (var item in dictionary.Items.Values) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!IsSupportedOutlineGraph(sourceObjects, item, visitedReferences, cancellationToken)) {
                        return false;
                    }
                }
    
                return true;
            default:
                return false;
        }
    }
    
    private static bool IsSupportedOutlineAction(Dictionary<int, PdfIndirectObject> sourceObjects, PdfObject action,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return ResolveDictionary(sourceObjects, action) is PdfDictionary dictionary &&
            dictionary.Items.Count == 2 &&
            dictionary.Get<PdfName>("S")?.Name == "GoTo" &&
            dictionary.Items.TryGetValue("D", out var destination) &&
            IsDestinationForKnownPage(sourceObjects, destination, cancellationToken);
    }
    
    private static bool OutlineDestinationsReferenceOnlyCopiedPages(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject value,
        HashSet<int> copiedPageObjectIds,
        HashSet<int> visitedReferences,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (value) {
            case PdfNumber:
            case PdfBoolean:
            case PdfName:
            case PdfStringObj:
            case PdfNull:
                return true;
            case PdfReference reference:
                if (!PdfObjectLookup.TryGet(sourceObjects, reference, out var indirect)) {
                    return false;
                }
    
                if (IsPageDictionary(indirect.Value)) {
                    return copiedPageObjectIds.Contains(reference.ObjectNumber);
                }
    
                if (!visitedReferences.Add(reference.ObjectNumber)) {
                    return true;
                }
    
                return OutlineDestinationsReferenceOnlyCopiedPages(sourceObjects, indirect.Value, copiedPageObjectIds, visitedReferences, cancellationToken);
            case PdfArray array:
                foreach (var item in array.Items) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!OutlineDestinationsReferenceOnlyCopiedPages(sourceObjects, item, copiedPageObjectIds, visitedReferences, cancellationToken)) {
                        return false;
                    }
                }
    
                return true;
            case PdfDictionary dictionary:
                if (IsPageDictionary(dictionary)) {
                    return false;
                }
    
                foreach (var item in dictionary.Items.Values) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!OutlineDestinationsReferenceOnlyCopiedPages(sourceObjects, item, copiedPageObjectIds, visitedReferences, cancellationToken)) {
                        return false;
                    }
                }
    
                return true;
            default:
                return false;
        }
    }
    
    private static bool IsPageDictionary(PdfObject value) {
        return value is PdfDictionary dictionary &&
            dictionary.Get<PdfName>("Type")?.Name == "Page";
    }
    
    private static PdfDictionary? BuildViewerPreferences(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? viewerPreferences,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfDictionary? sourceDictionary = ResolveDictionary(sourceObjects, viewerPreferences);
        if (sourceDictionary is null) {
            return null;
        }
    
        var result = new PdfDictionary();
        foreach (var entry in sourceDictionary.Items) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryCloneSimpleCatalogValue(entry.Value, out var cloned, cancellationToken)) {
                return null;
            }
    
            result.Items[entry.Key] = cloned;
        }
    
        return result;
    }
    
    private static bool TryCloneSimpleCatalogValue(PdfObject value, out PdfObject cloned,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (value) {
            case PdfNumber number:
                cloned = new PdfNumber(number.Value);
                return true;
            case PdfBoolean boolean:
                cloned = new PdfBoolean(boolean.Value);
                return true;
            case PdfName name:
                cloned = new PdfName(name.Name);
                return true;
            case PdfStringObj text:
                cloned = new PdfStringObj(text.Value);
                return true;
            case PdfNull:
                cloned = PdfNull.Instance;
                return true;
            case PdfArray array:
                var clonedArray = new PdfArray();
                foreach (var item in array.Items) {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (!TryCloneSimpleCatalogValue(item, out var clonedItem, cancellationToken)) {
                        cloned = PdfNull.Instance;
                        return false;
                    }
    
                    clonedArray.Items.Add(clonedItem);
                }
    
                cloned = clonedArray;
                return true;
            default:
                cloned = PdfNull.Instance;
                return false;
        }
    }
    
    private static PdfObject? BuildOpenActionForPages(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? openAction,
        HashSet<int> copiedPageObjectIds,
        CancellationToken cancellationToken) {
        PdfObject? destination = ResolveObject(sourceObjects, openAction);
        if (destination is PdfArray array && IsDestinationForCopiedPages(array, copiedPageObjectIds, cancellationToken)) {
            return array;
        }
    
        if (destination is PdfDictionary dictionary &&
            dictionary.Items.Count == 2 &&
            dictionary.Get<PdfName>("S")?.Name == "GoTo" &&
            dictionary.Items.TryGetValue("D", out var actionDestination) &&
            IsDestinationForCopiedPages(actionDestination, copiedPageObjectIds, cancellationToken)) {
            var result = new PdfDictionary();
            result.Items["S"] = new PdfName("GoTo");
            result.Items["D"] = actionDestination;
            return result;
        }
    
        return null;
    }
    
    private static PdfDictionary? BuildNamedDestinationsForPages(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? namedDestinations,
        HashSet<int> copiedPageObjectIds,
        Dictionary<int, List<DirectNamedDestinationEntry>>? pageIndex = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (pageIndex is not null) {
            if (copiedPageObjectIds.Count == 1) {
                var pageEnumerator = copiedPageObjectIds.GetEnumerator();
                pageEnumerator.MoveNext();
                return pageIndex.TryGetValue(pageEnumerator.Current, out var pageEntries)
                    ? BuildIndexedNamedDestinations(sourceObjects, copiedPageObjectIds, pageEntries, cancellationToken)
                    : null;
            }

            var candidates = new List<DirectNamedDestinationEntry>();
            foreach (int pageObjectId in copiedPageObjectIds) {
                cancellationToken.ThrowIfCancellationRequested();
                if (pageIndex.TryGetValue(pageObjectId, out var entries)) {
                    foreach (DirectNamedDestinationEntry entry in entries) {
                        cancellationToken.ThrowIfCancellationRequested();
                        candidates.Add(entry);
                    }
                }
            }

            if (candidates.Count == 0) return null;
            try {
                candidates.Sort((left, right) => {
                    cancellationToken.ThrowIfCancellationRequested();
                    return left.Order.CompareTo(right.Order);
                });
            } catch (InvalidOperationException error) when (error.InnerException is OperationCanceledException) {
                throw error.InnerException!;
            }
            return BuildIndexedNamedDestinations(sourceObjects, copiedPageObjectIds, candidates, cancellationToken);
        }

        PdfDictionary? sourceDictionary = ResolveDictionary(sourceObjects, namedDestinations);
        if (sourceDictionary is null) {
            return null;
        }
    
        var result = new PdfDictionary();
        foreach (var entry in sourceDictionary.Items) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject? destination = ResolveObject(sourceObjects, entry.Value);
            if (destination is null) {
                continue;
            }
    
            if (IsDestinationForCopiedPages(destination, copiedPageObjectIds, cancellationToken)) {
                result.Items[entry.Key] = destination;
            }
        }
    
        return result.Items.Count == 0 ? null : result;
    }

    private static PdfDictionary? BuildIndexedNamedDestinations(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        HashSet<int> copiedPageObjectIds,
        List<DirectNamedDestinationEntry> entries,
        CancellationToken cancellationToken = default) {
        var result = new PdfDictionary();
        for (int index = 0; index < entries.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            DirectNamedDestinationEntry entry = entries[index];
            PdfObject? destination = ResolveObject(sourceObjects, entry.Destination);
            if (destination is null) return null;
            if (IsDestinationForCopiedPages(destination, copiedPageObjectIds, cancellationToken)) {
                result.Items[entry.Name] = destination;
            }
        }

        return result.Items.Count == 0 ? null : result;
    }

    private static Dictionary<int, List<DirectNamedDestinationEntry>>? BuildDirectNamedDestinationPageIndex(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject namedDestinations,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfDictionary? sourceDictionary = ResolveDictionary(sourceObjects, namedDestinations);
        if (sourceDictionary is null) return null;

        var index = new Dictionary<int, List<DirectNamedDestinationEntry>>();
        int order = 0;
        foreach (var entry in sourceDictionary.Items) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!TryGetNamedDestinationPageObjectId(sourceObjects, entry.Value, out int pageObjectId)) {
                return null;
            }

            if (!index.TryGetValue(pageObjectId, out var entries)) {
                entries = new List<DirectNamedDestinationEntry>();
                index[pageObjectId] = entries;
            }
            entries.Add(new DirectNamedDestinationEntry(entry.Key, entry.Value, order++));
        }

        return index;
    }
    
    private static bool IsDestinationForCopiedPages(PdfObject destination, HashSet<int> copiedPageObjectIds,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (destination is PdfArray array) {
            return array.Items.Count > 0 &&
                array.Items[0] is PdfReference pageReference &&
                copiedPageObjectIds.Contains(pageReference.ObjectNumber) &&
                ReferencesOnlyCopiedPages(array, copiedPageObjectIds, cancellationToken);
        }
    
        if (destination is PdfDictionary dictionary &&
            dictionary.Items.TryGetValue("D", out var explicitDestination)) {
            return IsDestinationForCopiedPages(explicitDestination, copiedPageObjectIds, cancellationToken) &&
                ReferencesOnlyCopiedPages(dictionary, copiedPageObjectIds, cancellationToken);
        }
    
        return false;
    }
    
    private static bool IsDestinationForKnownPage(Dictionary<int, PdfIndirectObject> sourceObjects, PdfObject destination,
        CancellationToken cancellationToken) {
        var visitedReferences = new HashSet<int>();
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            if (destination is PdfReference reference) {
                if (!visitedReferences.Add(reference.ObjectNumber) ||
                    !PdfObjectLookup.TryGet(sourceObjects, reference, out var indirect)) {
                    return false;
                }
    
                destination = indirect.Value;
                continue;
            }
    
            if (destination is PdfDictionary dictionary &&
                dictionary.Items.TryGetValue("D", out var explicitDestination)) {
                destination = explicitDestination;
                continue;
            }
    
            return destination is PdfArray array &&
                array.Items.Count > 0 &&
                array.Items[0] is PdfReference pageReference &&
                PdfObjectLookup.TryGet(sourceObjects, pageReference, out var pageObject) &&
                IsPageDictionary(pageObject.Value);
        }
    }
    
    private static bool ReferencesOnlyCopiedPages(PdfObject value, HashSet<int> copiedPageObjectIds,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        switch (value) {
            case PdfReference reference:
                return copiedPageObjectIds.Contains(reference.ObjectNumber);
            case PdfArray array:
                foreach (var item in array.Items) {
                    if (!ReferencesOnlyCopiedPages(item, copiedPageObjectIds, cancellationToken)) {
                        return false;
                    }
                }
    
                return true;
            case PdfDictionary dictionary:
                foreach (var item in dictionary.Items.Values) {
                    if (!ReferencesOnlyCopiedPages(item, copiedPageObjectIds, cancellationToken)) {
                        return false;
                    }
                }
    
                return true;
            default:
                return true;
        }
    }
    
    private static PdfObject? ResolveObject(Dictionary<int, PdfIndirectObject> sourceObjects, PdfObject? value) {
        return PdfObjectLookup.Resolve(sourceObjects, value);
    }
    
    private static PdfDictionary? ResolveDictionary(Dictionary<int, PdfIndirectObject> sourceObjects, PdfObject? value) {
        return ResolveObject(sourceObjects, value) as PdfDictionary;
    }
}
