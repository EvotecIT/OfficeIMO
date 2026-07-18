using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    private static bool IsSimpleCatalogDictionary(PdfDictionary dictionary) {
        foreach (var value in dictionary.Items.Values) {
            if (!IsSimpleCatalogValue(value)) {
                return false;
            }
        }
    
        return true;
    }
    
    private static bool IsSimpleCatalogValue(PdfObject value) {
        switch (value) {
            case PdfNumber:
            case PdfBoolean:
            case PdfName:
            case PdfStringObj:
            case PdfNull:
                return true;
            case PdfArray array:
                foreach (var item in array.Items) {
                    if (!IsSimpleCatalogValue(item)) {
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
        HashSet<int> visitedReferences) {
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
                    IsSupportedCatalogMetadataGraph(sourceObjects, indirect.Value, visitedReferences);
            case PdfArray array:
                foreach (var item in array.Items) {
                    if (!IsSupportedCatalogMetadataGraph(sourceObjects, item, visitedReferences)) {
                        return false;
                    }
                }
    
                return true;
            case PdfDictionary dictionary:
                if (IsPageDictionary(dictionary)) {
                    return false;
                }
    
                foreach (var item in dictionary.Items.Values) {
                    if (!IsSupportedCatalogMetadataGraph(sourceObjects, item, visitedReferences)) {
                        return false;
                    }
                }
    
                return true;
            case PdfStream stream:
                if (IsPageDictionary(stream.Dictionary)) {
                    return false;
                }
    
                foreach (var item in stream.Dictionary.Items.Values) {
                    if (!IsSupportedCatalogMetadataGraph(sourceObjects, item, visitedReferences)) {
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
        HashSet<int> visitedReferences) {
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
                    IsSupportedOutlineGraph(sourceObjects, indirect.Value, visitedReferences);
            case PdfArray array:
                foreach (var item in array.Items) {
                    if (!IsSupportedOutlineGraph(sourceObjects, item, visitedReferences)) {
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
                    !IsSupportedOutlineAction(sourceObjects, action)) {
                    return false;
                }
    
                foreach (var item in dictionary.Items.Values) {
                    if (!IsSupportedOutlineGraph(sourceObjects, item, visitedReferences)) {
                        return false;
                    }
                }
    
                return true;
            default:
                return false;
        }
    }
    
    private static bool IsSupportedOutlineAction(Dictionary<int, PdfIndirectObject> sourceObjects, PdfObject action) {
        return ResolveDictionary(sourceObjects, action) is PdfDictionary dictionary &&
            dictionary.Items.Count == 2 &&
            dictionary.Get<PdfName>("S")?.Name == "GoTo" &&
            dictionary.Items.TryGetValue("D", out var destination) &&
            IsDestinationForKnownPage(sourceObjects, destination);
    }
    
    private static bool OutlineDestinationsReferenceOnlyCopiedPages(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject value,
        HashSet<int> copiedPageObjectIds,
        HashSet<int> visitedReferences) {
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
    
                return OutlineDestinationsReferenceOnlyCopiedPages(sourceObjects, indirect.Value, copiedPageObjectIds, visitedReferences);
            case PdfArray array:
                foreach (var item in array.Items) {
                    if (!OutlineDestinationsReferenceOnlyCopiedPages(sourceObjects, item, copiedPageObjectIds, visitedReferences)) {
                        return false;
                    }
                }
    
                return true;
            case PdfDictionary dictionary:
                if (IsPageDictionary(dictionary)) {
                    return false;
                }
    
                foreach (var item in dictionary.Items.Values) {
                    if (!OutlineDestinationsReferenceOnlyCopiedPages(sourceObjects, item, copiedPageObjectIds, visitedReferences)) {
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
        PdfObject? viewerPreferences) {
        PdfDictionary? sourceDictionary = ResolveDictionary(sourceObjects, viewerPreferences);
        if (sourceDictionary is null) {
            return null;
        }
    
        var result = new PdfDictionary();
        foreach (var entry in sourceDictionary.Items) {
            if (!TryCloneSimpleCatalogValue(entry.Value, out var cloned)) {
                return null;
            }
    
            result.Items[entry.Key] = cloned;
        }
    
        return result;
    }
    
    private static bool TryCloneSimpleCatalogValue(PdfObject value, out PdfObject cloned) {
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
                    if (!TryCloneSimpleCatalogValue(item, out var clonedItem)) {
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
        HashSet<int> copiedPageObjectIds) {
        PdfObject? destination = ResolveObject(sourceObjects, openAction);
        if (destination is PdfArray array && IsDestinationForCopiedPages(array, copiedPageObjectIds)) {
            return array;
        }
    
        if (destination is PdfDictionary dictionary &&
            dictionary.Items.Count == 2 &&
            dictionary.Get<PdfName>("S")?.Name == "GoTo" &&
            dictionary.Items.TryGetValue("D", out var actionDestination) &&
            IsDestinationForCopiedPages(actionDestination, copiedPageObjectIds)) {
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
        HashSet<int> copiedPageObjectIds) {
        PdfDictionary? sourceDictionary = ResolveDictionary(sourceObjects, namedDestinations);
        if (sourceDictionary is null) {
            return null;
        }
    
        var result = new PdfDictionary();
        foreach (var entry in sourceDictionary.Items) {
            PdfObject? destination = ResolveObject(sourceObjects, entry.Value);
            if (destination is null) {
                continue;
            }
    
            if (IsDestinationForCopiedPages(destination, copiedPageObjectIds)) {
                result.Items[entry.Key] = destination;
            }
        }
    
        return result.Items.Count == 0 ? null : result;
    }
    
    private static bool IsDestinationForCopiedPages(PdfObject destination, HashSet<int> copiedPageObjectIds) {
        if (destination is PdfArray array) {
            return array.Items.Count > 0 &&
                array.Items[0] is PdfReference pageReference &&
                copiedPageObjectIds.Contains(pageReference.ObjectNumber) &&
                ReferencesOnlyCopiedPages(array, copiedPageObjectIds);
        }
    
        if (destination is PdfDictionary dictionary &&
            dictionary.Items.TryGetValue("D", out var explicitDestination)) {
            return IsDestinationForCopiedPages(explicitDestination, copiedPageObjectIds) &&
                ReferencesOnlyCopiedPages(dictionary, copiedPageObjectIds);
        }
    
        return false;
    }
    
    private static bool IsDestinationForKnownPage(Dictionary<int, PdfIndirectObject> sourceObjects, PdfObject destination) {
        var visitedReferences = new HashSet<int>();
        while (true) {
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
    
    private static bool ReferencesOnlyCopiedPages(PdfObject value, HashSet<int> copiedPageObjectIds) {
        switch (value) {
            case PdfReference reference:
                return copiedPageObjectIds.Contains(reference.ObjectNumber);
            case PdfArray array:
                foreach (var item in array.Items) {
                    if (!ReferencesOnlyCopiedPages(item, copiedPageObjectIds)) {
                        return false;
                    }
                }
    
                return true;
            case PdfDictionary dictionary:
                foreach (var item in dictionary.Items.Values) {
                    if (!ReferencesOnlyCopiedPages(item, copiedPageObjectIds)) {
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
