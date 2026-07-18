using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    private static PdfDictionary? BuildNamedDestinationNameTree(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? names) {
        if (!TryGetNamedDestinationNameTree(sourceObjects, names, out var namedDestinations)) {
            return null;
        }
    
        return TryBuildFlattenedNamedDestinationNameTree(sourceObjects, namedDestinations, null, out var flattenedTree)
            ? flattenedTree
            : null;
    }
    
    private static PdfDictionary? BuildNamedDestinationNameTreeForPages(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? namedDestinationNameTree,
        HashSet<int> copiedPageObjectIds) {
        return TryBuildFlattenedNamedDestinationNameTree(sourceObjects, namedDestinationNameTree, copiedPageObjectIds, out var filteredTree)
            ? filteredTree
            : null;
    }
    
    private static bool TryGetNamedDestinationNameTree(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? names,
        out PdfObject namedDestinations) {
        namedDestinations = PdfNull.Instance;
        PdfDictionary? namesDictionary = ResolveDictionary(sourceObjects, names);
        if (namesDictionary is null ||
            !namesDictionary.Items.TryGetValue("Dests", out var namedDestinationTree)) {
            return false;
        }
    
        namedDestinations = namedDestinationTree;
        return true;
    }
    
    private static bool IsSupportedNamedDestinationNameTree(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject namedDestinations) {
        return TryBuildFlattenedNamedDestinationNameTree(sourceObjects, namedDestinations, null, out _);
    }
    
    private static bool TryBuildFlattenedNamedDestinationNameTree(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? namedDestinationNameTree,
        HashSet<int>? copiedPageObjectIds,
        out PdfDictionary result) {
        result = new PdfDictionary();
        var entries = new List<NamedDestinationNameTreeEntry>();
        if (!TryCollectNamedDestinationNameTreeEntries(sourceObjects, namedDestinationNameTree, entries, new HashSet<int>())) {
            return false;
        }
    
        var names = new PdfArray();
        foreach (var entry in entries) {
            PdfObject? resolvedDestination = ResolveObject(sourceObjects, entry.Destination);
            if (resolvedDestination is null) {
                return false;
            }
    
            bool supportedDestination = copiedPageObjectIds is null
                ? IsDestinationForKnownPage(sourceObjects, resolvedDestination)
                : IsDestinationForCopiedPages(resolvedDestination, copiedPageObjectIds);
            if (!supportedDestination) {
                if (copiedPageObjectIds is null) {
                    return false;
                }
    
                continue;
            }
    
            names.Items.Add(entry.Name);
            names.Items.Add(entry.Destination);
        }
    
        if (names.Items.Count == 0) {
            return false;
        }
    
        result.Items["Names"] = names;
        return true;
    }
    
    private static bool TryCollectNamedDestinationNameTreeEntries(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? value,
        List<NamedDestinationNameTreeEntry> entries,
        HashSet<int> visitedReferences) {
        if (value is PdfReference reference) {
            if (!visitedReferences.Add(reference.ObjectNumber) ||
                !PdfObjectLookup.TryGet(sourceObjects, reference, out var indirect)) {
                return false;
            }
    
            return TryCollectNamedDestinationNameTreeEntries(sourceObjects, indirect.Value, entries, visitedReferences);
        }
    
        if (value is not PdfDictionary tree) {
            return false;
        }
    
        bool hasNames = false;
        if (tree.Items.TryGetValue("Names", out var namesObject)) {
            hasNames = true;
            if (ResolveObject(sourceObjects, namesObject) is not PdfArray names ||
                names.Items.Count % 2 != 0) {
                return false;
            }
    
            for (int i = 0; i < names.Items.Count; i += 2) {
                if (names.Items[i] is not PdfStringObj name) {
                    return false;
                }
    
                entries.Add(new NamedDestinationNameTreeEntry(name, names.Items[i + 1]));
            }
        }
    
        bool hasKids = false;
        if (tree.Items.TryGetValue("Kids", out var kidsObject)) {
            hasKids = true;
            if (ResolveObject(sourceObjects, kidsObject) is not PdfArray kids) {
                return false;
            }
    
            foreach (var kid in kids.Items) {
                if (kid is not PdfReference) {
                    return false;
                }
    
                if (!TryCollectNamedDestinationNameTreeEntries(sourceObjects, kid, entries, visitedReferences)) {
                    return false;
                }
            }
        }
    
        return hasNames != hasKids;
    }
    
    private static PdfObject? BuildEmbeddedFiles(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? names) {
        if (!TryGetEmbeddedFilesNameTree(sourceObjects, names, out var embeddedFiles)) {
            return null;
        }
    
        return IsSupportedCatalogMetadataGraph(sourceObjects, embeddedFiles, new HashSet<int>())
            ? embeddedFiles
            : null;
    }
    
    private static PdfObject? BuildAssociatedFiles(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? associatedFiles) {
        return associatedFiles is not null &&
            IsSupportedCatalogMetadataGraph(sourceObjects, associatedFiles, new HashSet<int>())
            ? associatedFiles
            : null;
    }
    
    private static bool TryGetEmbeddedFilesNameTree(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? names,
        out PdfObject embeddedFiles) {
        embeddedFiles = PdfNull.Instance;
        PdfDictionary? namesDictionary = ResolveDictionary(sourceObjects, names);
        if (namesDictionary is null ||
            !namesDictionary.Items.TryGetValue("EmbeddedFiles", out var embeddedFileTree)) {
            return false;
        }
    
        embeddedFiles = embeddedFileTree;
        return true;
    }
    
    private static PdfObject? BuildOutputIntents(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? outputIntents) {
        return outputIntents is not null &&
            IsSupportedCatalogMetadataGraph(sourceObjects, outputIntents, new HashSet<int>())
            ? outputIntents
            : null;
    }
    
    private static PdfReference? BuildXmpMetadata(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? xmpMetadata) {
        if (xmpMetadata is not PdfReference reference ||
            !PdfObjectLookup.TryGet(sourceObjects, reference, out var indirect) ||
            indirect.Value is not PdfStream stream ||
            !IsXmpMetadataStream(stream)) {
            return null;
        }
    
        return reference;
    }
    
    private static PdfObject? BuildCatalogUri(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? catalogUri) {
        return catalogUri is not null &&
            ResolveDictionary(sourceObjects, catalogUri) is PdfDictionary dictionary &&
            IsSimpleCatalogDictionary(dictionary)
            ? catalogUri
            : null;
    }
}
