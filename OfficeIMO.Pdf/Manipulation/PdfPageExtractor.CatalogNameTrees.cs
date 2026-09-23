using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    private static PdfDictionary? BuildNamedDestinationNameTree(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? names,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!TryGetNamedDestinationNameTree(sourceObjects, names, out var namedDestinations)) {
            return null;
        }
    
        return TryBuildFlattenedNamedDestinationNameTree(sourceObjects, namedDestinations, null, out var flattenedTree,
            cancellationToken: cancellationToken)
            ? flattenedTree
            : null;
    }
    
    private static PdfDictionary? BuildNamedDestinationNameTreeForPages(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? namedDestinationNameTree,
        HashSet<int> copiedPageObjectIds,
        Dictionary<int, List<NamedDestinationNameTreeEntry>>? pageIndex = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (pageIndex is not null) {
            var candidates = new List<NamedDestinationNameTreeEntry>();
            foreach (int pageObjectId in copiedPageObjectIds) {
                cancellationToken.ThrowIfCancellationRequested();
                if (pageIndex.TryGetValue(pageObjectId, out var entries)) {
                    foreach (NamedDestinationNameTreeEntry entry in entries) {
                        cancellationToken.ThrowIfCancellationRequested();
                        candidates.Add(entry);
                    }
                }
            }

            if (candidates.Count == 0) {
                return null;
            }

            try {
                candidates.Sort((left, right) => {
                    cancellationToken.ThrowIfCancellationRequested();
                    return left.Order.CompareTo(right.Order);
                });
            } catch (InvalidOperationException error) when (error.InnerException is OperationCanceledException) {
                throw error.InnerException!;
            }
            var names = new PdfArray();
            foreach (var entry in candidates) {
                cancellationToken.ThrowIfCancellationRequested();
                PdfObject? destination = ResolveObject(sourceObjects, entry.Destination);
                if (destination is null) {
                    return null;
                }

                if (IsDestinationForCopiedPages(destination, copiedPageObjectIds, cancellationToken)) {
                    names.Items.Add(entry.Name);
                    names.Items.Add(entry.Destination);
                }
            }

            if (names.Items.Count == 0) {
                return null;
            }

            var filtered = new PdfDictionary();
            filtered.Items["Names"] = names;
            return filtered;
        }

        return TryBuildFlattenedNamedDestinationNameTree(sourceObjects, namedDestinationNameTree, copiedPageObjectIds,
            out var filteredTree, cancellationToken: cancellationToken)
            ? filteredTree
            : null;
    }

    private static Dictionary<int, List<NamedDestinationNameTreeEntry>>? BuildNamedDestinationPageIndex(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject namedDestinationNameTree,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (ResolveDictionary(sourceObjects, namedDestinationNameTree) is not PdfDictionary tree ||
            !tree.Items.TryGetValue("Names", out var namesObject) ||
            ResolveObject(sourceObjects, namesObject) is not PdfArray names ||
            names.Items.Count % 2 != 0) {
            return null;
        }

        var index = new Dictionary<int, List<NamedDestinationNameTreeEntry>>();
        for (int item = 0; item < names.Items.Count; item += 2) {
            cancellationToken.ThrowIfCancellationRequested();
            if (names.Items[item] is not PdfStringObj name ||
                !TryGetNamedDestinationPageObjectId(sourceObjects, names.Items[item + 1], out int pageObjectId)) {
                return null;
            }

            if (!index.TryGetValue(pageObjectId, out var entries)) {
                entries = new List<NamedDestinationNameTreeEntry>();
                index[pageObjectId] = entries;
            }

            entries.Add(new NamedDestinationNameTreeEntry(name, names.Items[item + 1], item / 2));
        }

        return index;
    }

    private static bool TryGetNamedDestinationPageObjectId(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject destination,
        out int pageObjectId) {
        pageObjectId = 0;
        PdfObject? current = destination;
        for (int depth = 0; depth < 32; depth++) {
            current = ResolveObject(sourceObjects, current);
            if (current is PdfDictionary dictionary && dictionary.Items.TryGetValue("D", out var nestedDestination)) {
                current = nestedDestination;
                continue;
            }

            if (current is PdfArray array && array.Items.Count > 0 && array.Items[0] is PdfReference pageReference) {
                pageObjectId = pageReference.ObjectNumber;
                return true;
            }

            return false;
        }

        return false;
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
    
    internal static bool TryBuildFlattenedNamedDestinationNameTree(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? namedDestinationNameTree,
        HashSet<int>? copiedPageObjectIds,
        out PdfDictionary result,
        int maximumNodes = PdfReadLimits.DefaultMaxNameTreeNodes,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        result = new PdfDictionary();
        var entries = new List<NamedDestinationNameTreeEntry>();
        int traversedNodes = 0;
        if (!TryCollectNamedDestinationNameTreeEntries(sourceObjects, namedDestinationNameTree, entries,
            new HashSet<int>(), 0, maximumNodes, ref traversedNodes, cancellationToken)) {
            return false;
        }
    
        var names = new PdfArray();
        foreach (var entry in entries) {
            cancellationToken.ThrowIfCancellationRequested();
            PdfObject? resolvedDestination = ResolveObject(sourceObjects, entry.Destination);
            if (resolvedDestination is null) {
                return false;
            }
    
            bool supportedDestination = copiedPageObjectIds is null
                ? IsDestinationForKnownPage(sourceObjects, resolvedDestination)
                : IsDestinationForCopiedPages(resolvedDestination, copiedPageObjectIds, cancellationToken);
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
        HashSet<int> visitedReferences,
        int depth,
        int maximumNodes,
        ref int traversedNodes,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (depth > PdfReadLimits.DefaultMaxNameTreeDepth) {
            return false;
        }

        if (value is PdfReference reference) {
            if (!visitedReferences.Add(reference.ObjectNumber) ||
                !PdfObjectLookup.TryGet(sourceObjects, reference, out var indirect)) {
                return false;
            }

            if (++traversedNodes > maximumNodes) {
                return false;
            }

            value = indirect.Value;
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
                cancellationToken.ThrowIfCancellationRequested();
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
                cancellationToken.ThrowIfCancellationRequested();
                if (kid is not PdfReference) {
                    return false;
                }
    
                if (!TryCollectNamedDestinationNameTreeEntries(sourceObjects, kid, entries, visitedReferences,
                    depth + 1, maximumNodes, ref traversedNodes, cancellationToken)) {
                    return false;
                }
            }
        }
    
        return hasNames != hasKids;
    }
    
    private static PdfObject? BuildEmbeddedFiles(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? names,
        CancellationToken cancellationToken = default) {
        if (!TryGetEmbeddedFilesNameTree(sourceObjects, names, out var embeddedFiles)) {
            return null;
        }
    
        return IsSupportedCatalogMetadataGraph(sourceObjects, embeddedFiles, new HashSet<int>(), cancellationToken)
            ? embeddedFiles
            : null;
    }
    
    private static PdfObject? BuildAssociatedFiles(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? associatedFiles,
        CancellationToken cancellationToken = default) {
        return associatedFiles is not null &&
            IsSupportedCatalogMetadataGraph(sourceObjects, associatedFiles, new HashSet<int>(), cancellationToken)
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
        PdfObject? outputIntents,
        CancellationToken cancellationToken = default) {
        return outputIntents is not null &&
            IsSupportedCatalogMetadataGraph(sourceObjects, outputIntents, new HashSet<int>(), cancellationToken)
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
        PdfObject? catalogUri,
        CancellationToken cancellationToken = default) {
        return catalogUri is not null &&
            ResolveDictionary(sourceObjects, catalogUri) is PdfDictionary dictionary &&
            IsSimpleCatalogDictionary(dictionary, cancellationToken)
            ? catalogUri
            : null;
    }
}
