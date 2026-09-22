using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    internal static CatalogRewriteState ExtractCatalogRewriteState(Dictionary<int, PdfIndirectObject> sourceObjects,
        string? trailerRaw = null, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfDictionary? dictionary = PdfSyntax.FindCatalog(sourceObjects, trailerRaw);
        if (dictionary is not null) {
            string? pageMode = dictionary.Get<PdfName>("PageMode")?.Name;
            string? pageLayout = dictionary.Get<PdfName>("PageLayout")?.Name;
            dictionary.Items.TryGetValue("Version", out var catalogVersion);
            dictionary.Items.TryGetValue("Lang", out var catalogLanguage);
            dictionary.Items.TryGetValue("PageLabels", out var pageLabels);
            dictionary.Items.TryGetValue("Dests", out var namedDestinations);
            dictionary.Items.TryGetValue("OpenAction", out var openAction);
            dictionary.Items.TryGetValue("Outlines", out var outlines);
            dictionary.Items.TryGetValue("ViewerPreferences", out var viewerPreferences);
            dictionary.Items.TryGetValue("Metadata", out var xmpMetadata);
            dictionary.Items.TryGetValue("URI", out var catalogUri);
            dictionary.Items.TryGetValue("OutputIntents", out var outputIntents);
            dictionary.Items.TryGetValue("Names", out var names);
            dictionary.Items.TryGetValue("AF", out var associatedFiles);
            dictionary.Items.TryGetValue("OCProperties", out var optionalContent);
            return new CatalogRewriteState(pageMode, pageLayout, BuildCatalogVersion(sourceObjects, catalogVersion), BuildCatalogLanguage(sourceObjects, catalogLanguage), BuildOutlines(sourceObjects, outlines, cancellationToken), pageLabels, namedDestinations, BuildNamedDestinationNameTree(sourceObjects, names, cancellationToken), openAction, BuildViewerPreferences(sourceObjects, viewerPreferences), BuildXmpMetadata(sourceObjects, xmpMetadata), BuildCatalogUri(sourceObjects, catalogUri), BuildOutputIntents(sourceObjects, outputIntents, cancellationToken), BuildEmbeddedFiles(sourceObjects, names, cancellationToken), BuildAssociatedFiles(sourceObjects, associatedFiles, cancellationToken), BuildOptionalContent(sourceObjects, optionalContent, cancellationToken), GetPageObjectNumbersInDocumentOrder(sourceObjects, dictionary, cancellationToken), sourceObjects, cancellationToken);
        }
    
        return CatalogRewriteState.Empty;
    }
    
    internal static CatalogRewriteState PruneCatalogStateForPages(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        CatalogRewriteState catalogState,
        HashSet<int> copiedPageObjectIds,
        IReadOnlyList<int>? orderedPageObjectNumbers = null,
        int outputPageIndexOffset = 0,
        IReadOnlyDictionary<int, int>? outputPageIndexByPageObjectNumber = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var namedDestinations = BuildNamedDestinationsForPages(sourceObjects, catalogState.NamedDestinations, copiedPageObjectIds, catalogState.GetDirectNamedDestinationPageIndexForRepeatedUse(), cancellationToken);
        var namedDestinationNameTree = BuildNamedDestinationNameTreeForPages(sourceObjects, catalogState.NamedDestinationNameTree, copiedPageObjectIds, catalogState.GetNamedDestinationPageIndexForRepeatedUse(), cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        var openAction = BuildOpenActionForPages(sourceObjects, catalogState.OpenAction, copiedPageObjectIds);
        var outlines = BuildOutlinesForPages(sourceObjects, catalogState.Outlines, copiedPageObjectIds, cancellationToken);
        var pageLabels = BuildPageLabelsForPages(sourceObjects, catalogState.PageLabels, orderedPageObjectNumbers, outputPageIndexOffset, outputPageIndexByPageObjectNumber, catalogState.SourcePageIndexes, catalogState.PageLabelEntries, cancellationToken);
        string? pageMode = outlines is null && string.Equals(catalogState.PageMode, "UseOutlines", StringComparison.Ordinal)
            ? null
            : catalogState.PageMode;
        return new CatalogRewriteState(pageMode, catalogState.PageLayout, catalogState.CatalogVersion, catalogState.CatalogLanguage, outlines, pageLabels, namedDestinations, namedDestinationNameTree, openAction, catalogState.ViewerPreferences, catalogState.XmpMetadata, catalogState.CatalogUri, catalogState.OutputIntents, catalogState.EmbeddedFiles, catalogState.AssociatedFiles, catalogState.OptionalContent);
    }
    
    internal static Dictionary<int, Dictionary<string, PdfObject>>? BuildPageOverridesWithFilteredDestinationLinks(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        IReadOnlyList<int> pageObjectNumbers,
        Dictionary<int, Dictionary<string, PdfObject>>? pageOverrides,
        CatalogRewriteState catalogState,
        HashSet<int> copiedPageObjectIds,
        CancellationToken cancellationToken = default) {
        var availableDestinationNames = GetNamedDestinationNames(sourceObjects, catalogState, cancellationToken);
        Dictionary<int, Dictionary<string, PdfObject>>? result = null;
    
        if (pageOverrides is not null && pageOverrides.Count > 0) {
            result = new Dictionary<int, Dictionary<string, PdfObject>>();
            foreach (var pageEntry in pageOverrides) {
                cancellationToken.ThrowIfCancellationRequested();
                result[pageEntry.Key] = new Dictionary<string, PdfObject>(pageEntry.Value);
            }
        }
    
        var visitedPages = new HashSet<int>();
        foreach (int pageObjectNumber in pageObjectNumbers) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!visitedPages.Add(pageObjectNumber) ||
                !sourceObjects.TryGetValue(pageObjectNumber, out var pageObject) ||
                pageObject.Value is not PdfDictionary pageDictionary) {
                continue;
            }
    
            Dictionary<string, PdfObject>? existingOverrides = null;
            if (result is not null) {
                result.TryGetValue(pageObjectNumber, out existingOverrides);
            }
            PdfObject? annotationsObject = existingOverrides is not null && existingOverrides.TryGetValue("Annots", out var overrideAnnotations)
                ? overrideAnnotations
                : pageDictionary.Items.TryGetValue("Annots", out var pageAnnotations) ? pageAnnotations : null;
    
            if (!TryFilterLinkAnnotations(sourceObjects, annotationsObject, availableDestinationNames, copiedPageObjectIds, out var filteredAnnotations, cancellationToken)) {
                continue;
            }
    
            result ??= new Dictionary<int, Dictionary<string, PdfObject>>();
            if (!result.TryGetValue(pageObjectNumber, out var overrides)) {
                overrides = new Dictionary<string, PdfObject>();
                result[pageObjectNumber] = overrides;
            }
    
            overrides["Annots"] = filteredAnnotations;
        }
    
        return result ?? pageOverrides;
    }
    
    private static HashSet<string> GetNamedDestinationNames(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        CatalogRewriteState catalogState,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var names = new HashSet<string>(StringComparer.Ordinal);
    
        if (ResolveDictionary(sourceObjects, catalogState.NamedDestinations) is PdfDictionary directDestinations) {
            foreach (var name in directDestinations.Items.Keys) {
                cancellationToken.ThrowIfCancellationRequested();
                names.Add(name);
            }
        }
    
        if (ResolveDictionary(sourceObjects, catalogState.NamedDestinationNameTree) is PdfDictionary nameTree &&
            nameTree.Items.TryGetValue("Names", out var namesObject) &&
            ResolveObject(sourceObjects, namesObject) is PdfArray nameArray) {
            for (int i = 0; i + 1 < nameArray.Items.Count; i += 2) {
                cancellationToken.ThrowIfCancellationRequested();
                if (TryGetNamedDestinationName(sourceObjects, nameArray.Items[i], out string? destinationName)) {
                    names.Add(destinationName!);
                }
            }
        }
    
        return names;
    }
    
    private static bool TryFilterLinkAnnotations(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? annotationsObject,
        HashSet<string> availableDestinationNames,
        HashSet<int> copiedPageObjectIds,
        out PdfArray filteredAnnotations,
        CancellationToken cancellationToken = default) {
        filteredAnnotations = new PdfArray();
        if (ResolveObject(sourceObjects, annotationsObject) is not PdfArray annotations) {
            return false;
        }
    
        bool removed = false;
        foreach (var annotation in annotations.Items) {
            cancellationToken.ThrowIfCancellationRequested();
            if (TryGetNamedDestinationLinkName(sourceObjects, annotation, out string? destinationName) &&
                !availableDestinationNames.Contains(destinationName!)) {
                removed = true;
                continue;
            }

            if (TryGetDirectDestinationLink(sourceObjects, annotation, out PdfObject? destination) &&
                !IsDestinationForCopiedPages(destination!, copiedPageObjectIds)) {
                removed = true;
                continue;
            }
    
            filteredAnnotations.Items.Add(annotation);
        }
    
        return removed;
    }
    
    private static bool TryGetNamedDestinationLinkName(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject annotationObject,
        out string? destinationName) {
        destinationName = null;
        if (ResolveDictionary(sourceObjects, annotationObject) is not PdfDictionary annotation ||
            annotation.Get<PdfName>("Subtype")?.Name != "Link") {
            return false;
        }
    
        if (annotation.Items.TryGetValue("Dest", out var directDestination) &&
            TryGetNamedDestinationName(sourceObjects, directDestination, out destinationName)) {
            return true;
        }
    
        if (!annotation.Items.TryGetValue("A", out var actionObject) ||
            ResolveDictionary(sourceObjects, actionObject) is not PdfDictionary action ||
            action.Get<PdfName>("S")?.Name != "GoTo" ||
            !action.Items.TryGetValue("D", out var actionDestination)) {
            return false;
        }
    
        return TryGetNamedDestinationName(sourceObjects, actionDestination, out destinationName);
    }

    private static bool TryGetDirectDestinationLink(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject annotationObject,
        out PdfObject? destination) {
        destination = null;
        if (ResolveDictionary(sourceObjects, annotationObject) is not PdfDictionary annotation ||
            annotation.Get<PdfName>("Subtype")?.Name != "Link") {
            return false;
        }

        if (annotation.Items.TryGetValue("Dest", out var directDestination) &&
            TryResolveDirectDestination(sourceObjects, directDestination, out destination)) {
            return true;
        }

        if (!annotation.Items.TryGetValue("A", out var actionObject) ||
            ResolveDictionary(sourceObjects, actionObject) is not PdfDictionary action ||
            action.Get<PdfName>("S")?.Name != "GoTo" ||
            !action.Items.TryGetValue("D", out var actionDestination) ||
            !TryResolveDirectDestination(sourceObjects, actionDestination, out destination)) {
            return false;
        }

        return true;
    }

    private static bool TryResolveDirectDestination(Dictionary<int, PdfIndirectObject> sourceObjects, PdfObject? destinationObject, out PdfObject? destination) {
        destination = ResolveObject(sourceObjects, destinationObject);
        if (destination is PdfDictionary dictionary &&
            dictionary.Items.TryGetValue("D", out var explicitDestination)) {
            destination = ResolveObject(sourceObjects, explicitDestination);
        }

        if (destination is PdfArray array &&
            array.Items.Count > 0 &&
            array.Items[0] is PdfReference) {
            return true;
        }

        destination = null;
        return false;
    }
    
    private static bool TryGetNamedDestinationName(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? destinationObject,
        out string? destinationName) {
        destinationName = null;
        PdfObject? destination = ResolveObject(sourceObjects, destinationObject);
        if (destination is PdfStringObj text) {
            destinationName = text.Value;
            return true;
        }
    
        if (destination is PdfName name) {
            destinationName = name.Name;
            return true;
        }
    
        return false;
    }
}
