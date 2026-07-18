using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    internal static string BuildCatalogDictionary(int pagesId, CatalogRewriteState? catalogState, SerializationContext? context = null) {
        var sb = new StringBuilder();
        PdfCatalogDictionaryBuilder.AppendCatalogStart(sb, pagesId);
    
        string? pageMode = catalogState?.PageMode;
        if (pageMode is not null && pageMode.Length > 0) {
            PdfCatalogDictionaryBuilder.AppendNameEntry(sb, "PageMode", pageMode);
        }
    
        string? pageLayout = catalogState?.PageLayout;
        if (pageLayout is not null && pageLayout.Length > 0) {
            PdfCatalogDictionaryBuilder.AppendNameEntry(sb, "PageLayout", pageLayout);
        }
    
        PdfObject? catalogVersion = catalogState?.CatalogVersion;
        if (catalogVersion is not null) {
            if (context is null) {
                throw new InvalidOperationException("A serialization context is required to preserve catalog version.");
            }
    
            sb.Append(" /Version ");
            AppendObject(sb, catalogVersion, context);
        }
    
        PdfObject? catalogLanguage = catalogState?.CatalogLanguage;
        if (catalogLanguage is not null) {
            if (context is null) {
                throw new InvalidOperationException("A serialization context is required to preserve catalog language.");
            }
    
            sb.Append(" /Lang ");
            AppendObject(sb, catalogLanguage, context);
        }
    
        PdfObject? pageLabels = catalogState?.PageLabels;
        if (pageLabels is not null) {
            if (context is null) {
                throw new InvalidOperationException("A serialization context is required to preserve catalog page labels.");
            }
    
            sb.Append(" /PageLabels ");
            AppendObject(sb, pageLabels, context);
        }
    
        PdfObject? outlines = catalogState?.Outlines;
        if (outlines is not null) {
            if (context is null) {
                throw new InvalidOperationException("A serialization context is required to preserve catalog outlines.");
            }
    
            sb.Append(" /Outlines ");
            AppendObject(sb, outlines, context);
        }
    
        PdfObject? namedDestinations = catalogState?.NamedDestinations;
        if (namedDestinations is not null) {
            if (context is null) {
                throw new InvalidOperationException("A serialization context is required to preserve catalog named destinations.");
            }
    
            sb.Append(" /Dests ");
            AppendObject(sb, namedDestinations, context);
        }
    
        PdfObject? openAction = catalogState?.OpenAction;
        if (openAction is not null) {
            if (context is null) {
                throw new InvalidOperationException("A serialization context is required to preserve catalog open actions.");
            }
    
            sb.Append(" /OpenAction ");
            AppendObject(sb, openAction, context);
        }
    
        PdfObject? viewerPreferences = catalogState?.ViewerPreferences;
        if (viewerPreferences is not null) {
            if (context is null) {
                throw new InvalidOperationException("A serialization context is required to preserve catalog viewer preferences.");
            }
    
            sb.Append(" /ViewerPreferences ");
            AppendObject(sb, viewerPreferences, context);
        }
    
        PdfObject? xmpMetadata = catalogState?.XmpMetadata;
        if (xmpMetadata is not null) {
            if (context is null) {
                throw new InvalidOperationException("A serialization context is required to preserve catalog XMP metadata.");
            }
    
            sb.Append(" /Metadata ");
            AppendObject(sb, xmpMetadata, context);
        }
    
        PdfObject? catalogUri = catalogState?.CatalogUri;
        if (catalogUri is not null) {
            if (context is null) {
                throw new InvalidOperationException("A serialization context is required to preserve catalog URI settings.");
            }
    
            sb.Append(" /URI ");
            AppendObject(sb, catalogUri, context);
        }
    
        PdfObject? outputIntents = catalogState?.OutputIntents;
        if (outputIntents is not null) {
            if (context is null) {
                throw new InvalidOperationException("A serialization context is required to preserve catalog output intents.");
            }
    
            sb.Append(" /OutputIntents ");
            AppendObject(sb, outputIntents, context);
        }
    
        PdfObject? namedDestinationNameTree = catalogState?.NamedDestinationNameTree;
        PdfObject? embeddedFiles = catalogState?.EmbeddedFiles;
        if (namedDestinationNameTree is not null || embeddedFiles is not null) {
            if (context is null) {
                throw new InvalidOperationException("A serialization context is required to preserve catalog name trees.");
            }
    
            sb.Append(" /Names <<");
            if (namedDestinationNameTree is not null) {
                sb.Append(" /Dests ");
                AppendObject(sb, namedDestinationNameTree, context);
            }
    
            if (embeddedFiles is not null) {
                sb.Append(" /EmbeddedFiles ");
                AppendObject(sb, embeddedFiles, context);
            }
    
            sb.Append(" >>");
        }
    
        PdfObject? associatedFiles = catalogState?.AssociatedFiles;
        if (associatedFiles is not null) {
            if (context is null) {
                throw new InvalidOperationException("A serialization context is required to preserve catalog associated files.");
            }
    
            sb.Append(" /AF ");
            AppendObject(sb, associatedFiles, context);
        }
    
        PdfObject? optionalContent = catalogState?.OptionalContent;
        if (optionalContent is not null) {
            if (context is null) {
                throw new InvalidOperationException("A serialization context is required to preserve catalog optional content.");
            }
    
            sb.Append(" /OCProperties ");
            AppendObject(sb, optionalContent, context);
        }
    
        sb.Append(" >>\n");
        return sb.ToString();
    }
    
    private static PdfObject? BuildOutlines(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? outlines) {
        return outlines is not null &&
            IsSupportedOutlineGraph(sourceObjects, outlines, new HashSet<int>())
            ? outlines
            : null;
    }
    
    private static PdfObject? BuildOutlinesForPages(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? outlines,
        HashSet<int> copiedPageObjectIds) {
        return outlines is not null &&
            OutlineDestinationsReferenceOnlyCopiedPages(sourceObjects, outlines, copiedPageObjectIds, new HashSet<int>())
            ? outlines
            : null;
    }
    
    private static PdfObject? BuildOptionalContent(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? optionalContent) {
        return optionalContent is not null &&
            IsSupportedCatalogMetadataGraph(sourceObjects, optionalContent, new HashSet<int>())
            ? optionalContent
            : null;
    }
    
    private static PdfName? BuildCatalogVersion(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? catalogVersion) {
        return ResolveObject(sourceObjects, catalogVersion) is PdfName name
            ? name
            : null;
    }
    
    private static PdfStringObj? BuildCatalogLanguage(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? catalogLanguage) {
        return ResolveObject(sourceObjects, catalogLanguage) is PdfStringObj text
            ? text
            : null;
    }
    
    private static PdfObject? BuildPageLabelsForPages(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? pageLabels,
        IReadOnlyList<int>? orderedPageObjectNumbers,
        int outputPageIndexOffset,
        IReadOnlyDictionary<int, int>? outputPageIndexByPageObjectNumber,
        IReadOnlyList<int>? sourcePageOrder = null) {
        if (pageLabels is null || orderedPageObjectNumbers is null || orderedPageObjectNumbers.Count == 0) {
            return pageLabels;
        }
    
        PdfDictionary? labelTree = ResolveDictionary(sourceObjects, pageLabels);
        if (labelTree is null ||
            labelTree.Items.ContainsKey("Kids") ||
            !labelTree.Items.TryGetValue("Nums", out var numsObject) ||
            ResolveObject(sourceObjects, numsObject) is not PdfArray nums ||
            nums.Items.Count % 2 != 0) {
            return pageLabels;
        }
    
        sourcePageOrder ??= GetPageObjectNumbersInDocumentOrder(sourceObjects);
        if (sourcePageOrder.Count == 0) {
            return pageLabels;
        }
    
        var sourcePageIndexes = new Dictionary<int, int>();
        for (int i = 0; i < sourcePageOrder.Count; i++) {
            if (!sourcePageIndexes.ContainsKey(sourcePageOrder[i])) {
                sourcePageIndexes[sourcePageOrder[i]] = i;
            }
        }
    
        var entries = new List<PageLabelEntry>();
        for (int i = 0; i < nums.Items.Count; i += 2) {
            if (ResolveObject(sourceObjects, nums.Items[i]) is not PdfNumber pageIndexNumber ||
                !TryGetNonNegativeInteger(pageIndexNumber, out int pageIndex) ||
                ResolveObject(sourceObjects, nums.Items[i + 1]) is not PdfDictionary labelDictionary) {
                return pageLabels;
            }
    
            entries.Add(new PageLabelEntry(pageIndex, labelDictionary));
        }
    
        if (entries.Count == 0) {
            return pageLabels;
        }
    
        entries.Sort((left, right) => left.StartPageIndex.CompareTo(right.StartPageIndex));
        var rewrittenNums = new PdfArray();
        PageLabelEntry? previousEntry = null;
        int previousSourcePageIndex = -1;
        int previousOutputPageIndex = -1;
    
        for (int outputIndex = 0; outputIndex < orderedPageObjectNumbers.Count; outputIndex++) {
            if (!sourcePageIndexes.TryGetValue(orderedPageObjectNumbers[outputIndex], out int sourcePageIndex)) {
                return pageLabels;
            }
    
            PageLabelEntry? entry = FindPageLabelEntry(entries, sourcePageIndex);
            if (entry is null) {
                continue;
            }
    
            int rewrittenOutputIndex = outputPageIndexByPageObjectNumber is not null &&
                outputPageIndexByPageObjectNumber.TryGetValue(orderedPageObjectNumbers[outputIndex], out int mappedOutputIndex)
                ? mappedOutputIndex
                : outputPageIndexOffset + outputIndex;
    
            bool continuesPreviousRun = previousEntry is not null &&
                ReferenceEquals(previousEntry.LabelDictionary, entry.LabelDictionary) &&
                sourcePageIndex == previousSourcePageIndex + 1 &&
                rewrittenOutputIndex == previousOutputPageIndex + 1;
    
            if (!continuesPreviousRun) {
                rewrittenNums.Items.Add(new PdfNumber(rewrittenOutputIndex));
                rewrittenNums.Items.Add(ClonePageLabelDictionary(entry.LabelDictionary, sourcePageIndex - entry.StartPageIndex));
            }
    
            previousEntry = entry;
            previousSourcePageIndex = sourcePageIndex;
            previousOutputPageIndex = rewrittenOutputIndex;
        }
    
        if (rewrittenNums.Items.Count == 0) {
            return null;
        }
    
        var rewrittenTree = new PdfDictionary();
        rewrittenTree.Items["Nums"] = rewrittenNums;
        return rewrittenTree;
    }
    
    private static List<int> GetPageObjectNumbersInDocumentOrder(Dictionary<int, PdfIndirectObject> sourceObjects, PdfDictionary? catalog = null) {
        var pages = new List<int>();
        if (catalog is not null &&
            catalog.Get<PdfName>("Type")?.Name == "Catalog" &&
            catalog.Items.TryGetValue("Pages", out var pagesRoot)) {
            CollectPageObjectNumbers(sourceObjects, pagesRoot, pages, new HashSet<int>());
            return pages;
        }
    
        foreach (var entry in sourceObjects) {
            if (entry.Value.Value is PdfDictionary scannedCatalog &&
                scannedCatalog.Get<PdfName>("Type")?.Name == "Catalog" &&
                scannedCatalog.Items.TryGetValue("Pages", out pagesRoot)) {
                CollectPageObjectNumbers(sourceObjects, pagesRoot, pages, new HashSet<int>());
                break;
            }
        }
    
        return pages;
    }
    
    private static void CollectPageObjectNumbers(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject pageNode,
        List<int> pages,
        HashSet<int> visitedObjects) {
        if (pageNode is PdfReference reference) {
            if (!visitedObjects.Add(reference.ObjectNumber) ||
                !PdfObjectLookup.TryGet(sourceObjects, reference, out var indirect)) {
                return;
            }
    
            if (indirect.Value is PdfDictionary referencedDictionary &&
                referencedDictionary.Get<PdfName>("Type")?.Name == "Page") {
                pages.Add(reference.ObjectNumber);
                return;
            }
    
            CollectPageObjectNumbers(sourceObjects, indirect.Value, pages, visitedObjects);
            return;
        }
    
        if (pageNode is not PdfDictionary dictionary ||
            !dictionary.Items.TryGetValue("Kids", out var kidsObject) ||
            ResolveObject(sourceObjects, kidsObject) is not PdfArray kids) {
            return;
        }
    
        foreach (var kid in kids.Items) {
            CollectPageObjectNumbers(sourceObjects, kid, pages, visitedObjects);
        }
    }
    
    private static PageLabelEntry? FindPageLabelEntry(IReadOnlyList<PageLabelEntry> entries, int sourcePageIndex) {
        PageLabelEntry? selected = null;
        for (int i = 0; i < entries.Count; i++) {
            if (entries[i].StartPageIndex > sourcePageIndex) {
                break;
            }
    
            selected = entries[i];
        }
    
        return selected;
    }
    
    private static PdfDictionary ClonePageLabelDictionary(PdfDictionary source, int sourcePageOffset) {
        var clone = new PdfDictionary();
        foreach (var entry in source.Items) {
            clone.Items[entry.Key] = entry.Value;
        }
    
        if (source.Items.ContainsKey("S")) {
            int start = 1;
            if (source.Get<PdfNumber>("St") is PdfNumber startNumber &&
                TryGetPositiveInteger(startNumber, out int parsedStart)) {
                start = parsedStart;
            }
    
            clone.Items["St"] = new PdfNumber(start + sourcePageOffset);
        }
    
        return clone;
    }
    
    private static bool TryGetNonNegativeInteger(PdfNumber number, out int value) {
        value = 0;
        if (number.Value < 0 || number.Value > int.MaxValue || Math.Truncate(number.Value) != number.Value) {
            return false;
        }
    
        value = (int)number.Value;
        return true;
    }
    
    private static bool TryGetPositiveInteger(PdfNumber number, out int value) {
        if (TryGetNonNegativeInteger(number, out value) && value > 0) {
            return true;
        }
    
        value = 0;
        return false;
    }
    
}
