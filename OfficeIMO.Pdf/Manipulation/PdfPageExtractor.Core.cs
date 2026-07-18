using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    internal static byte[] ExtractPages(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfMetadata metadata,
        int[] pageObjectNumbers,
        Dictionary<int, Dictionary<string, PdfObject>>? pageOverrides = null,
        IEnumerable<AdditionalObject>? additionalObjects = null,
        CatalogRewriteState? catalogState = null,
        PdfFileVersion fileVersion = PdfFileVersion.Pdf14) {
        catalogState ??= CatalogRewriteState.Empty;
        var copiedPageObjectIds = new HashSet<int>(pageObjectNumbers);
        catalogState = PruneCatalogStateForPages(sourceObjects, catalogState, copiedPageObjectIds, pageObjectNumbers);
        pageOverrides = BuildPageOverridesWithFilteredDestinationLinks(sourceObjects, pageObjectNumbers, pageOverrides, catalogState, copiedPageObjectIds);
    
        var collector = new ObjectCollector(sourceObjects, pageOverrides);
        foreach (int pageObjectNumber in pageObjectNumbers) {
            collector.CollectPage(pageObjectNumber);
        }
    
        collector.CollectObjectGraph(catalogState.Outlines);
        collector.CollectObjectGraph(catalogState.PageLabels);
        collector.CollectObjectGraph(catalogState.NamedDestinationNameTree);
        collector.CollectObjectGraph(catalogState.OpenAction);
        collector.CollectObjectGraph(catalogState.XmpMetadata);
        collector.CollectObjectGraph(catalogState.CatalogUri);
        collector.CollectObjectGraph(catalogState.OutputIntents);
        collector.CollectObjectGraph(catalogState.EmbeddedFiles);
        collector.CollectObjectGraph(catalogState.AssociatedFiles);
        collector.CollectObjectGraph(catalogState.OptionalContent);
        var extraObjects = additionalObjects?.ToArray() ?? Array.Empty<AdditionalObject>();
        foreach (var extraObject in extraObjects) {
            collector.CollectObjectGraph(extraObject.Value);
        }

        var sourceIds = collector.ObjectIds;
        var numberMap = new Dictionary<int, int>();
        for (int i = 0; i < sourceIds.Count; i++) {
            numberMap[sourceIds[i]] = i + 1;
        }
    
        int nextObjectId = sourceIds.Count + 1;
        foreach (var extraObject in extraObjects) {
            if (numberMap.ContainsKey(extraObject.PseudoObjectNumber)) {
                throw new InvalidOperationException("Additional PDF object id collides with a copied source object.");
            }
    
            numberMap[extraObject.PseudoObjectNumber] = nextObjectId++;
        }
    
        var clonedPages = new List<ClonedPageObject>();
        var seenPages = new HashSet<int>();
        var outputPageObjectIds = new int[pageObjectNumbers.Length];
        for (int i = 0; i < pageObjectNumbers.Length; i++) {
            int pageObjectNumber = pageObjectNumbers[i];
            if (seenPages.Add(pageObjectNumber)) {
                outputPageObjectIds[i] = numberMap[pageObjectNumber];
                continue;
            }
    
            int clonedPageObjectId = nextObjectId++;
            outputPageObjectIds[i] = clonedPageObjectId;
            Dictionary<string, PdfObject>? sourcePageOverrides = pageOverrides is not null && pageOverrides.TryGetValue(pageObjectNumber, out var overrides)
                ? overrides
                : null;
            var clonedAnnotationState = BuildClonedAnnotationState(sourceObjects, pageObjectNumber, sourcePageOverrides, ref nextObjectId);
            clonedPages.Add(new ClonedPageObject(pageObjectNumber, clonedPageObjectId, clonedAnnotationState.PageOverrides, clonedAnnotationState.AnnotationObjectMap));
        }
    
        int pagesId = nextObjectId++;
        int catalogId = nextObjectId++;
        int infoId = nextObjectId;
        var context = new SerializationContext(numberMap, pagesId, collector.MaterializedPageValues, sourceObjects, pageOverrides);
        var objects = new List<byte[]>(sourceIds.Count + 3);
    
        foreach (int sourceId in sourceIds) {
            if (!sourceObjects.TryGetValue(sourceId, out var sourceObject)) {
                throw new InvalidOperationException("PDF object " + sourceId.ToString(CultureInfo.InvariantCulture) + " was referenced but not found.");
            }
    
            int newId = numberMap[sourceId];
            byte[] body = sourceObject.Value is PdfDictionary dictionary && collector.PageObjectIds.Contains(sourceId)
                ? SerializePageDictionary(dictionary, sourceId, context)
                : SerializeObject(sourceObject.Value, context);
    
            objects.Add(WrapObject(newId, body));
        }
    
        foreach (var extraObject in extraObjects) {
            objects.Add(WrapObject(numberMap[extraObject.PseudoObjectNumber], SerializeObject(extraObject.Value, context)));
        }
    
        foreach (var clonedPage in clonedPages) {
            if (!sourceObjects.TryGetValue(clonedPage.SourcePageObjectNumber, out var sourceObject) ||
                sourceObject.Value is not PdfDictionary dictionary) {
                throw new InvalidOperationException("PDF page object " + clonedPage.SourcePageObjectNumber.ToString(CultureInfo.InvariantCulture) + " was referenced but not found.");
            }
    
            var clonedNumberMap = new Dictionary<int, int>(numberMap) {
                [clonedPage.SourcePageObjectNumber] = clonedPage.OutputPageObjectNumber
            };
            foreach (var annotation in clonedPage.AnnotationObjectMap) {
                clonedNumberMap[annotation.Key] = annotation.Value;
            }
    
            var clonedPageOverrides = clonedPage.PageOverrides is null
                ? null
                : new Dictionary<int, Dictionary<string, PdfObject>> {
                    [clonedPage.SourcePageObjectNumber] = clonedPage.PageOverrides
                };
            var clonedContext = new SerializationContext(clonedNumberMap, pagesId, collector.MaterializedPageValues, sourceObjects, clonedPageOverrides);
            byte[] body = SerializePageDictionary(dictionary, clonedPage.SourcePageObjectNumber, clonedContext);
            objects.Add(WrapObject(clonedPage.OutputPageObjectNumber, body));
    
            foreach (var annotation in clonedPage.AnnotationObjectMap) {
                if (!sourceObjects.TryGetValue(annotation.Key, out var annotationObject)) {
                    throw new InvalidOperationException("PDF annotation object " + annotation.Key.ToString(CultureInfo.InvariantCulture) + " was referenced but not found.");
                }
    
                objects.Add(WrapObject(annotation.Value, SerializeObject(annotationObject.Value, clonedContext)));
            }
        }
    
        objects.Add(WrapObject(pagesId, PdfEncoding.Latin1GetBytes(PdfPageTreeBuilder.BuildPagesDictionary(outputPageObjectIds))));
        objects.Add(WrapObject(catalogId, PdfEncoding.Latin1GetBytes(BuildCatalogDictionary(pagesId, catalogState, context))));
        objects.Add(WrapObject(infoId, PdfEncoding.Latin1GetBytes(BuildInfoDictionary(metadata))));
    
        return Assemble(objects, catalogId, infoId, fileVersion);
    }
    
    private static ClonedAnnotationState BuildClonedAnnotationState(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        int pageObjectNumber,
        Dictionary<string, PdfObject>? pageOverrides,
        ref int nextObjectId) {
        PdfObject? annotationsObject = pageOverrides is not null && pageOverrides.TryGetValue("Annots", out var overrideAnnotations)
            ? overrideAnnotations
            : null;
        if (!sourceObjects.TryGetValue(pageObjectNumber, out var pageObject) ||
            pageObject.Value is not PdfDictionary pageDictionary) {
            return ClonedAnnotationState.Empty;
        }

        if (annotationsObject is null &&
            !pageDictionary.Items.TryGetValue("Annots", out annotationsObject)) {
            return ClonedAnnotationState.Empty;
        }

        if (ResolveObject(sourceObjects, annotationsObject) is not PdfArray annotations) {
            return ClonedAnnotationState.Empty;
        }
    
        var annotationObjectMap = new Dictionary<int, int>();
        var clonedAnnotations = new PdfArray();
        bool hasClonedIndirectAnnotation = false;
    
        foreach (var annotation in annotations.Items) {
            if (annotation is PdfReference annotationReference &&
                PdfObjectLookup.TryGet(sourceObjects, annotationReference, out _)) {
                if (!annotationObjectMap.TryGetValue(annotationReference.ObjectNumber, out int clonedAnnotationObjectNumber)) {
                    clonedAnnotationObjectNumber = nextObjectId++;
                    annotationObjectMap[annotationReference.ObjectNumber] = clonedAnnotationObjectNumber;
                }
    
                clonedAnnotations.Items.Add(new PdfReference(annotationReference.ObjectNumber, annotationReference.Generation));
                hasClonedIndirectAnnotation = true;
                continue;
            }
    
            clonedAnnotations.Items.Add(annotation);
        }
    
        if (!hasClonedIndirectAnnotation && pageOverrides is null) {
            return ClonedAnnotationState.Empty;
        }

        var clonedPageOverrides = pageOverrides is null
            ? new Dictionary<string, PdfObject>(StringComparer.Ordinal)
            : new Dictionary<string, PdfObject>(pageOverrides, StringComparer.Ordinal);
        clonedPageOverrides["Annots"] = clonedAnnotations;

        return new ClonedAnnotationState(
            clonedPageOverrides,
            annotationObjectMap);
    }
    
    private static void ValidatePageNumbers(int[] pageNumbers, int pageCount, string paramName) {
        for (int i = 0; i < pageNumbers.Length; i++) {
            int pageNumber = pageNumbers[i];
            if (pageNumber < 1 || pageNumber > pageCount) {
                throw new ArgumentOutOfRangeException(paramName, "Page number " + pageNumber.ToString(CultureInfo.InvariantCulture) + " is outside the document page range 1-" + pageCount.ToString(CultureInfo.InvariantCulture) + ".");
            }
        }
    }
    
    private static void ValidatePageRanges(PdfPageRange[] ranges, int pageCount, string paramName) {
        for (int i = 0; i < ranges.Length; i++) {
            var range = ranges[i];
            if (range.FirstPage < 1) {
                throw new ArgumentOutOfRangeException(paramName, "Page range first page must be 1 or greater.");
            }
    
            if (range.LastPage < range.FirstPage) {
                throw new ArgumentOutOfRangeException(paramName, "Page range last page must be greater than or equal to first page.");
            }
    
            if (range.LastPage > pageCount) {
                throw new ArgumentOutOfRangeException(paramName, "Page range " + range.ToString() + " is outside the document page range 1-" + pageCount.ToString(CultureInfo.InvariantCulture) + ".");
            }
        }
    }
}
