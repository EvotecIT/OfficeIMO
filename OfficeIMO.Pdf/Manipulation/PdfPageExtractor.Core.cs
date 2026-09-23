using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf;

internal static partial class PdfPageExtractor {
    internal static byte[] ExtractPages(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfMetadata metadata,
        int[] pageObjectNumbers,
        Dictionary<int, Dictionary<string, PdfObject>>? pageOverrides = null,
        IEnumerable<AdditionalObject>? additionalObjects = null,
        CatalogRewriteState? catalogState = null,
        PdfFileVersion fileVersion = PdfFileVersion.Pdf14,
        long? maximumOutputBytes = null,
        Action<IReadOnlyDictionary<int, int>>? captureObjectNumbers = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (maximumOutputBytes <= 0L) throw new ArgumentOutOfRangeException(nameof(maximumOutputBytes));
        catalogState ??= CatalogRewriteState.Empty;
        var copiedPageObjectIds = new HashSet<int>(pageObjectNumbers);
        catalogState = PruneCatalogStateForPages(sourceObjects, catalogState, copiedPageObjectIds, pageObjectNumbers,
            cancellationToken: cancellationToken);
        pageOverrides = BuildPageOverridesWithFilteredDestinationLinks(sourceObjects, pageObjectNumbers, pageOverrides,
            catalogState, copiedPageObjectIds, cancellationToken);
    
        var collector = new ObjectCollector(sourceObjects, pageOverrides, cancellationToken);
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
        int numberMapCapacity = GetBoundedExtractionCollectionCount(sourceIds.Count, extraObjects.Length);
        var numberMap = new Dictionary<int, int>(numberMapCapacity);
        for (int i = 0; i < sourceIds.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
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
        var seenPages = PdfCollectionSizing.CreateHashSet<int>(pageObjectNumbers.Length, pageObjectNumbers.Length);
        var outputPageObjectIds = new int[pageObjectNumbers.Length];
        for (int i = 0; i < pageObjectNumbers.Length; i++) {
            cancellationToken.ThrowIfCancellationRequested();
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
        var context = new SerializationContext(numberMap, pagesId, collector.MaterializedPageValues, sourceObjects, pageOverrides,
            cancellationToken: cancellationToken);
        int serializedObjectCapacity = GetSerializedObjectCapacity(sourceIds.Count, extraObjects.Length, clonedPages);
        var objects = new List<PdfSerializedObject>(serializedObjectCapacity);
        long serializedObjectBytes = 0L;
        bool enforceOutputLimit = maximumOutputBytes.HasValue;
        long objectBytesLimit = maximumOutputBytes ?? long.MaxValue;
    
        foreach (int sourceId in sourceIds) {
            cancellationToken.ThrowIfCancellationRequested();
            if (!sourceObjects.TryGetValue(sourceId, out var sourceObject)) {
                throw new InvalidOperationException("PDF object " + sourceId.ToString(CultureInfo.InvariantCulture) + " was referenced but not found.");
            }
    
            int newId = numberMap[sourceId];
            if (enforceOutputLimit) {
                PdfObject sizeCheckValue = sourceObject.Value is PdfDictionary pageDictionary && collector.PageObjectIds.Contains(sourceId)
                    ? BuildPageDictionaryForSizeCheck(pageDictionary, sourceId, context)
                    : sourceObject.Value;
                EnsureSerializedIndirectObjectWithinLimit(
                    sizeCheckValue,
                    context,
                    newId,
                    objectBytesLimit - serializedObjectBytes - (ReferenceEquals(sizeCheckValue, sourceObject.Value) ? 0L : 64L));
            }
            PdfSerializedObject serializedObject = sourceObject.Value is PdfDictionary dictionary && collector.PageObjectIds.Contains(sourceId)
                ? PdfSerializedObject.FromBytes(WrapObject(newId, SerializePageDictionary(dictionary, sourceId, context)))
                : SerializeIndirectObjectForAssembly(newId, sourceObject.Value, context);

            AddBoundedObject(objects, serializedObject, objectBytesLimit, ref serializedObjectBytes);
        }
    
        foreach (var extraObject in extraObjects) {
            cancellationToken.ThrowIfCancellationRequested();
            int newId = numberMap[extraObject.PseudoObjectNumber];
            if (enforceOutputLimit) {
                EnsureSerializedIndirectObjectWithinLimit(
                    extraObject.Value,
                    context,
                    newId,
                    objectBytesLimit - serializedObjectBytes);
            }
            PdfSerializedObject serializedObject = SerializeIndirectObjectForAssembly(newId, extraObject.Value, context);
            AddBoundedObject(objects, serializedObject, objectBytesLimit, ref serializedObjectBytes);
        }
    
        foreach (var clonedPage in clonedPages) {
            cancellationToken.ThrowIfCancellationRequested();
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
            var clonedContext = new SerializationContext(clonedNumberMap, pagesId, collector.MaterializedPageValues, sourceObjects, clonedPageOverrides,
                cancellationToken: cancellationToken);
            if (enforceOutputLimit) {
                PdfDictionary clonedSizeCheckDictionary = BuildPageDictionaryForSizeCheck(dictionary, clonedPage.SourcePageObjectNumber, clonedContext);
                EnsureSerializedIndirectObjectWithinLimit(
                    clonedSizeCheckDictionary,
                    clonedContext,
                    clonedPage.OutputPageObjectNumber,
                    objectBytesLimit - serializedObjectBytes - 64L);
            }
            PdfSerializedObject serializedPage = PdfSerializedObject.FromBytes(WrapObject(
                clonedPage.OutputPageObjectNumber,
                SerializePageDictionary(dictionary, clonedPage.SourcePageObjectNumber, clonedContext)));
            AddBoundedObject(objects, serializedPage, objectBytesLimit, ref serializedObjectBytes);
    
            foreach (var annotation in clonedPage.AnnotationObjectMap) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!sourceObjects.TryGetValue(annotation.Key, out var annotationObject)) {
                    throw new InvalidOperationException("PDF annotation object " + annotation.Key.ToString(CultureInfo.InvariantCulture) + " was referenced but not found.");
                }
    
                if (enforceOutputLimit) {
                    EnsureSerializedIndirectObjectWithinLimit(
                        annotationObject.Value,
                        clonedContext,
                        annotation.Value,
                        objectBytesLimit - serializedObjectBytes);
                }
                PdfSerializedObject serializedAnnotation = SerializeIndirectObjectForAssembly(annotation.Value, annotationObject.Value, clonedContext);
                AddBoundedObject(objects, serializedAnnotation, objectBytesLimit, ref serializedObjectBytes);
            }
        }
    
        AddBoundedObject(objects, PdfSerializedObject.FromBytes(WrapObject(pagesId, PdfEncoding.Latin1GetBytes(PdfPageTreeBuilder.BuildPagesDictionary(outputPageObjectIds, cancellationToken)))), objectBytesLimit, ref serializedObjectBytes);
        AddBoundedObject(objects, PdfSerializedObject.FromBytes(WrapObject(catalogId, PdfEncoding.Latin1GetBytes(BuildCatalogDictionary(pagesId, catalogState, context)))), objectBytesLimit, ref serializedObjectBytes);
        AddBoundedObject(objects, PdfSerializedObject.FromBytes(WrapObject(infoId, PdfEncoding.Latin1GetBytes(BuildInfoDictionary(metadata)))), objectBytesLimit, ref serializedObjectBytes);
    
        byte[] result = maximumOutputBytes.HasValue
            ? AssembleBounded(objects, catalogId, infoId, fileVersion, maximumOutputBytes.Value, cancellationToken)
            : Assemble(objects, catalogId, infoId, fileVersion, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        captureObjectNumbers?.Invoke(numberMap);
        return result;
    }

    private static int GetSerializedObjectCapacity(
        int sourceObjectCount,
        int additionalObjectCount,
        IReadOnlyList<ClonedPageObject> clonedPages) {
        long total = (long)sourceObjectCount + additionalObjectCount + clonedPages.Count + 3L;
        for (int index = 0; index < clonedPages.Count; index++) {
            total += clonedPages[index].AnnotationObjectMap.Count;
        }
        return GetBoundedExtractionCollectionCount(total);
    }

    private static int GetBoundedExtractionCollectionCount(int first, int second) =>
        GetBoundedExtractionCollectionCount((long)first + second);

    private static int GetBoundedExtractionCollectionCount(long total) {
        if (total > int.MaxValue) {
            throw PdfOutputLimitErrors.Create("The extracted PDF exceeds the supported in-memory collection limits.");
        }
        return (int)total;
    }

    private static void AddBoundedObject(
        List<PdfSerializedObject> objects,
        PdfSerializedObject indirectObject,
        long maximumObjectBytes,
        ref long serializedObjectBytes) {
        if (serializedObjectBytes > maximumObjectBytes - indirectObject.Length) {
            throw new InvalidDataException("The extracted PDF exceeds the configured output limit.");
        }
        serializedObjectBytes += indirectObject.Length;
        objects.Add(indirectObject);
    }

    private static byte[] AssembleBounded(
        IReadOnlyList<PdfSerializedObject> objects,
        int catalogId,
        int infoId,
        PdfFileVersion fileVersion,
        long maximumOutputBytes,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        long assembledLength = PdfFileAssembler.GetAssembledLength(objects, catalogId, infoId, fileVersion,
            cancellationToken: cancellationToken);
        if (assembledLength > maximumOutputBytes) {
            throw new InvalidDataException("The extracted PDF exceeds the configured output limit.");
        }
        if (assembledLength > int.MaxValue) {
            throw new InvalidDataException("The extracted PDF exceeds the supported in-memory result size.");
        }
        using FileStream output = PdfTemporaryFile.Create(".extract", FileOptions.RandomAccess, out _);
        using var boundedOutput = new PdfBoundedWriteStream(
            output,
            Math.Min(maximumOutputBytes, int.MaxValue),
            "The extracted PDF exceeds the configured output limit.");
        PdfFileAssembler.Assemble(
            boundedOutput,
            objects,
            catalogId,
            infoId,
            fileVersion,
            cancellationToken: cancellationToken);
        boundedOutput.Flush();
        if (output.Length > int.MaxValue) {
            throw new InvalidDataException("The extracted PDF exceeds the supported in-memory result size.");
        }
        var bytes = new byte[(int)output.Length];
        output.Position = 0L;
        int read = 0;
        while (read < bytes.Length) {
            cancellationToken.ThrowIfCancellationRequested();
            int count = output.Read(bytes, read, bytes.Length - read);
            if (count == 0) throw new EndOfStreamException("The temporary extracted PDF ended unexpectedly.");
            read += count;
        }
        return bytes;
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
