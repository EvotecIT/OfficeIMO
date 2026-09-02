using OfficeIMO.Core.Internal;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>
/// Provides first-party PDF merge helpers for PDFs that can be parsed by OfficeIMO.Pdf.
/// </summary>
internal static partial class PdfMerger {
    /// <summary>
    /// Merges all pages from the supplied PDFs into one new PDF.
    /// </summary>
    public static byte[] Merge(params byte[][] pdfs) {
        return Merge((IEnumerable<byte[]>)pdfs);
    }

    /// <summary>
    /// Merges all pages from the supplied PDFs into one new PDF, applying optional source preparation first.
    /// </summary>
    public static byte[] Merge(PdfMergeOptions options, params byte[][] pdfs) {
        return Merge(options, (IEnumerable<byte[]>)pdfs);
    }

    /// <summary>
    /// Merges all pages from the supplied readable PDF streams into one new PDF.
    /// </summary>
    public static byte[] Merge(params Stream[] streams) {
        return Merge((IEnumerable<Stream>)streams);
    }

    /// <summary>
    /// Merges all pages from the supplied PDFs into one new PDF.
    /// </summary>
    public static byte[] Merge(IEnumerable<byte[]> pdfs) {
        return MergeCore(pdfs, primarySourceIndex: 0, options: null).ToBytes();
    }

    internal static byte[] Merge(IReadOnlyList<byte[]> pdfs, IReadOnlyList<PdfLoadOptions> readOptions) {
        Guard.NotNull(readOptions, nameof(readOptions));
        return MergeCore(pdfs, primarySourceIndex: 0, options: null, readOptions).ToBytes();
    }

    internal static PdfMergeResult MergeOwned(
        IReadOnlyList<byte[]> pdfs,
        IReadOnlyList<PdfLoadOptions> readOptions) {
        Guard.NotNull(readOptions, nameof(readOptions));
        return MergeCore(pdfs, primarySourceIndex: 0, options: null, readOptions);
    }

    internal static PdfMergeResult MergeOwned(
        IReadOnlyList<byte[]> pdfs,
        IReadOnlyList<PdfLoadOptions> readOptions,
        IReadOnlyList<Func<PdfReadDocument>?> readDocumentFactories,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(readOptions, nameof(readOptions));
        Guard.NotNull(readDocumentFactories, nameof(readDocumentFactories));
        return MergeCore(
            pdfs,
            primarySourceIndex: 0,
            options: null,
            readOptions,
            readDocumentFactories,
            cancellationToken);
    }

    internal static PdfMergeResult MergeWithReport(PdfMergeOptions options, IReadOnlyList<byte[]> pdfs, IReadOnlyList<PdfLoadOptions> readOptions) {
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(readOptions, nameof(readOptions));
        return MergeCore(pdfs, primarySourceIndex: 0, options, readOptions);
    }

    /// <summary>
    /// Merges all pages from the supplied PDFs into one new PDF, applying optional source preparation first.
    /// </summary>
    public static byte[] Merge(PdfMergeOptions options, IEnumerable<byte[]> pdfs) {
        Guard.NotNull(options, nameof(options));
        return MergeCore(pdfs, primarySourceIndex: 0, options).ToBytes();
    }

    /// <summary>Merges PDFs and returns the applied document-structure policy report.</summary>
    public static PdfMergeResult MergeWithReport(PdfMergeOptions options, params byte[][] pdfs) {
        return MergeWithReport(options, (IEnumerable<byte[]>)pdfs);
    }

    /// <summary>Merges PDFs and returns the applied document-structure policy report.</summary>
    public static PdfMergeResult MergeWithReport(PdfMergeOptions options, IEnumerable<byte[]> pdfs) {
        Guard.NotNull(options, nameof(options));
        return MergeCore(pdfs, primarySourceIndex: 0, options);
    }

    internal static byte[] MergeWithPrimarySource(int primarySourceIndex, params byte[][] pdfs) {
        return MergeCore(pdfs, primarySourceIndex, options: null).ToBytes();
    }

    internal static byte[] MergeWithPrimarySource(
        int primarySourceIndex,
        IReadOnlyList<byte[]> pdfs,
        IReadOnlyList<PdfLoadOptions> readOptions) {
        Guard.NotNull(readOptions, nameof(readOptions));
        return MergeCore(pdfs, primarySourceIndex, options: null, readOptions).ToBytes();
    }

    internal static byte[] MergePrimaryWithInsertedPages(byte[] primaryPdf, byte[] insertedPdf, int insertBeforePageNumber) {
        return MergePrimaryWithInsertedPages(primaryPdf, insertedPdf, insertBeforePageNumber, primaryReadOptions: null);
    }

    internal static byte[] MergePrimaryWithInsertedPages(
        byte[] primaryPdf,
        byte[] insertedPdf,
        int insertBeforePageNumber,
        PdfLoadOptions? primaryReadOptions) {
        Guard.NotNull(primaryPdf, nameof(primaryPdf));
        Guard.NotNull(insertedPdf, nameof(insertedPdf));

        if (PdfReadDocument.Open(primaryPdf, primaryReadOptions).AcroFormXfa is not null ||
            PdfReadDocument.Open(insertedPdf).AcroFormXfa is not null) {
            throw new NotSupportedException("Page insertion does not preserve XFA form packets. Flatten or remove XFA before inserting pages.");
        }

        var (_, primaryDocument) = PdfMutationPlanner.RequireFullRewriteDocument(
            primaryPdf,
            PdfMutationOperation.ModifyPageTree,
            primaryReadOptions);
        if (primaryDocument.Pages.Count == 0) {
            throw new ArgumentException("Primary PDF does not contain any pages.", nameof(primaryPdf));
        }

        if (insertBeforePageNumber < 1 || insertBeforePageNumber > primaryDocument.Pages.Count + 1) {
            throw new ArgumentOutOfRangeException(nameof(insertBeforePageNumber), "Insert-before page must be in the primary document page range.");
        }

        var (_, insertedDocument) = PdfMutationPlanner.RequireFullRewriteDocument(
            insertedPdf,
            PdfMutationOperation.ExtractPages);
        if (insertedDocument.Pages.Count == 0) {
            throw new ArgumentException("Inserted PDF does not contain any pages.", nameof(insertedPdf));
        }
        int[] primaryPageObjectNumbers = primaryDocument.Pages.Select(page => page.ObjectNumber).ToArray();
        int[] insertedPageObjectNumbers = insertedDocument.Pages.Select(page => page.ObjectNumber).ToArray();
        var outputOrder = new List<OutputPageReference>(primaryPageObjectNumbers.Length + insertedPageObjectNumbers.Length);
        var primaryPageIndexMap = new Dictionary<int, int>();

        for (int i = 0; i < insertBeforePageNumber - 1; i++) {
            primaryPageIndexMap[primaryPageObjectNumbers[i]] = outputOrder.Count;
            outputOrder.Add(new OutputPageReference(0, primaryPageObjectNumbers[i]));
        }

        for (int i = 0; i < insertedPageObjectNumbers.Length; i++) {
            outputOrder.Add(new OutputPageReference(1, insertedPageObjectNumbers[i]));
        }

        for (int i = insertBeforePageNumber - 1; i < primaryPageObjectNumbers.Length; i++) {
            primaryPageIndexMap[primaryPageObjectNumbers[i]] = outputOrder.Count;
            outputOrder.Add(new OutputPageReference(0, primaryPageObjectNumbers[i]));
        }

        var importedSources = new[] {
            ImportSource(primaryPdf, 0, primaryPageObjectNumbers, 0, primaryPageIndexMap, primaryReadOptions, primaryDocument),
            ImportSource(insertedPdf, 1, insertedPageObjectNumbers, insertBeforePageNumber - 1, null, plannedDocument: insertedDocument)
        };
        return WriteMerged(importedSources, primarySourceIndex: 0, outputOrder, out _);
    }

    private static PdfMergeResult MergeCore(
        IEnumerable<byte[]> pdfs,
        int primarySourceIndex,
        PdfMergeOptions? options,
        IReadOnlyList<PdfLoadOptions>? readOptions = null,
        IReadOnlyList<Func<PdfReadDocument>?>? readDocumentFactories = null,
        CancellationToken cancellationToken = default) {
        Guard.NotNull(pdfs, nameof(pdfs));
        cancellationToken.ThrowIfCancellationRequested();
        var sourceList = new List<byte[]>();
        foreach (byte[] source in pdfs) {
            cancellationToken.ThrowIfCancellationRequested();
            sourceList.Add(source);
        }
        byte[][] sources = sourceList.ToArray();
        if (sources.Length == 0) {
            throw new ArgumentException("At least one PDF must be supplied.", nameof(pdfs));
        }

        if (primarySourceIndex < 0 || primarySourceIndex >= sources.Length) {
            throw new ArgumentOutOfRangeException(nameof(primarySourceIndex), "Primary source index must refer to one of the supplied PDFs.");
        }

        if (readOptions is not null && readOptions.Count != sources.Length) {
            throw new ArgumentException("Read options must contain one entry for every PDF input.", nameof(readOptions));
        }
        if (readDocumentFactories is not null && readDocumentFactories.Count != sources.Length) {
            throw new ArgumentException("Read-document factories must contain one entry for every PDF input.", nameof(readDocumentFactories));
        }
        var importedSources = new List<ImportedSource>(sources.Length);
        int mergedPageOffset = 0;
        for (int i = 0; i < sources.Length; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            byte[] source = sources[i];
            if (source is null) {
                throw new ArgumentException("PDF input " + i.ToString(CultureInfo.InvariantCulture) + " cannot be null.", nameof(pdfs));
            }

            PdfLoadOptions? sourceReadOptions = readOptions?[i];
            Func<PdfReadDocument>? readDocumentFactory = readDocumentFactories?[i];
            (PdfMutationPlan sourceMergePlan, PdfReadDocument plannedDocument) = readDocumentFactory is null
                ? PdfMutationPlanner.RequireFullRewriteDocument(
                    source,
                    PdfMutationOperation.MergeDocuments,
                    sourceReadOptions,
                    cancellationToken: cancellationToken)
                : PdfMutationPlanner.RequireFullRewriteDocument(
                    source,
                    PdfMutationOperation.MergeDocuments,
                    readDocumentFactory,
                    sourceReadOptions,
                    cancellationToken: cancellationToken);
            PdfDocumentSecurityInfo sourceSecurity = sourceMergePlan.Preflight.Probe.Security;
            PdfPermissionPolicy sourcePermissionPolicy = sourceMergePlan.Preflight.PermissionPolicy;
            ValidateXfaSourceBeforePreparation(
                plannedDocument,
                i,
                primarySourceIndex,
                options?.Policy?.Forms ?? PdfMergeStructureMode.KeepPrimary);
            byte[] plannedSource = source;
            source = PrepareMergeSource(source, options, sourceReadOptions, out sourceReadOptions);
            importedSources.Add(ImportSource(
                source,
                i,
                null,
                mergedPageOffset,
                null,
                sourceReadOptions,
                plannedDocument: ReferenceEquals(source, plannedSource) ? plannedDocument : null,
                sourceSecurity: sourceSecurity,
                sourcePermissionPolicy: sourcePermissionPolicy,
                cancellationToken: cancellationToken));
            mergedPageOffset += importedSources[importedSources.Count - 1].PageObjectNumbers.Length;
        }

        byte[] merged = WriteMerged(importedSources, primarySourceIndex, outputOrder: null, out int mergedObjectCount, cancellationToken);
        PdfLoadOptions outputReadOptions = PdfLoadOptions.ForComposedOutput(
            importedSources[primarySourceIndex].Document.ReadOptions,
            importedSources.Select(static source => source.Document.ReadOptions),
            merged.LongLength,
            mergedObjectCount);
        return ApplyMergePolicy(merged, importedSources, primarySourceIndex, options, outputReadOptions, cancellationToken: cancellationToken);
    }

    /// <summary>
    /// Merges all pages from the supplied readable PDF streams into one new PDF, reading each stream from its current position.
    /// </summary>
    public static byte[] Merge(IEnumerable<Stream> streams) {
        Guard.NotNull(streams, nameof(streams));

        var sources = streams.ToArray();
        if (sources.Length == 0) {
            throw new ArgumentException("At least one PDF stream must be supplied.", nameof(streams));
        }

        var pdfs = new byte[sources.Length][];
        for (int i = 0; i < sources.Length; i++) {
            Stream stream = sources[i];
            if (stream is null) {
                throw new ArgumentException("PDF stream input " + i.ToString(CultureInfo.InvariantCulture) + " cannot be null.", nameof(streams));
            }

            if (!stream.CanRead) {
                throw new ArgumentException("PDF stream input " + i.ToString(CultureInfo.InvariantCulture) + " must be readable.", nameof(streams));
            }

            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            pdfs[i] = buffer.ToArray();
        }

        return Merge((IEnumerable<byte[]>)pdfs);
    }

    /// <summary>
    /// Merges all pages from the supplied readable PDF streams into one new PDF, applying optional source preparation first.
    /// </summary>
    public static byte[] Merge(PdfMergeOptions options, IEnumerable<Stream> streams) {
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(streams, nameof(streams));

        var sources = streams.ToArray();
        if (sources.Length == 0) {
            throw new ArgumentException("At least one PDF stream must be supplied.", nameof(streams));
        }

        var pdfs = new byte[sources.Length][];
        for (int i = 0; i < sources.Length; i++) {
            Stream stream = sources[i];
            if (stream is null) {
                throw new ArgumentException("PDF stream input " + i.ToString(CultureInfo.InvariantCulture) + " cannot be null.", nameof(streams));
            }

            if (!stream.CanRead) {
                throw new ArgumentException("PDF stream input " + i.ToString(CultureInfo.InvariantCulture) + " must be readable.", nameof(streams));
            }

            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            pdfs[i] = buffer.ToArray();
        }

        return Merge(options, (IEnumerable<byte[]>)pdfs);
    }

    /// <summary>
    /// Merges all pages from the supplied PDFs and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void Merge(IEnumerable<byte[]> pdfs, Stream outputStream) {
        WriteOutput(outputStream, Merge(pdfs));
    }

    /// <summary>
    /// Merges all pages from the supplied PDFs and writes the result to <paramref name="outputStream"/>, applying optional source preparation first.
    /// </summary>
    public static void Merge(PdfMergeOptions options, IEnumerable<byte[]> pdfs, Stream outputStream) {
        WriteOutput(outputStream, Merge(options, pdfs));
    }

    /// <summary>
    /// Merges all pages from the supplied readable PDF streams and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void Merge(IEnumerable<Stream> streams, Stream outputStream) {
        WriteOutput(outputStream, Merge(streams));
    }

    /// <summary>
    /// Merges all pages from the supplied readable PDF streams and writes the result to <paramref name="outputStream"/>, applying optional source preparation first.
    /// </summary>
    public static void Merge(PdfMergeOptions options, IEnumerable<Stream> streams, Stream outputStream) {
        WriteOutput(outputStream, Merge(options, streams));
    }

    /// <summary>
    /// Merges PDFs from file paths and writes the result to the output path.
    /// </summary>
    public static void MergeFiles(string outputPath, params string[] inputPaths) {
        Guard.NotNull(outputPath, nameof(outputPath));
        Guard.NotNull(inputPaths, nameof(inputPaths));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var merged = MergeFilesToBytes((IEnumerable<string>)inputPaths);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        OfficeFileCommit.WriteAllBytes(fullOutputPath, merged);
    }

    /// <summary>
    /// Merges PDFs from file paths and writes the result to the output path.
    /// </summary>
    public static void MergeFiles(IEnumerable<string> inputPaths, string outputPath) {
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var merged = MergeFilesToBytes(inputPaths);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        OfficeFileCommit.WriteAllBytes(fullOutputPath, merged);
    }

    /// <summary>
    /// Merges PDFs from file paths and writes the result to the output path, applying optional source preparation first.
    /// </summary>
    public static void MergeFiles(PdfMergeOptions options, IEnumerable<string> inputPaths, string outputPath) {
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var merged = MergeFilesToBytes(options, inputPaths);
        var directory = Path.GetDirectoryName(fullOutputPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        OfficeFileCommit.WriteAllBytes(fullOutputPath, merged);
    }

    /// <summary>
    /// Merges PDFs from file paths and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void MergeFiles(IEnumerable<string> inputPaths, Stream outputStream) {
        Guard.NotNull(outputStream, nameof(outputStream));
        if (!outputStream.CanWrite) {
            throw new ArgumentException("Stream must be writable.", nameof(outputStream));
        }

        WriteOutput(outputStream, MergeFilesToBytes(inputPaths));
    }

    /// <summary>
    /// Merges PDFs from file paths and returns the merged PDF bytes.
    /// </summary>
    public static byte[] MergeFilesToBytes(params string[] inputPaths) {
        Guard.NotNull(inputPaths, nameof(inputPaths));
        return MergeFilesToBytes((IEnumerable<string>)inputPaths);
    }

    /// <summary>
    /// Merges PDFs from file paths and returns the merged PDF bytes.
    /// </summary>
    public static byte[] MergeFilesToBytes(IEnumerable<string> inputPaths) {
        Guard.NotNull(inputPaths, nameof(inputPaths));

        var paths = inputPaths.ToArray();
        if (paths.Length == 0) {
            throw new ArgumentException("At least one input path must be supplied.", nameof(inputPaths));
        }

        var pdfs = new byte[paths.Length][];
        for (int i = 0; i < paths.Length; i++) {
            string inputPath = paths[i];
            if (inputPath is null) {
                throw new ArgumentException("Input path " + i.ToString(CultureInfo.InvariantCulture) + " cannot be null.", nameof(inputPaths));
            }

            if (string.IsNullOrWhiteSpace(inputPath)) {
                throw new ArgumentException("Input path " + i.ToString(CultureInfo.InvariantCulture) + " cannot be empty or whitespace.", nameof(inputPaths));
            }

            pdfs[i] = File.ReadAllBytes(inputPath);
        }

        return Merge(pdfs);
    }

    /// <summary>
    /// Merges PDFs from file paths and returns the merged PDF bytes, applying optional source preparation first.
    /// </summary>
    public static byte[] MergeFilesToBytes(PdfMergeOptions options, IEnumerable<string> inputPaths) {
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(inputPaths, nameof(inputPaths));

        var paths = inputPaths.ToArray();
        if (paths.Length == 0) {
            throw new ArgumentException("At least one input path must be supplied.", nameof(inputPaths));
        }

        var pdfs = new byte[paths.Length][];
        for (int i = 0; i < paths.Length; i++) {
            string inputPath = paths[i];
            if (inputPath is null) {
                throw new ArgumentException("Input path " + i.ToString(CultureInfo.InvariantCulture) + " cannot be null.", nameof(inputPaths));
            }

            if (string.IsNullOrWhiteSpace(inputPath)) {
                throw new ArgumentException("Input path " + i.ToString(CultureInfo.InvariantCulture) + " cannot be empty or whitespace.", nameof(inputPaths));
            }

            pdfs[i] = File.ReadAllBytes(inputPath);
        }

        return Merge(options, (IEnumerable<byte[]>)pdfs);
    }

    private static byte[] PrepareMergeSource(
        byte[] source,
        PdfMergeOptions? options,
        PdfLoadOptions? readOptions,
        out PdfLoadOptions preparedReadOptions) {
        preparedReadOptions = PdfLoadOptions.Resolve(readOptions);
        if (options is null) {
            return source;
        }

        if (options.FlattenVisualAnnotations) {
            byte[] input = source;
            source = PdfAnnotationFlattener.FlattenVisualAnnotations(input, options: null, preparedReadOptions, out PdfGeneratedOutputGrowth growth);
            preparedReadOptions = PdfLoadOptions.ForGeneratedOutput(preparedReadOptions, input, source, growth);
        }

        if (options.ResizePages is not null) {
            byte[] input = source;
            source = PdfPageEditor.ResizePages(input, options.ResizePages, preparedReadOptions);
            preparedReadOptions = PdfLoadOptions.ForGeneratedOutput(preparedReadOptions, input, source);
        }

        return source;
    }

    private static void WriteOutput(Stream outputStream, byte[] bytes) {
        Guard.NotNull(outputStream, nameof(outputStream));
        if (!outputStream.CanWrite) {
            throw new ArgumentException("Stream must be writable.", nameof(outputStream));
        }

        outputStream.Write(bytes, 0, bytes.Length);
    }

    private static string ValidateOutputPath(string outputPath) {
        Guard.NotNull(outputPath, nameof(outputPath));
        if (string.IsNullOrWhiteSpace(outputPath)) {
            throw new ArgumentException("Output path cannot be empty or whitespace.", nameof(outputPath));
        }

        string fullPath;
        try {
            fullPath = Path.GetFullPath(outputPath);
        } catch (Exception ex) {
            throw new ArgumentException("Output path is invalid.", nameof(outputPath), ex);
        }

        if (Directory.Exists(fullPath) && (File.GetAttributes(fullPath) & FileAttributes.Directory) == FileAttributes.Directory) {
            throw new ArgumentException("Output path refers to a directory; a file path is required.", nameof(outputPath));
        }

        var fileName = Path.GetFileName(fullPath);
        if (string.IsNullOrEmpty(fileName)) {
            throw new ArgumentException("Output path must include a file name.", nameof(outputPath));
        }

        if (fileName.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0) {
            throw new ArgumentException("Output path contains invalid file name characters.", nameof(outputPath));
        }

        return fullPath;
    }

    private static ImportedSource ImportSource(
        byte[] source,
        int sourceIndex,
        int[]? knownPageObjectNumbers,
        int mergedPageOffset,
        IReadOnlyDictionary<int, int>? outputPageIndexByPageObjectNumber,
        PdfLoadOptions? readOptions = null,
        PdfReadDocument? plannedDocument = null,
        PdfDocumentSecurityInfo? sourceSecurity = null,
        PdfPermissionPolicy? sourcePermissionPolicy = null,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfReadDocument document;
        if (plannedDocument is null) {
            (_, document) = PdfMutationPlanner.RequireFullRewriteDocument(
                source,
                PdfMutationOperation.MergeDocuments,
                readOptions,
                cancellationToken: cancellationToken);
        } else {
            document = plannedDocument;
        }

        int[] pageObjectNumbers = knownPageObjectNumbers ?? document.Pages.Select(page => page.ObjectNumber).ToArray();
        bool partialPageSelection = pageObjectNumbers.Length != document.Pages.Count;
        Dictionary<int, PdfIndirectObject> objects = partialPageSelection
            ? new Dictionary<int, PdfIndirectObject>(document.Objects)
            : document.Objects;
        string trailerRaw = document.TrailerRaw;
        if (document.Pages.Count == 0) {
            throw new ArgumentException("PDF input " + sourceIndex.ToString(CultureInfo.InvariantCulture) + " does not contain any pages.", nameof(source));
        }

        int formFieldCount = 0;
        int[]? selectedFormFieldRootObjectNumbers = partialPageSelection
            ? PrepareSelectedAcroFormFieldRoots(objects, document, pageObjectNumbers, out formFieldCount)
            : null;
        var copiedPageObjectIds = new HashSet<int>(pageObjectNumbers);
        var catalogState = PdfPageExtractor.PruneCatalogStateForPages(
            objects,
            PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw),
            copiedPageObjectIds,
            pageObjectNumbers,
            mergedPageOffset,
            outputPageIndexByPageObjectNumber);
        Dictionary<int, Dictionary<string, PdfObject>>? pageOverrides =
            PdfPageExtractor.BuildPageOverridesWithFilteredDestinationLinks(
                objects,
                pageObjectNumbers,
                pageOverrides: null,
                catalogState,
                copiedPageObjectIds);
        var collector = new PdfPageExtractor.ObjectCollector(objects, pageOverrides, cancellationToken);
        foreach (int pageObjectNumber in pageObjectNumbers) {
            cancellationToken.ThrowIfCancellationRequested();
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
        int[] formFieldRootObjectNumbers;
        if (selectedFormFieldRootObjectNumbers is null) {
            formFieldRootObjectNumbers = CollectAcroFormFieldRoots(objects, document, collector);
            formFieldCount = document.FormFields.Count;
        } else {
            formFieldRootObjectNumbers = selectedFormFieldRootObjectNumbers;
            if (formFieldRootObjectNumbers.Length > 0) {
                CollectAcroFormResources(objects, document, collector);
                foreach (int rootObjectNumber in formFieldRootObjectNumbers) {
                    int generation = objects.TryGetValue(rootObjectNumber, out PdfIndirectObject? rootObject) ? rootObject.Generation : 0;
                    collector.CollectObjectGraph(new PdfReference(rootObjectNumber, generation));
                }
            }
        }
        return new ImportedSource(
            objects,
            document,
            pageObjectNumbers,
            collector,
            pageOverrides,
            catalogState,
            formFieldRootObjectNumbers,
            formFieldCount,
            sourceSecurity ?? document.Security,
            sourcePermissionPolicy ?? document.ReadOptions.PermissionPolicy);
    }

    private static byte[] WriteMerged(
        IReadOnlyList<ImportedSource> sources,
        int primarySourceIndex,
        IReadOnlyList<OutputPageReference>? outputOrder,
        out int outputObjectCount,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        var objects = new List<byte[]>();
        var allPageObjectIds = new List<int>();
        var plans = new List<SourceWritePlan>(sources.Count);
        int nextObjectId = 1;

        foreach (var source in sources) {
            cancellationToken.ThrowIfCancellationRequested();
            var numberMap = new Dictionary<int, int>();
            foreach (int sourceId in source.Collector.ObjectIds) {
                numberMap[sourceId] = nextObjectId++;
            }

            source.OutputNumberMap = numberMap;
            plans.Add(new SourceWritePlan(source, numberMap));
        }

        if (outputOrder is null) {
            foreach (var plan in plans) {
                cancellationToken.ThrowIfCancellationRequested();
                foreach (int pageObjectNumber in plan.Source.PageObjectNumbers) {
                    allPageObjectIds.Add(plan.NumberMap[pageObjectNumber]);
                }
            }
        } else {
            foreach (var page in outputOrder) {
                cancellationToken.ThrowIfCancellationRequested();
                allPageObjectIds.Add(plans[page.SourceIndex].NumberMap[page.PageObjectNumber]);
            }
        }

        int pagesId = nextObjectId++;
        int catalogId = nextObjectId++;
        int infoId = nextObjectId;

        foreach (var plan in plans) {
            cancellationToken.ThrowIfCancellationRequested();
            var source = plan.Source;
            var context = new PdfPageExtractor.SerializationContext(plan.NumberMap, pagesId, source.Collector.MaterializedPageValues, source.Objects, source.PageOverrides);
            foreach (int sourceId in source.Collector.ObjectIds) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!source.Objects.TryGetValue(sourceId, out var sourceObject)) {
                    throw new InvalidOperationException("PDF object " + sourceId.ToString(CultureInfo.InvariantCulture) + " was referenced but not found.");
                }

                int newId = plan.NumberMap[sourceId];
                byte[] body = sourceObject.Value is PdfDictionary dictionary && source.Collector.PageObjectIds.Contains(sourceId)
                    ? PdfPageExtractor.SerializePageDictionary(dictionary, sourceId, context)
                    : PdfPageExtractor.SerializeObject(sourceObject.Value, context);

                objects.Add(PdfPageExtractor.WrapObject(newId, body));
            }
        }

        objects.Add(PdfPageExtractor.WrapObject(pagesId, PdfEncoding.Latin1GetBytes(PdfPageTreeBuilder.BuildPagesDictionary(allPageObjectIds))));
        var primaryPlan = plans[primarySourceIndex];
        var primaryCatalogContext = new PdfPageExtractor.SerializationContext(primaryPlan.NumberMap, pagesId, primaryPlan.Source.Collector.MaterializedPageValues, primaryPlan.Source.Objects);
        objects.Add(PdfPageExtractor.WrapObject(catalogId, PdfEncoding.Latin1GetBytes(PdfPageExtractor.BuildCatalogDictionary(pagesId, sources[primarySourceIndex].CatalogState, primaryCatalogContext))));
        objects.Add(PdfPageExtractor.WrapObject(infoId, PdfEncoding.Latin1GetBytes(PdfPageExtractor.BuildInfoDictionary(BuildMergedMetadata(sources, primarySourceIndex)))));

        outputObjectCount = objects.Count;
        return PdfPageExtractor.Assemble(objects, catalogId, infoId, cancellationToken: cancellationToken);
    }

    private static PdfMetadata BuildMergedMetadata(IReadOnlyList<ImportedSource> sources, int primarySourceIndex) {
        var primary = sources[primarySourceIndex].Metadata;
        return new PdfMetadata {
            Title = string.IsNullOrEmpty(primary.Title) ? "Merged PDF" : primary.Title,
            Author = primary.Author,
            Subject = primary.Subject,
            Keywords = primary.Keywords
        };
    }

    private sealed class ImportedSource {
        public ImportedSource(
            Dictionary<int, PdfIndirectObject> objects,
            PdfReadDocument document,
            int[] pageObjectNumbers,
            PdfPageExtractor.ObjectCollector collector,
            Dictionary<int, Dictionary<string, PdfObject>>? pageOverrides,
            PdfPageExtractor.CatalogRewriteState catalogState,
            int[] formFieldRootObjectNumbers,
            int formFieldCount,
            PdfDocumentSecurityInfo sourceSecurity,
            PdfPermissionPolicy sourcePermissionPolicy) {
            Objects = objects;
            Document = document;
            PageObjectNumbers = pageObjectNumbers;
            var pageNumberByObjectNumber = document.Pages
                .Select((page, index) => new { page.ObjectNumber, PageNumber = index + 1 })
                .ToDictionary(static item => item.ObjectNumber, static item => item.PageNumber);
            SelectedPageNumbers = pageObjectNumbers.Select(pageObjectNumber => pageNumberByObjectNumber[pageObjectNumber]).ToArray();
            var selectedPageNumberSet = new HashSet<int>(SelectedPageNumbers);
            NamedDestinationCount = document.NamedDestinations.Count(destination =>
                destination.PageNumber.HasValue && selectedPageNumberSet.Contains(destination.PageNumber.Value));
            OutlineCount = catalogState.Outlines is null ? 0 : CountOutlines(document.Outlines);
            PageLabelCount = CountPageLabelRules(objects, catalogState.PageLabels);
            Collector = collector;
            PageOverrides = pageOverrides;
            CatalogState = catalogState;
            FormFieldRootObjectNumbers = formFieldRootObjectNumbers;
            FormFieldCount = formFieldCount;
            SourceSecurity = sourceSecurity;
            SourcePermissionPolicy = sourcePermissionPolicy;
        }

        public Dictionary<int, PdfIndirectObject> Objects { get; }

        public PdfReadDocument Document { get; }

        public PdfMetadata Metadata => Document.UncheckedMetadata;

        public int[] PageObjectNumbers { get; }

        public int[] SelectedPageNumbers { get; }

        public int NamedDestinationCount { get; }

        public int OutlineCount { get; }

        public int PageLabelCount { get; }

        public PdfPageExtractor.ObjectCollector Collector { get; }

        public Dictionary<int, Dictionary<string, PdfObject>>? PageOverrides { get; }

        public PdfPageExtractor.CatalogRewriteState CatalogState { get; }

        public int[] FormFieldRootObjectNumbers { get; }

        public int FormFieldCount { get; }

        public PdfDocumentSecurityInfo SourceSecurity { get; }

        public PdfPermissionPolicy SourcePermissionPolicy { get; }

        public IReadOnlyDictionary<int, int>? OutputNumberMap { get; set; }

        private static int CountPageLabelRules(Dictionary<int, PdfIndirectObject> objects, PdfObject? pageLabels) {
            if (PdfObjectLookup.Resolve(objects, pageLabels) is not PdfDictionary dictionary ||
                !dictionary.Items.TryGetValue("Nums", out PdfObject? numsObject) ||
                PdfObjectLookup.Resolve(objects, numsObject) is not PdfArray nums) {
                return 0;
            }

            return nums.Items.Count / 2;
        }
    }

    private sealed class SourceWritePlan {
        public SourceWritePlan(ImportedSource source, Dictionary<int, int> numberMap) {
            Source = source;
            NumberMap = numberMap;
        }

        public ImportedSource Source { get; }

        public Dictionary<int, int> NumberMap { get; }
    }

    private readonly struct OutputPageReference {
        public OutputPageReference(int sourceIndex, int pageObjectNumber) {
            SourceIndex = sourceIndex;
            PageObjectNumber = pageObjectNumber;
        }

        public int SourceIndex { get; }

        public int PageObjectNumber { get; }
    }
}
