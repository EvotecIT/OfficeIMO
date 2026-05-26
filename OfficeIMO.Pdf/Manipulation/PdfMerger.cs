using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Provides first-party PDF merge helpers for PDFs that can be parsed by OfficeIMO.Pdf.
/// </summary>
public static class PdfMerger {
    /// <summary>
    /// Merges all pages from the supplied PDFs into one new PDF.
    /// </summary>
    public static byte[] Merge(params byte[][] pdfs) {
        return Merge((IEnumerable<byte[]>)pdfs);
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
        return MergeCore(pdfs, primarySourceIndex: 0);
    }

    internal static byte[] MergeWithPrimarySource(int primarySourceIndex, params byte[][] pdfs) {
        return MergeCore(pdfs, primarySourceIndex);
    }

    internal static byte[] MergePrimaryWithInsertedPages(byte[] primaryPdf, byte[] insertedPdf, int insertBeforePageNumber) {
        Guard.NotNull(primaryPdf, nameof(primaryPdf));
        Guard.NotNull(insertedPdf, nameof(insertedPdf));

        PdfSyntax.ThrowIfUnsafeForRewrite(primaryPdf);
        PdfSyntax.ThrowIfUnsafeForRewrite(insertedPdf);

        var primaryDocument = PdfReadDocument.Load(primaryPdf);
        if (primaryDocument.Pages.Count == 0) {
            throw new ArgumentException("Primary PDF does not contain any pages.", nameof(primaryPdf));
        }

        if (insertBeforePageNumber < 1 || insertBeforePageNumber > primaryDocument.Pages.Count + 1) {
            throw new ArgumentOutOfRangeException(nameof(insertBeforePageNumber), "Insert-before page must be in the primary document page range.");
        }

        var insertedDocument = PdfReadDocument.Load(insertedPdf);
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
            ImportSource(primaryPdf, 0, primaryPageObjectNumbers, 0, primaryPageIndexMap),
            ImportSource(insertedPdf, 1, insertedPageObjectNumbers, insertBeforePageNumber - 1, null)
        };
        return WriteMerged(importedSources, primarySourceIndex: 0, outputOrder);
    }

    private static byte[] MergeCore(IEnumerable<byte[]> pdfs, int primarySourceIndex) {
        Guard.NotNull(pdfs, nameof(pdfs));

        var sources = pdfs.ToArray();
        if (sources.Length == 0) {
            throw new ArgumentException("At least one PDF must be supplied.", nameof(pdfs));
        }

        if (primarySourceIndex < 0 || primarySourceIndex >= sources.Length) {
            throw new ArgumentOutOfRangeException(nameof(primarySourceIndex), "Primary source index must refer to one of the supplied PDFs.");
        }

        var importedSources = new List<ImportedSource>(sources.Length);
        int mergedPageOffset = 0;
        for (int i = 0; i < sources.Length; i++) {
            byte[] source = sources[i];
            if (source is null) {
                throw new ArgumentException("PDF input " + i.ToString(CultureInfo.InvariantCulture) + " cannot be null.", nameof(pdfs));
            }

            importedSources.Add(ImportSource(source, i, null, mergedPageOffset, null));
            mergedPageOffset += importedSources[importedSources.Count - 1].PageObjectNumbers.Length;
        }

        return WriteMerged(importedSources, primarySourceIndex);
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
    /// Merges all pages from the supplied PDFs and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void Merge(IEnumerable<byte[]> pdfs, Stream outputStream) {
        WriteOutput(outputStream, Merge(pdfs));
    }

    /// <summary>
    /// Merges all pages from the supplied readable PDF streams and writes the result to <paramref name="outputStream"/>.
    /// </summary>
    public static void Merge(IEnumerable<Stream> streams, Stream outputStream) {
        WriteOutput(outputStream, Merge(streams));
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
        File.WriteAllBytes(fullOutputPath, merged);
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
        File.WriteAllBytes(fullOutputPath, merged);
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
        IReadOnlyDictionary<int, int>? outputPageIndexByPageObjectNumber) {
        PdfSyntax.ThrowIfUnsafeForRewrite(source);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(source);
        var document = PdfReadDocument.Load(source);
        if (document.Pages.Count == 0) {
            throw new ArgumentException("PDF input " + sourceIndex.ToString(CultureInfo.InvariantCulture) + " does not contain any pages.", nameof(source));
        }

        var collector = new PdfPageExtractor.ObjectCollector(objects);
        int[] pageObjectNumbers = knownPageObjectNumbers ?? document.Pages.Select(page => page.ObjectNumber).ToArray();
        foreach (int pageObjectNumber in pageObjectNumbers) {
            collector.CollectPage(pageObjectNumber);
        }

        var catalogState = PdfPageExtractor.PruneCatalogStateForPages(
            objects,
            PdfPageExtractor.ExtractCatalogRewriteState(objects, trailerRaw),
            collector.PageObjectIds,
            pageObjectNumbers,
            mergedPageOffset,
            outputPageIndexByPageObjectNumber);
        collector.CollectObjectGraph(catalogState.Outlines);
        collector.CollectObjectGraph(catalogState.PageLabels);
        collector.CollectObjectGraph(catalogState.NamedDestinationNameTree);
        collector.CollectObjectGraph(catalogState.XmpMetadata);
        collector.CollectObjectGraph(catalogState.CatalogUri);
        collector.CollectObjectGraph(catalogState.OutputIntents);
        collector.CollectObjectGraph(catalogState.EmbeddedFiles);
        collector.CollectObjectGraph(catalogState.AssociatedFiles);
        collector.CollectObjectGraph(catalogState.OptionalContent);
        return new ImportedSource(objects, document.Metadata, pageObjectNumbers, collector, catalogState);
    }

    private static byte[] WriteMerged(
        IReadOnlyList<ImportedSource> sources,
        int primarySourceIndex,
        IReadOnlyList<OutputPageReference>? outputOrder = null) {
        var objects = new List<byte[]>();
        var allPageObjectIds = new List<int>();
        var plans = new List<SourceWritePlan>(sources.Count);
        int nextObjectId = 1;

        foreach (var source in sources) {
            var numberMap = new Dictionary<int, int>();
            foreach (int sourceId in source.Collector.ObjectIds) {
                numberMap[sourceId] = nextObjectId++;
            }

            plans.Add(new SourceWritePlan(source, numberMap));
        }

        if (outputOrder is null) {
            foreach (var plan in plans) {
                foreach (int pageObjectNumber in plan.Source.PageObjectNumbers) {
                    allPageObjectIds.Add(plan.NumberMap[pageObjectNumber]);
                }
            }
        } else {
            foreach (var page in outputOrder) {
                allPageObjectIds.Add(plans[page.SourceIndex].NumberMap[page.PageObjectNumber]);
            }
        }

        int pagesId = nextObjectId++;
        int catalogId = nextObjectId++;
        int infoId = nextObjectId;

        foreach (var plan in plans) {
            var source = plan.Source;
            var context = new PdfPageExtractor.SerializationContext(plan.NumberMap, pagesId, source.Collector.MaterializedPageValues);
            foreach (int sourceId in source.Collector.ObjectIds) {
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
        var primaryCatalogContext = new PdfPageExtractor.SerializationContext(primaryPlan.NumberMap, pagesId, primaryPlan.Source.Collector.MaterializedPageValues);
        objects.Add(PdfPageExtractor.WrapObject(catalogId, PdfEncoding.Latin1GetBytes(PdfPageExtractor.BuildCatalogDictionary(pagesId, sources[primarySourceIndex].CatalogState, primaryCatalogContext))));
        objects.Add(PdfPageExtractor.WrapObject(infoId, PdfEncoding.Latin1GetBytes(PdfPageExtractor.BuildInfoDictionary(BuildMergedMetadata(sources, primarySourceIndex)))));

        return PdfPageExtractor.Assemble(objects, catalogId, infoId);
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
            PdfMetadata metadata,
            int[] pageObjectNumbers,
            PdfPageExtractor.ObjectCollector collector,
            PdfPageExtractor.CatalogRewriteState catalogState) {
            Objects = objects;
            Metadata = metadata;
            PageObjectNumbers = pageObjectNumbers;
            Collector = collector;
            CatalogState = catalogState;
        }

        public Dictionary<int, PdfIndirectObject> Objects { get; }

        public PdfMetadata Metadata { get; }

        public int[] PageObjectNumbers { get; }

        public PdfPageExtractor.ObjectCollector Collector { get; }

        public PdfPageExtractor.CatalogRewriteState CatalogState { get; }
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
