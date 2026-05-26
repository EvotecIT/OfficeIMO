using System.Globalization;

namespace OfficeIMO.Pdf;

/// <summary>
/// Provides first-party page extraction and simple splitting for PDFs that can be parsed by OfficeIMO.Pdf.
/// </summary>
public static class PdfPageExtractor {
    /// <summary>
    /// Creates a new PDF containing the selected one-based page numbers in the requested order.
    /// </summary>
    public static byte[] ExtractPages(byte[] pdf, params int[] pageNumbers) {
        return ExtractPages(pdf, (IEnumerable<int>)pageNumbers);
    }

    /// <summary>
    /// Creates a new PDF containing the selected one-based page numbers in the requested order from the current position of a readable stream.
    /// </summary>
    public static byte[] ExtractPages(Stream stream, params int[] pageNumbers) {
        return ExtractPages(stream, (IEnumerable<int>)pageNumbers);
    }

    /// <summary>
    /// Creates a new PDF containing the selected one-based page numbers in the requested order.
    /// </summary>
    public static byte[] ExtractPages(byte[] pdf, IEnumerable<int> pageNumbers) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);

        var selected = pageNumbers.ToArray();
        if (selected.Length == 0) {
            throw new ArgumentException("At least one page number must be specified.", nameof(pageNumbers));
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Load(pdf);
        ValidatePageNumbers(selected, document.Pages.Count, nameof(pageNumbers));

        var pageObjectNumbers = selected.Select(pageNumber => document.Pages[pageNumber - 1].ObjectNumber).ToArray();
        return ExtractPages(objects, document.Metadata, pageObjectNumbers, catalogState: ExtractCatalogRewriteState(objects, trailerRaw));
    }

    /// <summary>
    /// Creates a new PDF containing the selected one-based page numbers in the requested order from the current position of a readable stream.
    /// </summary>
    public static byte[] ExtractPages(Stream stream, IEnumerable<int> pageNumbers) {
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        return ExtractPages(ReadStream(stream, nameof(stream)), pageNumbers);
    }

    /// <summary>
    /// Writes a new PDF containing the selected one-based page numbers to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPages(byte[] pdf, Stream outputStream, params int[] pageNumbers) {
        WriteOutput(outputStream, ExtractPages(pdf, pageNumbers));
    }

    /// <summary>
    /// Writes a new PDF containing the selected one-based page numbers from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPages(Stream inputStream, Stream outputStream, params int[] pageNumbers) {
        WriteOutput(outputStream, ExtractPages(inputStream, pageNumbers));
    }

    /// <summary>
    /// Creates a new PDF containing the inclusive one-based page range.
    /// </summary>
    public static byte[] ExtractPageRange(byte[] pdf, int firstPage, int lastPage) {
        if (firstPage > lastPage) {
            throw new ArgumentOutOfRangeException(nameof(lastPage), "Last page must be greater than or equal to first page.");
        }

        return ExtractPages(pdf, Enumerable.Range(firstPage, lastPage - firstPage + 1));
    }

    /// <summary>
    /// Creates a new PDF containing the inclusive one-based page range.
    /// </summary>
    public static byte[] ExtractPageRange(byte[] pdf, PdfPageRange pageRange) {
        return ExtractPages(pdf, pageRange.ToPageNumbers());
    }

    /// <summary>
    /// Creates a new PDF containing the inclusive one-based page range from the current position of a readable stream.
    /// </summary>
    public static byte[] ExtractPageRange(Stream stream, int firstPage, int lastPage) {
        return ExtractPageRange(ReadStream(stream, nameof(stream)), firstPage, lastPage);
    }

    /// <summary>
    /// Creates a new PDF containing the inclusive one-based page range from the current position of a readable stream.
    /// </summary>
    public static byte[] ExtractPageRange(Stream stream, PdfPageRange pageRange) {
        return ExtractPageRange(ReadStream(stream, nameof(stream)), pageRange);
    }

    /// <summary>
    /// Writes a new PDF containing the inclusive one-based page range to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRange(byte[] pdf, Stream outputStream, int firstPage, int lastPage) {
        WriteOutput(outputStream, ExtractPageRange(pdf, firstPage, lastPage));
    }

    /// <summary>
    /// Writes a new PDF containing the inclusive one-based page range to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRange(byte[] pdf, Stream outputStream, PdfPageRange pageRange) {
        WriteOutput(outputStream, ExtractPageRange(pdf, pageRange));
    }

    /// <summary>
    /// Writes a new PDF containing the inclusive one-based page range from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRange(Stream inputStream, Stream outputStream, int firstPage, int lastPage) {
        WriteOutput(outputStream, ExtractPageRange(inputStream, firstPage, lastPage));
    }

    /// <summary>
    /// Writes a new PDF containing the inclusive one-based page range from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRange(Stream inputStream, Stream outputStream, PdfPageRange pageRange) {
        WriteOutput(outputStream, ExtractPageRange(inputStream, pageRange));
    }

    /// <summary>
    /// Writes a new PDF containing the inclusive one-based page range.
    /// </summary>
    public static void ExtractPageRange(string inputPath, string outputPath, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = ExtractPageRange(File.ReadAllBytes(inputPath), firstPage, lastPage);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF containing the inclusive one-based page range from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRange(string inputPath, Stream outputStream, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        var bytes = ExtractPageRange(File.ReadAllBytes(inputPath), firstPage, lastPage);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Writes a new PDF containing the inclusive one-based page range.
    /// </summary>
    public static void ExtractPageRange(string inputPath, string outputPath, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = ExtractPageRange(File.ReadAllBytes(inputPath), pageRange);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF containing the inclusive one-based page range from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRange(string inputPath, Stream outputStream, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        var bytes = ExtractPageRange(File.ReadAllBytes(inputPath), pageRange);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Creates a new PDF containing the inclusive one-based page range from a file path.
    /// </summary>
    public static byte[] ExtractPageRange(string inputPath, int firstPage, int lastPage) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return ExtractPageRange(File.ReadAllBytes(inputPath), firstPage, lastPage);
    }

    /// <summary>
    /// Creates a new PDF containing the inclusive one-based page range from a file path.
    /// </summary>
    public static byte[] ExtractPageRange(string inputPath, PdfPageRange pageRange) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return ExtractPageRange(File.ReadAllBytes(inputPath), pageRange);
    }

    /// <summary>
    /// Creates a new PDF containing the supplied inclusive one-based page ranges in caller order.
    /// </summary>
    public static byte[] ExtractPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return ExtractPageRanges(pdf, (IEnumerable<PdfPageRange>)pageRanges);
    }

    /// <summary>
    /// Creates a new PDF containing the supplied inclusive one-based page ranges in caller order.
    /// </summary>
    public static byte[] ExtractPageRanges(byte[] pdf, IEnumerable<PdfPageRange> pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageRanges, nameof(pageRanges));
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);

        var ranges = pageRanges.ToArray();
        if (ranges.Length == 0) {
            throw new ArgumentException("At least one page range must be specified.", nameof(pageRanges));
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Load(pdf);
        ValidatePageRanges(ranges, document.Pages.Count, nameof(pageRanges));

        var pageObjectNumbers = new List<int>(ranges.Sum(range => range.PageCount));
        foreach (var range in ranges) {
            for (int pageNumber = range.FirstPage; pageNumber <= range.LastPage; pageNumber++) {
                pageObjectNumbers.Add(document.Pages[pageNumber - 1].ObjectNumber);
            }
        }

        return ExtractPages(objects, document.Metadata, pageObjectNumbers.ToArray(), catalogState: ExtractCatalogRewriteState(objects, trailerRaw));
    }

    /// <summary>
    /// Creates a new PDF containing the supplied inclusive one-based page ranges in caller order from the current position of a readable stream.
    /// </summary>
    public static byte[] ExtractPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return ExtractPageRanges(ReadStream(stream, nameof(stream)), pageRanges);
    }

    /// <summary>
    /// Creates a new PDF containing the supplied inclusive one-based page ranges in caller order from the current position of a readable stream.
    /// </summary>
    public static byte[] ExtractPageRanges(Stream stream, IEnumerable<PdfPageRange> pageRanges) {
        Guard.NotNull(pageRanges, nameof(pageRanges));
        return ExtractPageRanges(ReadStream(stream, nameof(stream)), pageRanges);
    }

    /// <summary>
    /// Writes a new PDF containing the supplied inclusive one-based page ranges in caller order to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRanges(byte[] pdf, Stream outputStream, params PdfPageRange[] pageRanges) {
        WriteOutput(outputStream, ExtractPageRanges(pdf, pageRanges));
    }

    /// <summary>
    /// Writes a new PDF containing the supplied inclusive one-based page ranges in caller order from the current position of a readable stream to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRanges(Stream inputStream, Stream outputStream, params PdfPageRange[] pageRanges) {
        WriteOutput(outputStream, ExtractPageRanges(inputStream, pageRanges));
    }

    /// <summary>
    /// Writes a new PDF containing the supplied inclusive one-based page ranges in caller order.
    /// </summary>
    public static void ExtractPageRanges(string inputPath, string outputPath, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = ExtractPageRanges(File.ReadAllBytes(inputPath), pageRanges);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF containing the supplied inclusive one-based page ranges from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPageRanges(string inputPath, Stream outputStream, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        var bytes = ExtractPageRanges(File.ReadAllBytes(inputPath), pageRanges);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Creates a new PDF containing the supplied inclusive one-based page ranges in caller order from a file path.
    /// </summary>
    public static byte[] ExtractPageRanges(string inputPath, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return ExtractPageRanges(File.ReadAllBytes(inputPath), pageRanges);
    }

    /// <summary>
    /// Splits a PDF into one single-page PDF per source page.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPages(byte[] pdf) {
        Guard.NotNull(pdf, nameof(pdf));
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Load(pdf);
        var catalogState = ExtractCatalogRewriteState(objects, trailerRaw);
        var result = new List<byte[]>(document.Pages.Count);

        foreach (var page in document.Pages) {
            result.Add(ExtractPages(objects, document.Metadata, new[] { page.ObjectNumber }, catalogState: catalogState));
        }

        return result;
    }

    /// <summary>
    /// Splits a PDF into one single-page PDF per source page from the current position of a readable stream.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPages(Stream stream) {
        return SplitPages(ReadStream(stream, nameof(stream)));
    }

    /// <summary>
    /// Splits a PDF into one output PDF per inclusive one-based page range.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPageRanges(byte[] pdf, params PdfPageRange[] pageRanges) {
        return SplitPageRanges(pdf, (IEnumerable<PdfPageRange>)pageRanges);
    }

    /// <summary>
    /// Splits a PDF into one output PDF per inclusive one-based page range.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPageRanges(byte[] pdf, IEnumerable<PdfPageRange> pageRanges) {
        Guard.NotNull(pdf, nameof(pdf));
        Guard.NotNull(pageRanges, nameof(pageRanges));
        PdfSyntax.ThrowIfUnsafeForRewrite(pdf);

        var ranges = pageRanges.ToArray();
        if (ranges.Length == 0) {
            throw new ArgumentException("At least one page range must be specified.", nameof(pageRanges));
        }

        var (objects, trailerRaw) = PdfSyntax.ParseObjects(pdf);
        var document = PdfReadDocument.Load(pdf);
        ValidatePageRanges(ranges, document.Pages.Count, nameof(pageRanges));
        var catalogState = ExtractCatalogRewriteState(objects, trailerRaw);
        var result = new List<byte[]>(ranges.Length);

        foreach (var range in ranges) {
            int[] pageObjectNumbers = Enumerable
                .Range(range.FirstPage, range.PageCount)
                .Select(pageNumber => document.Pages[pageNumber - 1].ObjectNumber)
                .ToArray();
            result.Add(ExtractPages(objects, document.Metadata, pageObjectNumbers, catalogState: catalogState));
        }

        return result;
    }

    /// <summary>
    /// Splits a PDF from the current position of a readable stream into one output PDF per inclusive one-based page range.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPageRanges(Stream stream, params PdfPageRange[] pageRanges) {
        return SplitPageRanges(stream, (IEnumerable<PdfPageRange>)pageRanges);
    }

    /// <summary>
    /// Splits a PDF from the current position of a readable stream into one output PDF per inclusive one-based page range.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPageRanges(Stream stream, IEnumerable<PdfPageRange> pageRanges) {
        Guard.NotNull(pageRanges, nameof(pageRanges));
        return SplitPageRanges(ReadStream(stream, nameof(stream)), pageRanges);
    }

    /// <summary>
    /// Splits a PDF into one single-page PDF per source page and writes the files to <paramref name="outputDirectory"/>.
    /// </summary>
    public static IReadOnlyList<string> SplitPages(string inputPath, string outputDirectory) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = SplitPages(File.ReadAllBytes(inputPath));
        string baseName = Path.GetFileNameWithoutExtension(inputPath);
        return WriteSplitPages(pages, fullOutputDirectory, baseName);
    }

    /// <summary>
    /// Splits a PDF from a file path into one single-page PDF per source page.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPages(string inputPath) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return SplitPages(File.ReadAllBytes(inputPath));
    }

    /// <summary>
    /// Splits a PDF from a file path into one output PDF per inclusive one-based page range.
    /// </summary>
    public static IReadOnlyList<byte[]> SplitPageRanges(string inputPath, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        var ranges = ValidatePageRangeArguments(pageRanges, nameof(pageRanges));
        return SplitPageRanges(File.ReadAllBytes(inputPath), ranges);
    }

    /// <summary>
    /// Splits a PDF from the current position of a readable stream and writes one single-page PDF per source page.
    /// </summary>
    public static IReadOnlyList<string> SplitPages(Stream stream, string outputDirectory, string baseName = "page") {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = SplitPages(stream);
        return WriteSplitPages(pages, fullOutputDirectory, baseName);
    }

    /// <summary>
    /// Splits a PDF from a file path into one output PDF per inclusive one-based page range and writes the files to <paramref name="outputDirectory"/>.
    /// </summary>
    public static IReadOnlyList<string> SplitPageRanges(string inputPath, string outputDirectory, params PdfPageRange[] pageRanges) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        var ranges = ValidatePageRangeArguments(pageRanges, nameof(pageRanges));
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = SplitPageRanges(File.ReadAllBytes(inputPath), ranges);
        string baseName = Path.GetFileNameWithoutExtension(inputPath);
        return WriteSplitPageRanges(pages, fullOutputDirectory, baseName, ranges);
    }

    /// <summary>
    /// Splits a PDF from the current position of a readable stream into one output PDF per inclusive one-based page range and writes the files to <paramref name="outputDirectory"/>.
    /// </summary>
    public static IReadOnlyList<string> SplitPageRanges(Stream stream, string outputDirectory, string baseName, params PdfPageRange[] pageRanges) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));

        var ranges = ValidatePageRangeArguments(pageRanges, nameof(pageRanges));
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);
        var pages = SplitPageRanges(stream, ranges);
        return WriteSplitPageRanges(pages, fullOutputDirectory, baseName, ranges);
    }

    /// <summary>
    /// Writes a new PDF containing the selected one-based page numbers.
    /// </summary>
    public static void ExtractPages(string inputPath, string outputPath, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        Guard.NotNull(outputPath, nameof(outputPath));

        string fullOutputPath = ValidateOutputPath(outputPath);
        var bytes = ExtractPages(File.ReadAllBytes(inputPath), pageNumbers);
        WriteOutput(fullOutputPath, bytes);
    }

    /// <summary>
    /// Writes a new PDF containing the selected one-based page numbers from a file path to <paramref name="outputStream"/>.
    /// </summary>
    public static void ExtractPages(string inputPath, Stream outputStream, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        ValidateWritableOutputStream(outputStream);

        var bytes = ExtractPages(File.ReadAllBytes(inputPath), pageNumbers);
        WriteOutput(outputStream, bytes);
    }

    /// <summary>
    /// Creates a new PDF containing the selected one-based page numbers from a file path.
    /// </summary>
    public static byte[] ExtractPages(string inputPath, params int[] pageNumbers) {
        Guard.NotNullOrWhiteSpace(inputPath, nameof(inputPath));
        return ExtractPages(File.ReadAllBytes(inputPath), pageNumbers);
    }

    private static void WriteOutput(Stream outputStream, byte[] bytes) {
        ValidateWritableOutputStream(outputStream);

        outputStream.Write(bytes, 0, bytes.Length);
    }

    private static void ValidateWritableOutputStream(Stream outputStream) {
        Guard.NotNull(outputStream, nameof(outputStream));
        if (!outputStream.CanWrite) {
            throw new ArgumentException("Stream must be writable.", nameof(outputStream));
        }
    }

    private static byte[] ReadStream(Stream stream, string paramName) {
        Guard.NotNull(stream, paramName);
        if (!stream.CanRead) {
            throw new ArgumentException("Stream must be readable.", paramName);
        }

        using var buffer = new MemoryStream();
        stream.CopyTo(buffer);
        return buffer.ToArray();
    }

    private static List<string> WriteSplitPages(IReadOnlyList<byte[]> pages, string outputDirectory, string? baseName) {
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);

        string safeBaseName = Path.GetFileNameWithoutExtension(baseName ?? string.Empty) ?? string.Empty;
        if (string.IsNullOrWhiteSpace(safeBaseName)) {
            safeBaseName = "page";
        }

        var paths = new List<string>(pages.Count);
        for (int i = 0; i < pages.Count; i++) {
            string outputPath = Path.Combine(fullOutputDirectory, safeBaseName + "-page-" + (i + 1).ToString("0000", CultureInfo.InvariantCulture) + ".pdf");
            File.WriteAllBytes(outputPath, pages[i]);
            paths.Add(outputPath);
        }

        return paths;
    }

    private static List<string> WriteSplitPageRanges(IReadOnlyList<byte[]> pages, string outputDirectory, string? baseName, PdfPageRange[] ranges) {
        string fullOutputDirectory = ValidateOutputDirectory(outputDirectory);

        string safeBaseName = Path.GetFileNameWithoutExtension(baseName ?? string.Empty) ?? string.Empty;
        if (string.IsNullOrWhiteSpace(safeBaseName)) {
            safeBaseName = "page";
        }

        var paths = new List<string>(pages.Count);
        var rangeOccurrences = new Dictionary<PdfPageRange, int>();
        for (int i = 0; i < pages.Count; i++) {
            var range = ranges[i];
            rangeOccurrences.TryGetValue(range, out int occurrence);
            occurrence++;
            rangeOccurrences[range] = occurrence;

            string outputPath = Path.Combine(
                fullOutputDirectory,
                safeBaseName + "-pages-" +
                range.FirstPage.ToString("0000", CultureInfo.InvariantCulture) + "-" +
                range.LastPage.ToString("0000", CultureInfo.InvariantCulture) +
                (occurrence <= 1 ? string.Empty : "-occurrence-" + occurrence.ToString("0000", CultureInfo.InvariantCulture)) +
                ".pdf");
            File.WriteAllBytes(outputPath, pages[i]);
            paths.Add(outputPath);
        }

        return paths;
    }

    private static string ValidateOutputDirectory(string outputDirectory) {
        Guard.NotNull(outputDirectory, nameof(outputDirectory));
        if (string.IsNullOrWhiteSpace(outputDirectory)) {
            throw new ArgumentException("Output directory cannot be empty or whitespace.", nameof(outputDirectory));
        }

        string fullOutputDirectory;
        try {
            fullOutputDirectory = Path.GetFullPath(outputDirectory);
        } catch (Exception ex) {
            throw new ArgumentException("Output directory is invalid.", nameof(outputDirectory), ex);
        }

        if (File.Exists(fullOutputDirectory)) {
            throw new ArgumentException("Output directory refers to a file; a directory path is required.", nameof(outputDirectory));
        }

        Directory.CreateDirectory(fullOutputDirectory);
        return fullOutputDirectory;
    }

    private static PdfPageRange[] ValidatePageRangeArguments(PdfPageRange[]? pageRanges, string paramName) {
        if (pageRanges is null) {
            throw new ArgumentNullException(paramName);
        }

        if (pageRanges.Length == 0) {
            throw new ArgumentException("At least one page range must be specified.", paramName);
        }

        return (PdfPageRange[])pageRanges.Clone();
    }

    internal static byte[] ExtractPages(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfMetadata metadata,
        int[] pageObjectNumbers,
        Dictionary<int, Dictionary<string, PdfObject>>? pageOverrides = null,
        IEnumerable<AdditionalObject>? additionalObjects = null,
        CatalogRewriteState? catalogState = null) {
        catalogState ??= CatalogRewriteState.Empty;
        var copiedPageObjectIds = new HashSet<int>(pageObjectNumbers);
        catalogState = PruneCatalogStateForPages(sourceObjects, catalogState, copiedPageObjectIds, pageObjectNumbers);
        pageOverrides = BuildPageOverridesWithFilteredNamedDestinationLinks(sourceObjects, pageObjectNumbers, pageOverrides, catalogState);

        var collector = new ObjectCollector(sourceObjects, pageOverrides);
        foreach (int pageObjectNumber in pageObjectNumbers) {
            collector.CollectPage(pageObjectNumber);
        }

        collector.CollectObjectGraph(catalogState.Outlines);
        collector.CollectObjectGraph(catalogState.PageLabels);
        collector.CollectObjectGraph(catalogState.NamedDestinationNameTree);
        collector.CollectObjectGraph(catalogState.XmpMetadata);
        collector.CollectObjectGraph(catalogState.CatalogUri);
        collector.CollectObjectGraph(catalogState.OutputIntents);
        collector.CollectObjectGraph(catalogState.EmbeddedFiles);
        collector.CollectObjectGraph(catalogState.AssociatedFiles);
        collector.CollectObjectGraph(catalogState.OptionalContent);

        var sourceIds = collector.ObjectIds;
        var extraObjects = additionalObjects?.ToArray() ?? Array.Empty<AdditionalObject>();
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
            var clonedAnnotationState = BuildClonedAnnotationState(sourceObjects, pageObjectNumber, ref nextObjectId);
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

        return Assemble(objects, catalogId, infoId);
    }

    internal static CatalogRewriteState ExtractCatalogRewriteState(Dictionary<int, PdfIndirectObject> sourceObjects, string? trailerRaw = null) {
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
            return new CatalogRewriteState(pageMode, pageLayout, BuildCatalogVersion(sourceObjects, catalogVersion), BuildCatalogLanguage(sourceObjects, catalogLanguage), BuildOutlines(sourceObjects, outlines), pageLabels, namedDestinations, BuildNamedDestinationNameTree(sourceObjects, names), openAction, BuildViewerPreferences(sourceObjects, viewerPreferences), BuildXmpMetadata(sourceObjects, xmpMetadata), BuildCatalogUri(sourceObjects, catalogUri), BuildOutputIntents(sourceObjects, outputIntents), BuildEmbeddedFiles(sourceObjects, names), BuildAssociatedFiles(sourceObjects, associatedFiles), BuildOptionalContent(sourceObjects, optionalContent), GetPageObjectNumbersInDocumentOrder(sourceObjects, dictionary));
        }

        return CatalogRewriteState.Empty;
    }

    internal static CatalogRewriteState PruneCatalogStateForPages(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        CatalogRewriteState catalogState,
        HashSet<int> copiedPageObjectIds,
        IReadOnlyList<int>? orderedPageObjectNumbers = null,
        int outputPageIndexOffset = 0,
        IReadOnlyDictionary<int, int>? outputPageIndexByPageObjectNumber = null) {
        var namedDestinations = BuildNamedDestinationsForPages(sourceObjects, catalogState.NamedDestinations, copiedPageObjectIds);
        var namedDestinationNameTree = BuildNamedDestinationNameTreeForPages(sourceObjects, catalogState.NamedDestinationNameTree, copiedPageObjectIds);
        var openAction = BuildOpenActionForPages(sourceObjects, catalogState.OpenAction, copiedPageObjectIds);
        var outlines = BuildOutlinesForPages(sourceObjects, catalogState.Outlines, copiedPageObjectIds);
        var pageLabels = BuildPageLabelsForPages(sourceObjects, catalogState.PageLabels, orderedPageObjectNumbers, outputPageIndexOffset, outputPageIndexByPageObjectNumber, catalogState.SourcePageObjectNumbers);
        string? pageMode = outlines is null && string.Equals(catalogState.PageMode, "UseOutlines", StringComparison.Ordinal)
            ? null
            : catalogState.PageMode;
        return new CatalogRewriteState(pageMode, catalogState.PageLayout, catalogState.CatalogVersion, catalogState.CatalogLanguage, outlines, pageLabels, namedDestinations, namedDestinationNameTree, openAction, catalogState.ViewerPreferences, catalogState.XmpMetadata, catalogState.CatalogUri, catalogState.OutputIntents, catalogState.EmbeddedFiles, catalogState.AssociatedFiles, catalogState.OptionalContent);
    }

    private static Dictionary<int, Dictionary<string, PdfObject>>? BuildPageOverridesWithFilteredNamedDestinationLinks(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        IReadOnlyList<int> pageObjectNumbers,
        Dictionary<int, Dictionary<string, PdfObject>>? pageOverrides,
        CatalogRewriteState catalogState) {
        var availableDestinationNames = GetNamedDestinationNames(sourceObjects, catalogState);
        Dictionary<int, Dictionary<string, PdfObject>>? result = null;

        if (pageOverrides is not null && pageOverrides.Count > 0) {
            result = new Dictionary<int, Dictionary<string, PdfObject>>();
            foreach (var pageEntry in pageOverrides) {
                result[pageEntry.Key] = new Dictionary<string, PdfObject>(pageEntry.Value);
            }
        }

        var visitedPages = new HashSet<int>();
        foreach (int pageObjectNumber in pageObjectNumbers) {
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

            if (!TryFilterNamedDestinationLinkAnnotations(sourceObjects, annotationsObject, availableDestinationNames, out var filteredAnnotations)) {
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
        CatalogRewriteState catalogState) {
        var names = new HashSet<string>(StringComparer.Ordinal);

        if (ResolveDictionary(sourceObjects, catalogState.NamedDestinations) is PdfDictionary directDestinations) {
            foreach (var name in directDestinations.Items.Keys) {
                names.Add(name);
            }
        }

        if (ResolveDictionary(sourceObjects, catalogState.NamedDestinationNameTree) is PdfDictionary nameTree &&
            nameTree.Items.TryGetValue("Names", out var namesObject) &&
            ResolveObject(sourceObjects, namesObject) is PdfArray nameArray) {
            for (int i = 0; i + 1 < nameArray.Items.Count; i += 2) {
                if (TryGetNamedDestinationName(sourceObjects, nameArray.Items[i], out string? destinationName)) {
                    names.Add(destinationName!);
                }
            }
        }

        return names;
    }

    private static bool TryFilterNamedDestinationLinkAnnotations(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        PdfObject? annotationsObject,
        HashSet<string> availableDestinationNames,
        out PdfArray filteredAnnotations) {
        filteredAnnotations = new PdfArray();
        if (ResolveObject(sourceObjects, annotationsObject) is not PdfArray annotations) {
            return false;
        }

        bool removed = false;
        foreach (var annotation in annotations.Items) {
            if (TryGetNamedDestinationLinkName(sourceObjects, annotation, out string? destinationName) &&
                !availableDestinationNames.Contains(destinationName!)) {
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

    private static ClonedAnnotationState BuildClonedAnnotationState(
        Dictionary<int, PdfIndirectObject> sourceObjects,
        int pageObjectNumber,
        ref int nextObjectId) {
        if (!sourceObjects.TryGetValue(pageObjectNumber, out var pageObject) ||
            pageObject.Value is not PdfDictionary pageDictionary ||
            !pageDictionary.Items.TryGetValue("Annots", out var annotationsObject) ||
            ResolveObject(sourceObjects, annotationsObject) is not PdfArray annotations) {
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

        if (!hasClonedIndirectAnnotation) {
            return ClonedAnnotationState.Empty;
        }

        return new ClonedAnnotationState(
            new Dictionary<string, PdfObject>(StringComparer.Ordinal) {
                ["Annots"] = clonedAnnotations
            },
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

    internal static byte[] SerializePageDictionary(PdfDictionary dictionary, int sourceId, SerializationContext context) {
        var sb = new StringBuilder();
        sb.Append("<< ");

        bool hasType = false;
        context.PageOverrides.TryGetValue(sourceId, out var pageOverrides);
        foreach (var entry in dictionary.Items) {
            if (string.Equals(entry.Key, "Parent", StringComparison.Ordinal)) {
                continue;
            }

            if (pageOverrides is not null && pageOverrides.ContainsKey(entry.Key)) {
                continue;
            }

            if (string.Equals(entry.Key, "Type", StringComparison.Ordinal)) {
                hasType = true;
            }

            AppendDictionaryEntry(sb, entry.Key, entry.Value, context);
        }

        if (!hasType) {
            sb.Append("/Type /Page ");
        }

        sb.Append("/Parent ")
            .Append(PdfSyntaxEscaper.IndirectReference(context.PagesObjectId))
            .Append(' ');

        if (context.MaterializedPageValues.TryGetValue(sourceId, out var inherited)) {
            foreach (var entry in inherited) {
                if (pageOverrides is not null && pageOverrides.ContainsKey(entry.Key)) {
                    continue;
                }

                if (!dictionary.Items.ContainsKey(entry.Key)) {
                    AppendDictionaryEntry(sb, entry.Key, entry.Value, context);
                }
            }
        }

        if (pageOverrides is not null) {
            foreach (var entry in pageOverrides) {
                AppendDictionaryEntry(sb, entry.Key, entry.Value, context);
            }
        }

        sb.Append(">>\n");
        return PdfEncoding.Latin1GetBytes(sb.ToString());
    }

    internal static byte[] SerializeObject(PdfObject value, SerializationContext context) {
        if (value is PdfStream stream) {
            return SerializeStream(stream, context);
        }

        var sb = new StringBuilder();
        AppendObject(sb, value, context);
        sb.Append('\n');
        return PdfEncoding.Latin1GetBytes(sb.ToString());
    }

    private static byte[] SerializeStream(PdfStream stream, SerializationContext context) {
        string dictionary = BuildStreamDictionary(stream, context);
        return SerializeStreamBody(dictionary, stream.Data);
    }

    private static string BuildStreamDictionary(PdfStream stream, SerializationContext context) {
        var sb = new StringBuilder();
        sb.Append("<< ");
        foreach (var entry in stream.Dictionary.Items) {
            if (!string.Equals(entry.Key, "Length", StringComparison.Ordinal)) {
                AppendDictionaryEntry(sb, entry.Key, entry.Value, context);
            }
        }

        sb.Append("/Length ")
            .Append(stream.Data.Length.ToString(CultureInfo.InvariantCulture))
            .Append(" >>");

        return sb.ToString();
    }

    private static byte[] SerializeStreamBody(string dictionary, byte[] data) {
        return PdfObjectBytes.WrapStreamBody(dictionary, data);
    }

    private static void AppendDictionaryEntry(StringBuilder sb, string key, PdfObject value, SerializationContext context) {
        sb.Append('/').Append(PdfSyntaxEscaper.Name(key)).Append(' ');
        AppendObject(sb, value, context);
        sb.Append(' ');
    }

    private static void AppendObject(StringBuilder sb, PdfObject value, SerializationContext context) {
        switch (value) {
            case PdfNumber number:
                sb.Append(FormatNumber(number.Value));
                break;
            case PdfBoolean boolean:
                sb.Append(boolean.Value ? "true" : "false");
                break;
            case PdfName name:
                sb.Append('/').Append(PdfSyntaxEscaper.Name(name.Name));
                break;
            case PdfStringObj text:
                sb.Append(PdfSyntaxEscaper.LiteralString(text.Value));
                break;
            case PdfNull:
                sb.Append("null");
                break;
            case PdfReference reference:
                ValidateReferenceGeneration(reference, context);
                if (!context.NumberMap.TryGetValue(reference.ObjectNumber, out int newObjectNumber)) {
                    throw new InvalidOperationException("PDF object " + reference.ObjectNumber.ToString(CultureInfo.InvariantCulture) + " was referenced but not copied.");
                }

                sb.Append(PdfSyntaxEscaper.IndirectReference(newObjectNumber));
                break;
            case PdfArray array:
                sb.Append("[ ");
                foreach (var item in array.Items) {
                    AppendObject(sb, item, context);
                    sb.Append(' ');
                }
                sb.Append(']');
                break;
            case PdfDictionary dictionary:
                sb.Append("<< ");
                foreach (var entry in dictionary.Items) {
                    AppendDictionaryEntry(sb, entry.Key, entry.Value, context);
                }
                sb.Append(">>");
                break;
            case PdfStream:
                throw new NotSupportedException("Direct PDF streams inside arrays or dictionaries are not supported by page extraction yet.");
            default:
                throw new NotSupportedException("Unsupported PDF object type: " + value.GetType().Name);
        }
    }

    private static void ValidateReferenceGeneration(PdfReference reference, SerializationContext context) {
        if (context.SourceObjectGenerations.TryGetValue(reference.ObjectNumber, out int activeGeneration)) {
            if (reference.Generation != activeGeneration) {
                throw BuildGenerationMismatchException(reference, activeGeneration);
            }

            return;
        }

        if (reference.ObjectNumber < 0 && reference.Generation != 0) {
            throw new InvalidOperationException("Additional PDF object " + reference.ObjectNumber.ToString(CultureInfo.InvariantCulture) + " was referenced with generation " + reference.Generation.ToString(CultureInfo.InvariantCulture) + "; additional rewrite objects must use generation 0.");
        }
    }

    private static InvalidOperationException BuildGenerationMismatchException(PdfReference reference, int activeGeneration) {
        return new InvalidOperationException(
            "PDF object " +
            reference.ObjectNumber.ToString(CultureInfo.InvariantCulture) +
            " " +
            reference.Generation.ToString(CultureInfo.InvariantCulture) +
            " R was referenced, but the active object generation is " +
            activeGeneration.ToString(CultureInfo.InvariantCulture) +
            ".");
    }

    internal static string BuildInfoDictionary(PdfMetadata metadata) {
        return PdfInfoDictionaryBuilder.Build(metadata);
    }

    internal static byte[] WrapObject(int objectNumber, byte[] body) {
        return PdfObjectBytes.WrapIndirectObject(objectNumber, body);
    }

    internal static byte[] Assemble(List<byte[]> objects, int catalogId, int infoId) {
        return PdfFileAssembler.Assemble(objects, catalogId, infoId);
    }

    private static string FormatNumber(double value) {
        if (Math.Abs(value % 1) < 0.0000001) {
            return ((long)Math.Round(value)).ToString(CultureInfo.InvariantCulture);
        }

        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }

    private static void WriteOutput(string outputPath, byte[] bytes) {
        string fullPath = ValidateOutputPath(outputPath);
        var directory = Path.GetDirectoryName(fullPath);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        File.WriteAllBytes(fullPath, bytes);
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

    internal sealed class ObjectCollector {
        private static readonly string[] InheritablePageKeys = { "Resources", "MediaBox", "CropBox", "Rotate" };
        private readonly Dictionary<int, PdfIndirectObject> _sourceObjects;
        private readonly Dictionary<int, Dictionary<string, PdfObject>> _pageOverrides;
        private readonly List<int> _objectIds = new();
        private readonly HashSet<int> _visited = new();

        public ObjectCollector(Dictionary<int, PdfIndirectObject> sourceObjects, Dictionary<int, Dictionary<string, PdfObject>>? pageOverrides = null) {
            _sourceObjects = sourceObjects;
            _pageOverrides = pageOverrides ?? new Dictionary<int, Dictionary<string, PdfObject>>();
        }

        public IReadOnlyList<int> ObjectIds => _objectIds;

        public HashSet<int> PageObjectIds { get; } = new();

        public Dictionary<int, Dictionary<string, PdfObject>> MaterializedPageValues { get; } = new();

        public void CollectObjectGraph(PdfObject? value) {
            if (value is not null) {
                CollectReferences(value, isPageObject: false);
            }
        }

        public void CollectPage(int objectNumber) {
            if (!_sourceObjects.TryGetValue(objectNumber, out var indirect) || indirect.Value is not PdfDictionary pageDictionary) {
                throw new InvalidOperationException("PDF page object " + objectNumber.ToString(CultureInfo.InvariantCulture) + " was not found.");
            }

            PageObjectIds.Add(objectNumber);
            MaterializeInheritedPageValues(objectNumber, pageDictionary);
            CollectObject(objectNumber, isPageObject: true);
        }

        private void CollectObject(int objectNumber, bool isPageObject) {
            if (!_visited.Add(objectNumber)) {
                return;
            }

            if (!_sourceObjects.TryGetValue(objectNumber, out var indirect)) {
                if (objectNumber < 0) {
                    return;
                }

                throw new InvalidOperationException("PDF object " + objectNumber.ToString(CultureInfo.InvariantCulture) + " was referenced but not found.");
            }

            _objectIds.Add(objectNumber);
            _pageOverrides.TryGetValue(objectNumber, out var pageOverrides);
            CollectReferences(indirect.Value, isPageObject, pageOverrides);
        }

        private void CollectReferences(PdfObject value, bool isPageObject, Dictionary<string, PdfObject>? pageOverrides = null) {
            switch (value) {
                case PdfReference reference:
                    if (reference.ObjectNumber >= 0 &&
                        _sourceObjects.TryGetValue(reference.ObjectNumber, out var referenced) &&
                        referenced.Generation != reference.Generation) {
                        throw BuildGenerationMismatchException(reference, referenced.Generation);
                    }

                    CollectObject(reference.ObjectNumber, isPageObject: false);
                    break;
                case PdfArray array:
                    foreach (var item in array.Items) {
                        CollectReferences(item, isPageObject: false);
                    }

                    break;
                case PdfDictionary dictionary:
                    foreach (var entry in dictionary.Items) {
                        if (isPageObject &&
                            (string.Equals(entry.Key, "Parent", StringComparison.Ordinal) ||
                            (pageOverrides is not null && pageOverrides.ContainsKey(entry.Key)))) {
                            continue;
                        }

                        CollectReferences(entry.Value, isPageObject: false);
                    }

                    if (isPageObject && pageOverrides is not null) {
                        foreach (var entry in pageOverrides) {
                            CollectReferences(entry.Value, isPageObject: false);
                        }
                    }

                    break;
                case PdfStream stream:
                    foreach (var entry in stream.Dictionary.Items) {
                        if (!string.Equals(entry.Key, "Length", StringComparison.Ordinal)) {
                            CollectReferences(entry.Value, isPageObject: false);
                        }
                    }

                    break;
            }
        }

        private void MaterializeInheritedPageValues(int pageObjectNumber, PdfDictionary pageDictionary) {
            foreach (string key in InheritablePageKeys) {
                if (pageDictionary.Items.ContainsKey(key)) {
                    continue;
                }

                var inherited = ResolveInheritedValue(pageDictionary, key);
                if (inherited is null) {
                    continue;
                }

                if (!MaterializedPageValues.TryGetValue(pageObjectNumber, out var values)) {
                    values = new Dictionary<string, PdfObject>(StringComparer.Ordinal);
                    MaterializedPageValues[pageObjectNumber] = values;
                }

                values[key] = inherited;
                CollectReferences(inherited, isPageObject: false);
            }
        }

        private PdfObject? ResolveInheritedValue(PdfDictionary pageDictionary, string key) {
            PdfDictionary? current = pageDictionary;
            int guard = 0;
            while (current is not null && guard++ < 100) {
                if (current.Items.TryGetValue(key, out var value)) {
                    return value;
                }

                if (!current.Items.TryGetValue("Parent", out var parentObj) ||
                    parentObj is not PdfReference parentReference ||
                    !PdfObjectLookup.TryGet(_sourceObjects, parentReference, out var parentIndirect) ||
                    parentIndirect.Value is not PdfDictionary parentDictionary) {
                    return null;
                }

                current = parentDictionary;
            }

            return null;
        }
    }

    internal sealed class SerializationContext {
        public SerializationContext(
            Dictionary<int, int> numberMap,
            int pagesObjectId,
            Dictionary<int, Dictionary<string, PdfObject>> materializedPageValues,
            Dictionary<int, PdfIndirectObject>? sourceObjects = null,
            Dictionary<int, Dictionary<string, PdfObject>>? pageOverrides = null) {
            NumberMap = numberMap;
            PagesObjectId = pagesObjectId;
            MaterializedPageValues = materializedPageValues;
            SourceObjectGenerations = sourceObjects?.ToDictionary(entry => entry.Key, entry => entry.Value.Generation) ?? new Dictionary<int, int>();
            PageOverrides = pageOverrides ?? new Dictionary<int, Dictionary<string, PdfObject>>();
        }

        public Dictionary<int, int> NumberMap { get; }

        public int PagesObjectId { get; }

        public Dictionary<int, Dictionary<string, PdfObject>> MaterializedPageValues { get; }

        public Dictionary<int, int> SourceObjectGenerations { get; }

        public Dictionary<int, Dictionary<string, PdfObject>> PageOverrides { get; }
    }

    internal sealed class AdditionalObject {
        public AdditionalObject(int pseudoObjectNumber, PdfObject value) {
            PseudoObjectNumber = pseudoObjectNumber;
            Value = value;
        }

        public int PseudoObjectNumber { get; }

        public PdfObject Value { get; }
    }

    private sealed class ClonedPageObject {
        public ClonedPageObject(
            int sourcePageObjectNumber,
            int outputPageObjectNumber,
            Dictionary<string, PdfObject>? pageOverrides,
            Dictionary<int, int> annotationObjectMap) {
            SourcePageObjectNumber = sourcePageObjectNumber;
            OutputPageObjectNumber = outputPageObjectNumber;
            PageOverrides = pageOverrides;
            AnnotationObjectMap = annotationObjectMap;
        }

        public int SourcePageObjectNumber { get; }

        public int OutputPageObjectNumber { get; }

        public Dictionary<string, PdfObject>? PageOverrides { get; }

        public Dictionary<int, int> AnnotationObjectMap { get; }
    }

    private sealed class ClonedAnnotationState {
        public static readonly ClonedAnnotationState Empty = new ClonedAnnotationState(null, new Dictionary<int, int>());

        public ClonedAnnotationState(Dictionary<string, PdfObject>? pageOverrides, Dictionary<int, int> annotationObjectMap) {
            PageOverrides = pageOverrides;
            AnnotationObjectMap = annotationObjectMap;
        }

        public Dictionary<string, PdfObject>? PageOverrides { get; }

        public Dictionary<int, int> AnnotationObjectMap { get; }
    }

    private sealed class PageLabelEntry {
        public PageLabelEntry(int startPageIndex, PdfDictionary labelDictionary) {
            StartPageIndex = startPageIndex;
            LabelDictionary = labelDictionary;
        }

        public int StartPageIndex { get; }

        public PdfDictionary LabelDictionary { get; }
    }

    private sealed class NamedDestinationNameTreeEntry {
        public NamedDestinationNameTreeEntry(PdfStringObj name, PdfObject destination) {
            Name = name;
            Destination = destination;
        }

        public PdfStringObj Name { get; }

        public PdfObject Destination { get; }
    }

    internal sealed class CatalogRewriteState {
        public static readonly CatalogRewriteState Empty = new CatalogRewriteState(null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null);

        public CatalogRewriteState(string? pageMode, string? pageLayout, PdfObject? catalogVersion, PdfObject? catalogLanguage, PdfObject? outlines, PdfObject? pageLabels, PdfObject? namedDestinations, PdfObject? namedDestinationNameTree, PdfObject? openAction, PdfObject? viewerPreferences, PdfObject? xmpMetadata, PdfObject? catalogUri, PdfObject? outputIntents, PdfObject? embeddedFiles, PdfObject? associatedFiles, PdfObject? optionalContent, IReadOnlyList<int>? sourcePageObjectNumbers = null) {
            PageMode = string.IsNullOrEmpty(pageMode) ? null : pageMode;
            PageLayout = string.IsNullOrEmpty(pageLayout) ? null : pageLayout;
            CatalogVersion = catalogVersion;
            CatalogLanguage = catalogLanguage;
            Outlines = outlines;
            PageLabels = pageLabels;
            NamedDestinations = namedDestinations;
            NamedDestinationNameTree = namedDestinationNameTree;
            OpenAction = openAction;
            ViewerPreferences = viewerPreferences;
            XmpMetadata = xmpMetadata;
            CatalogUri = catalogUri;
            OutputIntents = outputIntents;
            EmbeddedFiles = embeddedFiles;
            AssociatedFiles = associatedFiles;
            OptionalContent = optionalContent;
            SourcePageObjectNumbers = sourcePageObjectNumbers;
        }

        public string? PageMode { get; }

        public string? PageLayout { get; }

        public PdfObject? CatalogVersion { get; }

        public PdfObject? CatalogLanguage { get; }

        public PdfObject? Outlines { get; }

        public PdfObject? PageLabels { get; }

        public PdfObject? NamedDestinations { get; }

        public PdfObject? NamedDestinationNameTree { get; }

        public PdfObject? OpenAction { get; }

        public PdfObject? ViewerPreferences { get; }

        public PdfObject? XmpMetadata { get; }

        public PdfObject? CatalogUri { get; }

        public PdfObject? OutputIntents { get; }

        public PdfObject? EmbeddedFiles { get; }

        public PdfObject? AssociatedFiles { get; }

        public PdfObject? OptionalContent { get; }

        public IReadOnlyList<int>? SourcePageObjectNumbers { get; }
    }
}
