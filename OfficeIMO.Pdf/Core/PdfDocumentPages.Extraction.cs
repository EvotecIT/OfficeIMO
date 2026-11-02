using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentPages {
    /// <summary>
    /// Creates a new PDF containing selected pages in caller order.
    /// </summary>
    public PdfDocument Extract(params int[] pageNumbers) {
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        byte[] input = _document.GetBytesForOperation();
        byte[] output = PdfPageExtractor.ExtractPages(input, pageNumbers, _document.ReadOptions, _document.GetOpenedReadDocumentFactory());
        return AdoptExtractedOutput(input, output, pageNumbers.Length, _document.ReadOptions);
    }

    /// <summary>Creates a new PDF containing selected pages in caller order with cooperative cancellation.</summary>
    public PdfDocument Extract(CancellationToken cancellationToken, params int[] pageNumbers) {
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        cancellationToken.ThrowIfCancellationRequested();
        byte[] input = _document.GetBytesForOperation(cancellationToken);
        byte[] output = PdfPageExtractor.ExtractPages(
            input,
            pageNumbers,
            _document.ReadOptions,
            _document.GetOpenedReadDocumentFactory(cancellationToken),
            cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        PdfDocument extracted = AdoptExtractedOutput(input, output, pageNumbers.Length, _document.ReadOptions,
            cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        return extracted;
    }

    /// <summary>Extracts pages with a byte budget enforced during the first canonical serialization.</summary>
    internal PdfDocument Extract(int[] pageNumbers, long maximumOutputBytes) {
        Guard.NotNull(pageNumbers, nameof(pageNumbers));
        byte[] input = _document.GetBytesForOperation();
        byte[] output = PdfPageExtractor.ExtractPages(input, pageNumbers, _document.ReadOptions, maximumOutputBytes, _document.GetOpenedReadDocumentFactory());
        return AdoptExtractedOutput(input, output, pageNumbers.Length, _document.ReadOptions);
    }

    /// <summary>
    /// Creates a new PDF containing one inclusive one-based page range.
    /// </summary>
    public PdfDocument Extract(PdfPageRange pageRange) {
        return Extract(pageRange.ToPageNumbers());
    }

    /// <summary>
    /// Creates a new PDF containing selected pages in caller order.
    /// </summary>
    public PdfDocument Extract(PdfPageSelection selection) {
        Guard.NotNull(selection, nameof(selection));
        return Extract(selection, _document.ReadOptions);
    }

    private PdfDocument Extract(PdfPageSelection selection, PdfLoadOptions? options, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        Guard.NotNull(selection, nameof(selection));
        byte[] input = _document.GetBytesForOperation(cancellationToken);
        byte[] output = PdfPageExtractor.ExtractPageRanges(input, selection.ToRanges(), options, _document.GetOpenedReadDocumentFactory(options, cancellationToken), cancellationToken);
        return AdoptExtractedOutput(input, output, selection.PageCount, options, cancellationToken);
    }

    /// <summary>
    /// Attempts to create a new PDF containing selected pages in caller order, returning diagnostics when blocked or failed.
    /// </summary>
    internal PdfOperationResult<PdfDocument> ExtractResult(PdfPageSelection selection, PdfLoadOptions? options = null) {
        Guard.NotNull(selection, nameof(selection));
        return TryPageExtractionOperation("Extract pages", effectiveOptions => Extract(selection, effectiveOptions), options);
    }

    /// <summary>
    /// Creates a new PDF containing comma- or semicolon-separated inclusive page ranges.
    /// </summary>
    public PdfDocument Extract(string pageRanges) {
        return Extract(PdfPageSelection.Parse(pageRanges));
    }

    /// <summary>
    /// Attempts to create a new PDF containing pages described by page ranges, returning diagnostics when blocked or failed.
    /// </summary>
    internal PdfOperationResult<PdfDocument> ExtractResult(string pageRanges, PdfLoadOptions? options = null) {
        return TryPageExtractionOperation("Extract pages", effectiveOptions => Extract(PdfPageSelection.Parse(pageRanges), effectiveOptions), options);
    }

    /// <summary>
    /// Creates one PDF per page.
    /// </summary>
    public IReadOnlyList<PdfDocument> Split() {
        return Split(_document.ReadOptions);
    }

    private PdfDocument[] Split(PdfLoadOptions? options, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        byte[] input = _document.GetBytesForOperation(cancellationToken);
        IReadOnlyList<byte[]> outputs = PdfPageExtractor.SplitPages(input, options, _document.GetOpenedReadDocumentFactory(options, cancellationToken), cancellationToken);
        return AdoptSplitOutputs(input, outputs, Enumerable.Repeat(1, outputs.Count).ToArray(), options, cancellationToken);
    }

    /// <summary>
    /// Creates PDFs containing consecutive groups of the requested page count.
    /// </summary>
    public IReadOnlyList<PdfDocument> Split(int pagesPerDocument) {
        if (pagesPerDocument <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pagesPerDocument), pagesPerDocument, "Pages per document must be greater than zero.");
        }

        return Split(pagesPerDocument, _document.ReadOptions);
    }

    private PdfDocument[] Split(int pagesPerDocument, PdfLoadOptions? options, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (pagesPerDocument <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pagesPerDocument), pagesPerDocument, "Pages per document must be greater than zero.");
        }

        PdfReadDocument readDocument = _document.GetReadDocument(options ?? _document.ReadOptions, cancellationToken);
        readDocument.DemandContentExtraction("logical object");
        int pageCount = readDocument.Pages.Count;
        if (pageCount == 0) {
            throw new InvalidOperationException("PDF does not contain any readable pages.");
        }

        var ranges = new List<PdfPageRange>();
        for (int firstPage = 1; firstPage <= pageCount;) {
            cancellationToken.ThrowIfCancellationRequested();
            int lastPage = firstPage + Math.Min(pagesPerDocument - 1, pageCount - firstPage);
            ranges.Add(PdfPageRange.From(firstPage, lastPage));
            if (lastPage == pageCount) break;
            firstPage = lastPage + 1;
        }

        return Split(ranges, options, cancellationToken);
    }

    /// <summary>
    /// Creates one PDF for each supplied page selection.
    /// </summary>
    public IReadOnlyList<PdfDocument> Split(params PdfPageSelection[] selections) {
        Guard.NotNull(selections, nameof(selections));
        if (selections.Length == 0) {
            throw new ArgumentException("At least one page selection must be specified.", nameof(selections));
        }

        return Split(selections, _document.ReadOptions);
    }

    private PdfDocument[] Split(PdfPageSelection[] selections, PdfLoadOptions? options, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        byte[] input = _document.GetBytesForOperation(cancellationToken);
        var outputs = new byte[selections.Length][];
        var outputPageCounts = new int[selections.Length];
        var source = new PdfPageExtractor.ExtractionSession(input, options, _document.GetOpenedReadDocumentFactory(options, cancellationToken), cancellationToken);
        for (int i = 0; i < selections.Length; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            Guard.NotNull(selections[i], nameof(selections));
            outputs[i] = source.Extract(selections[i].ToRanges());
            outputPageCounts[i] = selections[i].PageCount;
        }

        return AdoptSplitOutputs(input, outputs, outputPageCounts, options, cancellationToken);
    }

    /// <summary>
    /// Creates one PDF for each supplied inclusive page range.
    /// </summary>
    public IReadOnlyList<PdfDocument> Split(IEnumerable<PdfPageRange> pageRanges) {
        Guard.NotNull(pageRanges, nameof(pageRanges));
        PdfPageRange[] ranges = CaptureSplitRanges(pageRanges, CancellationToken.None);
        if (ranges.Length == 0) {
            throw new ArgumentException("At least one page range must be specified.", nameof(pageRanges));
        }

        return Split(ranges, _document.ReadOptions);
    }

    private PdfDocument[] Split(IEnumerable<PdfPageRange> pageRanges, PdfLoadOptions? options, CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        Guard.NotNull(pageRanges, nameof(pageRanges));
        PdfPageRange[] ranges = CaptureSplitRanges(pageRanges, cancellationToken);
        if (ranges.Length == 0) {
            throw new ArgumentException("At least one page range must be specified.", nameof(pageRanges));
        }

        byte[] input = _document.GetBytesForOperation(cancellationToken);
        return AdoptSplitOutputs(
            input,
            PdfPageExtractor.SplitPageRanges(input, ranges, options, _document.GetOpenedReadDocumentFactory(options, cancellationToken), cancellationToken),
            ranges.Select(static range => range.PageCount).ToArray(),
            options, cancellationToken);
    }

    /// <summary>Creates one split part within the remaining output budget, retaining split preservation semantics.</summary>
    internal PdfDocument Split(PdfPageRange pageRange, long maximumOutputBytes) {
        byte[] input = _document.GetBytesForOperation();
        byte[] output = PdfPageExtractor.ExtractPages(input, pageRange.ToPageNumbers(), _document.ReadOptions, maximumOutputBytes, _document.GetOpenedReadDocumentFactory());
        return AdoptSplitOutputs(input, new[] { output }, new[] { pageRange.PageCount }, _document.ReadOptions)[0];
    }

    private PdfDocument[] AdoptSplitOutputs(
        byte[] input,
        IReadOnlyList<byte[]> outputs,
        int[] outputPageCounts,
        PdfLoadOptions? options,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        if (outputs.Count != outputPageCounts.Length) {
            throw new InvalidOperationException("Split output count did not match the planned page-count evidence.");
        }

        PdfArtifactInfo inputArtifact = _document.Pipeline.Output ??
            PdfArtifactInfo.Capture(input, _document.ReadOptions, cancellationToken);
        var documents = new PdfDocument[outputs.Count];
        for (int i = 0; i < outputs.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();
            documents[i] = _document.WithCanonicalBytesKnownPageCount(
                input,
                inputArtifact,
                outputs[i],
                outputPageCounts[i],
                options,
                "Split",
                cancellationToken);
        }

        return documents;
    }

    private PdfDocument AdoptExtractedOutput(
        byte[] input,
        byte[] output,
        int outputPageCount,
        PdfLoadOptions? options,
        CancellationToken cancellationToken = default) {
        cancellationToken.ThrowIfCancellationRequested();
        PdfArtifactInfo inputArtifact = _document.Pipeline.Output ??
            PdfArtifactInfo.Capture(input, _document.ReadOptions, cancellationToken);
        return _document.WithCanonicalBytesKnownPageCount(
            input,
            inputArtifact,
            output,
            outputPageCount,
            options,
            "Extract",
            cancellationToken);
    }

    /// <summary>
    /// Creates one PDF for each comma- or semicolon-separated inclusive page range.
    /// </summary>
    public IReadOnlyList<PdfDocument> Split(string pageRanges) {
        return Split(PdfPageRange.ParseMany(pageRanges));
    }

    /// <summary>Splits this document using bounded page-count, target-size, and content-boundary production rules.</summary>
    public IReadOnlyList<PdfProductionSplitPart> SplitForProduction(PdfProductionSplitOptions options) {
        return SplitForProductionResult(options).RequireValue();
    }

    /// <summary>Splits this document and returns the generated parts with bounded-probe evidence.</summary>
    internal PdfProductionSplitResult SplitForProductionResult(PdfProductionSplitOptions options) {
        Guard.NotNull(options, nameof(options));
        return PdfProductionSplitter.Split(_document.GetBytesForOperation(options.CancellationToken), options, _document.ReadOptions);
    }

}
