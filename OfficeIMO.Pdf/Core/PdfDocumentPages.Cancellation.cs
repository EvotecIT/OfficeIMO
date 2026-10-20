using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocumentPages {
    /// <summary>Extracts one inclusive page range while observing cancellation through parsing, rewriting and output adoption.</summary>
    public PdfDocument Extract(PdfPageRange pageRange, CancellationToken cancellationToken) =>
        Extract(PdfPageSelection.FromRanges(pageRange), _document.ReadOptions, cancellationToken);

    /// <summary>Extracts a selection while observing cancellation and retaining this document's captured read limits.</summary>
    public PdfDocument Extract(PdfPageSelection selection, CancellationToken cancellationToken) =>
        Extract(selection, _document.ReadOptions, cancellationToken);

    /// <summary>Extracts parsed page ranges while observing cancellation and retaining the source document unchanged.</summary>
    public PdfDocument Extract(string pageRanges, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return Extract(PdfPageSelection.Parse(pageRanges), _document.ReadOptions, cancellationToken);
    }

    /// <summary>Extracts a document-relative selection while observing cancellation during resolution and rewriting.</summary>
    public PdfDocument Extract(PdfPageSelector selector, CancellationToken cancellationToken) =>
        Extract(ResolveSelector(selector, _document.ReadOptions, cancellationToken), _document.ReadOptions, cancellationToken);

    /// <summary>Creates one PDF per source page with cooperative cancellation; the source is unchanged.</summary>
    public IReadOnlyList<PdfDocument> Split(CancellationToken cancellationToken) =>
        Split(_document.ReadOptions, cancellationToken);

    /// <summary>Creates consecutive page groups with cooperative cancellation and the source's captured read limits.</summary>
    public IReadOnlyList<PdfDocument> Split(int pagesPerDocument, CancellationToken cancellationToken) =>
        Split(pagesPerDocument, _document.ReadOptions, cancellationToken);

    /// <summary>Creates caller-ordered selection parts with cooperative cancellation through every part.</summary>
    public IReadOnlyList<PdfDocument> Split(CancellationToken cancellationToken, params PdfPageSelection[] selections) {
        Guard.NotNull(selections, nameof(selections));
        if (selections.Length == 0) throw new ArgumentException("At least one page selection must be specified.", nameof(selections));
        return Split(selections, _document.ReadOptions, cancellationToken);
    }

    /// <summary>Creates caller-ordered range parts with cancellation during range enumeration and extraction.</summary>
    public IReadOnlyList<PdfDocument> Split(IEnumerable<PdfPageRange> pageRanges, CancellationToken cancellationToken) =>
        Split(pageRanges, _document.ReadOptions, cancellationToken);

    /// <summary>Creates parts from parsed inclusive ranges with cooperative cancellation.</summary>
    public IReadOnlyList<PdfDocument> Split(string pageRanges, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return Split(PdfPageRange.ParseMany(pageRanges), _document.ReadOptions, cancellationToken);
    }

    private static PdfPageRange[] CaptureSplitRanges(IEnumerable<PdfPageRange> pageRanges, CancellationToken cancellationToken) {
        Guard.NotNull(pageRanges, nameof(pageRanges));
        var ranges = new List<PdfPageRange>();
        foreach (PdfPageRange range in pageRanges) {
            cancellationToken.ThrowIfCancellationRequested();
            ranges.Add(range);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return ranges.ToArray();
    }
}
