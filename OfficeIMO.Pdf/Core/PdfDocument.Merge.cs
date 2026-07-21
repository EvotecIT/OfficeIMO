namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Merges loaded or generated documents with an explicit structure policy and returns readback evidence.</summary>
    public static PdfMergeResult MergeWithReport(PdfMergeOptions options, params PdfDocument[] documents) =>
        MergeWithReport(options, (IEnumerable<PdfDocument>)documents);

    /// <summary>Merges loaded or generated documents with an explicit structure policy and returns readback evidence.</summary>
    public static PdfMergeResult MergeWithReport(PdfMergeOptions options, IEnumerable<PdfDocument> documents) {
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(documents, nameof(documents));
        PdfDocument[] sources = documents.ToArray();
        if (sources.Length == 0) {
            throw new ArgumentException("At least one PDF document must be supplied.", nameof(documents));
        }

        if (sources.Any(static document => document is null)) {
            throw new ArgumentException("PDF documents cannot contain null entries.", nameof(documents));
        }

        byte[][] bytes = sources.Select(static document => document.GetBytesForOperation()).ToArray();
        PdfReadOptions[] readOptions = sources.Select(static document => document.ReadOptions).ToArray();
        return PdfMerger.MergeWithReport(options, bytes, readOptions);
    }

    /// <summary>Merges this PDF with another loaded or generated PDF using an explicit structure policy.</summary>
    public PdfDocument MergeWith(PdfDocument document, PdfMergeOptions options) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(options, nameof(options));
        return MergeWithReport(options, this, document).ToDocument();
    }
}
