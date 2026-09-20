using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Merges loaded or generated documents using an explicit structure policy.</summary>
    public static PdfDocument Merge(PdfMergeOptions options, params PdfDocument[] documents) =>
        Merge(options, (IEnumerable<PdfDocument>)documents);

    /// <summary>Merges loaded or generated documents using an explicit structure policy.</summary>
    public static PdfDocument Merge(PdfMergeOptions options, IEnumerable<PdfDocument> documents) =>
        MergeResult(options, documents).RequireValue();

    /// <summary>
    /// Merges caller-owned PDF byte payloads using an explicit structure policy.
    /// The inputs are consumed synchronously and are not retained by the returned document.
    /// </summary>
    public static PdfDocument MergeBytes(PdfMergeOptions options, IEnumerable<byte[]> pdfs) =>
        MergeBytesResult(options, pdfs).RequireValue();

    /// <summary>Merges loaded or generated documents with an explicit structure policy and returns readback evidence.</summary>
    public static PdfMergeResult MergeResult(PdfMergeOptions options, params PdfDocument[] documents) =>
        MergeResult(options, (IEnumerable<PdfDocument>)documents);

    /// <summary>Merges loaded or generated documents with an explicit structure policy and returns readback evidence.</summary>
    public static PdfMergeResult MergeResult(PdfMergeOptions options, IEnumerable<PdfDocument> documents) {
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
        PdfLoadOptions[] readOptions = sources.Select(static document => document.ReadOptions).ToArray();
        return PdfMerger.MergeResult(options, bytes, readOptions);
    }

    /// <summary>
    /// Merges caller-owned PDF byte payloads with an explicit structure policy and returns readback evidence.
    /// The inputs are consumed synchronously and are not retained by the result.
    /// </summary>
    public static PdfMergeResult MergeBytesResult(PdfMergeOptions options, IEnumerable<byte[]> pdfs) {
        Guard.NotNull(options, nameof(options));
        List<byte[]> sources = CollectMergeByteSources(pdfs, CancellationToken.None);
        var readOptions = new PdfLoadOptions[sources.Count];
        for (int index = 0; index < readOptions.Length; index++) {
            readOptions[index] = PdfLoadOptions.Default;
        }
        return PdfMerger.MergeResult(options, sources, readOptions);
    }

    /// <summary>Merges this PDF with another loaded or generated PDF using an explicit structure policy.</summary>
    public PdfDocument MergeWith(PdfDocument document, PdfMergeOptions options) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(options, nameof(options));
        return MergeResult(options, this, document).RequireValue();
    }

    /// <summary>Merges this PDF with another loaded or generated PDF and returns structure and readback evidence.</summary>
    public PdfMergeResult MergeWithResult(PdfDocument document, PdfMergeOptions options) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(options, nameof(options));
        return MergeResult(options, this, document);
    }
}
