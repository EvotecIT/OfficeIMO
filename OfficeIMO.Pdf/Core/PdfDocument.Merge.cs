using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Merges documents in order using one structure policy.</summary>
    public static PdfDocument Merge(IEnumerable<PdfDocument> documents, PdfMergeOptions? options = null,
        CancellationToken cancellationToken = default) => Merge(documents, options, out _, cancellationToken);

    /// <summary>Merges documents and returns the completed policy and readback report.</summary>
    public static PdfDocument Merge(IEnumerable<PdfDocument> documents, PdfMergeOptions? options,
        out PdfMergeReport report, CancellationToken cancellationToken = default) {
        PdfMergeResult result = MergeResult(options ?? new PdfMergeOptions(), documents, cancellationToken);
        report = result.Report;
        return result.RequireValue();
    }

    /// <summary>Merges loaded or generated documents using an explicit structure policy.</summary>
    internal static PdfDocument Merge(PdfMergeOptions options, params PdfDocument[] documents) =>
        Merge(options, (IEnumerable<PdfDocument>)documents);

    /// <summary>Merges loaded or generated documents using an explicit structure policy.</summary>
    internal static PdfDocument Merge(PdfMergeOptions options, IEnumerable<PdfDocument> documents) =>
        MergeResult(options, documents).RequireValue();

    /// <summary>Merges documents using an explicit structure policy with cancellation through readback.</summary>
    internal static PdfDocument Merge(PdfMergeOptions options, IEnumerable<PdfDocument> documents, CancellationToken cancellationToken) =>
        MergeResult(options, documents, cancellationToken).RequireValue();

    /// <summary>
    /// Merges caller-owned PDF byte payloads using an explicit structure policy.
    /// The inputs are consumed synchronously and are not retained by the returned document.
    /// </summary>
    internal static PdfDocument MergeBytes(PdfMergeOptions options, IEnumerable<byte[]> pdfs) =>
        MergeBytesResult(options, pdfs).RequireValue();

    /// <summary>Merges caller-owned PDF payloads using an explicit structure policy with cooperative cancellation.</summary>
    internal static PdfDocument MergeBytes(PdfMergeOptions options, IEnumerable<byte[]> pdfs, CancellationToken cancellationToken) =>
        MergeBytesResult(options, pdfs, cancellationToken).RequireValue();

    /// <summary>
    /// Merges caller-owned PDF byte payloads using an explicit structure policy and returns the merged artifact directly.
    /// The inputs are consumed synchronously and are not retained; the returned array is owned by the caller.
    /// </summary>
    internal static byte[] MergeToBytes(PdfMergeOptions options, params byte[][] pdfs) =>
        MergeToBytes(options, (IEnumerable<byte[]>)pdfs);

    /// <summary>
    /// Merges caller-owned PDF byte payloads using an explicit structure policy and returns the merged artifact directly.
    /// The inputs are consumed synchronously and are not retained; the returned array is owned by the caller.
    /// </summary>
    internal static byte[] MergeToBytes(PdfMergeOptions options, IEnumerable<byte[]> pdfs) =>
        MergeBytesResult(options, pdfs).OwnedBytes;

    /// <summary>Merges caller-owned payloads with an explicit structure policy and cancellation, returning caller-owned bytes.</summary>
    internal static byte[] MergeToBytes(PdfMergeOptions options, IEnumerable<byte[]> pdfs, CancellationToken cancellationToken) =>
        MergeBytesResult(options, pdfs, cancellationToken).OwnedBytes;

    /// <summary>Merges loaded or generated documents with an explicit structure policy and returns readback evidence.</summary>
    internal static PdfMergeResult MergeResult(PdfMergeOptions options, params PdfDocument[] documents) =>
        MergeResult(options, (IEnumerable<PdfDocument>)documents);

    /// <summary>Merges loaded or generated documents with an explicit structure policy and returns readback evidence.</summary>
    internal static PdfMergeResult MergeResult(PdfMergeOptions options, IEnumerable<PdfDocument> documents) =>
        MergeResult(options, documents, CancellationToken.None);

    /// <summary>Merges documents with an explicit structure policy and cooperative cancellation, returning readback evidence.</summary>
    internal static PdfMergeResult MergeResult(PdfMergeOptions options, IEnumerable<PdfDocument> documents, CancellationToken cancellationToken) {
        Guard.NotNull(options, nameof(options));
        Guard.NotNull(documents, nameof(documents));
        cancellationToken.ThrowIfCancellationRequested();
        var retained = new List<PdfDocument>();
        foreach (PdfDocument document in documents) { cancellationToken.ThrowIfCancellationRequested(); retained.Add(document); }
        PdfDocument[] sources = retained.ToArray();
        if (sources.Length == 0) {
            throw new ArgumentException("At least one PDF document must be supplied.", nameof(documents));
        }

        if (sources.Any(static document => document is null)) {
            throw new ArgumentException("PDF documents cannot contain null entries.", nameof(documents));
        }

        byte[][] bytes = sources.Select(document => document.GetBytesForOperation(cancellationToken)).ToArray();
        PdfLoadOptions[] readOptions = sources.Select(static document => document.ReadOptions).ToArray();
        return PdfMerger.MergeResult(options, bytes, readOptions, cancellationToken);
    }

    /// <summary>
    /// Merges caller-owned PDF byte payloads with an explicit structure policy and returns readback evidence.
    /// The inputs are consumed synchronously and are not retained by the result.
    /// </summary>
    internal static PdfMergeResult MergeBytesResult(PdfMergeOptions options, IEnumerable<byte[]> pdfs) =>
        MergeBytesResult(options, pdfs, CancellationToken.None);

    /// <summary>Merges caller-owned payloads with cancellation through input enumeration, rewriting and readback.</summary>
    internal static PdfMergeResult MergeBytesResult(PdfMergeOptions options, IEnumerable<byte[]> pdfs, CancellationToken cancellationToken) {
        Guard.NotNull(options, nameof(options));
        List<byte[]> sources = CollectMergeByteSources(pdfs, cancellationToken);
        var readOptions = new PdfLoadOptions[sources.Count];
        for (int index = 0; index < readOptions.Length; index++) {
            readOptions[index] = PdfLoadOptions.Default;
        }
        return PdfMerger.MergeResult(options, sources, readOptions, cancellationToken);
    }

    /// <summary>Merges this PDF with another loaded or generated PDF using an explicit structure policy.</summary>
    public PdfDocument MergeWith(PdfDocument document, PdfMergeOptions options) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(options, nameof(options));
        return MergeResult(options, this, document).RequireValue();
    }

    /// <summary>Merges this PDF with another loaded or generated PDF and returns structure and readback evidence.</summary>
    internal PdfMergeResult MergeWithResult(PdfDocument document, PdfMergeOptions options) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(options, nameof(options));
        return MergeResult(options, this, document);
    }

    /// <summary>Merges this PDF with another document using an explicit structure policy and cooperative cancellation.</summary>
    public PdfDocument MergeWith(PdfDocument document, PdfMergeOptions options, CancellationToken cancellationToken) =>
        MergeWithResult(document, options, cancellationToken).RequireValue();

    /// <summary>Merges this PDF with another document with cancellation through rewriting and structure readback.</summary>
    internal PdfMergeResult MergeWithResult(PdfDocument document, PdfMergeOptions options, CancellationToken cancellationToken) {
        Guard.NotNull(document, nameof(document));
        Guard.NotNull(options, nameof(options));
        return MergeResult(options, new[] { this, document }, cancellationToken);
    }
}
