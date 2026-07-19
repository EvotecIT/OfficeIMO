namespace OfficeIMO.Pdf;

/// <summary>Result of a dependency-free PDF annotation edit operation.</summary>
public sealed class PdfAnnotationEditResult {
    private readonly byte[] _bytes;
    private readonly PdfReadOptions? _readOptions;

    internal PdfAnnotationEditResult(
        byte[] bytes,
        int affectedAnnotationCount,
        PdfMutationPlan mutationPlan,
        PdfSignatureMutationReport? signatureMutationReport = null,
        PdfRewritePreservationReport? rewritePreservationReport = null,
        PdfReadOptions? readOptions = null) {
        _bytes = (byte[])bytes.Clone();
        _readOptions = PdfReadOptions.WithMinimumInputBytes(readOptions, _bytes.LongLength);
        AffectedAnnotationCount = affectedAnnotationCount;
        MutationPlan = mutationPlan;
        SignatureMutationReport = signatureMutationReport;
        RewritePreservationReport = rewritePreservationReport;
    }

    /// <summary>Rewritten PDF bytes.</summary>
    public byte[] Bytes => (byte[])_bytes.Clone();

    /// <summary>Number of annotations removed or updated.</summary>
    public int AffectedAnnotationCount { get; }

    /// <summary>Shared mutation decision used by the editor.</summary>
    public PdfMutationPlan MutationPlan { get; }

    /// <summary>Append-only signature and revision proof, when append-only mode was selected.</summary>
    public PdfSignatureMutationReport? SignatureMutationReport { get; }

    /// <summary>Full-rewrite preservation proof, when full rewrite mode was selected.</summary>
    public PdfRewritePreservationReport? RewritePreservationReport { get; }

    /// <summary>True when the operation changed at least one annotation.</summary>
    public bool Applied => AffectedAnnotationCount > 0;

    /// <summary>Opens the edited bytes through the fluent document API.</summary>
    public PdfDocument ToDocument(PdfReadOptions? readOptions = null) => PdfDocument.Open(_bytes, readOptions ?? _readOptions);
}
