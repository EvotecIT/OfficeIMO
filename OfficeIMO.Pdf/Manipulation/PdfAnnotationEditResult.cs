namespace OfficeIMO.Pdf;

/// <summary>Result of a dependency-free PDF annotation edit operation.</summary>
public sealed class PdfAnnotationEditResult {
    private readonly byte[] _bytes;
    private readonly PdfLoadOptions? _readOptions;

    internal PdfAnnotationEditResult(
        byte[] bytes,
        int affectedAnnotationCount,
        PdfMutationPlan mutationPlan,
        PdfSignatureMutationReport? signatureMutationReport = null,
        PdfRewritePreservationReport? rewritePreservationReport = null,
        PdfLoadOptions? readOptions = null,
        IReadOnlyDictionary<int, int>? annotationObjectNumberMap = null) {
        _bytes = (byte[])bytes.Clone();
        _readOptions = PdfLoadOptions.WithMinimumInputBytes(readOptions, _bytes.LongLength);
        AffectedAnnotationCount = affectedAnnotationCount;
        MutationPlan = mutationPlan;
        SignatureMutationReport = signatureMutationReport;
        RewritePreservationReport = rewritePreservationReport;
        AnnotationObjectNumberMap = annotationObjectNumberMap is null ? null : new System.Collections.ObjectModel.ReadOnlyDictionary<int, int>(annotationObjectNumberMap.ToDictionary(pair => pair.Key, pair => pair.Value));
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

    /// <summary>Original-to-output annotation object numbers reported by the full rewrite, or null when no map is available.</summary>
    /// <remarks>Removed annotations are absent. Append-only mutations preserve existing object numbers.</remarks>
    public IReadOnlyDictionary<int, int>? AnnotationObjectNumberMap { get; }

    /// <summary>True when the operation changed at least one annotation.</summary>
    public bool Applied => AffectedAnnotationCount > 0;

    /// <summary>Opens the edited bytes through the fluent document API.</summary>
    public PdfDocument ToDocument(PdfLoadOptions? readOptions = null) => PdfDocument.Load(_bytes, readOptions ?? _readOptions);

    internal PdfLoadOptions OutputReadOptions => _readOptions!;
}
