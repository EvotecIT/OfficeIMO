namespace OfficeIMO.Pdf;

/// <summary>Result of a dependency-free PDF annotation edit operation.</summary>
public sealed class PdfAnnotationEditResult {
    internal PdfAnnotationEditResult(byte[] bytes, int affectedAnnotationCount) {
        Bytes = bytes;
        AffectedAnnotationCount = affectedAnnotationCount;
    }

    /// <summary>Rewritten PDF bytes.</summary>
    public byte[] Bytes { get; }

    /// <summary>Number of annotations removed or updated.</summary>
    public int AffectedAnnotationCount { get; }

    /// <summary>True when the operation changed at least one annotation.</summary>
    public bool Applied => AffectedAnnotationCount > 0;
}
