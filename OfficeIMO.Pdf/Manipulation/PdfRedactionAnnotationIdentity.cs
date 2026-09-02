namespace OfficeIMO.Pdf;

/// <summary>Creates a bounded, object-number-independent identity for an annotation dictionary graph.</summary>
internal static class PdfRedactionAnnotationIdentity {
    private const int MaximumDepth = 64;
    private const int MaximumNodes = 16_384;

    internal static string Compute(
        PdfDictionary annotation,
        Dictionary<int, PdfIndirectObject> objects,
        IReadOnlyDictionary<int, string> stableReferenceLabels) {
        using var fingerprint = new PdfObjectGraphFingerprint(
            objects,
            MaximumDepth,
            MaximumNodes,
            stableReferenceLabels: stableReferenceLabels);
        fingerprint.AppendRoot(annotation);
        return Convert.ToBase64String(fingerprint.Complete());
    }
}
