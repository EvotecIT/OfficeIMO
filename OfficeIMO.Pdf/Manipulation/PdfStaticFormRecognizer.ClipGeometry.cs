namespace OfficeIMO.Pdf;

internal static partial class PdfStaticFormRecognizer {
    private static bool ClipProvesSeparate(PdfPageClipPath clip, PdfPageClipPath target,
        ref long candidateScanWork, int maxCandidateScanWork,
        PdfReadPage.VisualGeometryBudget geometryBudget) {
        long previousWork = geometryBudget.TotalWork;
        bool separate = clip.CanProveNoPositiveAreaIntersection(target, geometryBudget);
        candidateScanWork += geometryBudget.TotalWork - previousWork;
        if (geometryBudget.Exceeded || candidateScanWork > maxCandidateScanWork) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.UnderstandingArtifacts,
                maxCandidateScanWork, Math.Max(candidateScanWork, (long)maxCandidateScanWork + 1L));
        }
        return separate;
    }
}
