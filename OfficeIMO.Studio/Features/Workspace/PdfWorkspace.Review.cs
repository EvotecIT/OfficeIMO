using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    internal Task SetAnnotationReviewStateAsync(int objectNumber, PdfAnnotationReviewState state,
        CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateBytesAsync(PdfWorkspaceOperationKind.Annotation, "Updated comment review state", [],
            bytes => LoadDocument(bytes).Annotations.SetReviewState(objectNumber, state).Bytes,
            cancellationToken, progress);
}
