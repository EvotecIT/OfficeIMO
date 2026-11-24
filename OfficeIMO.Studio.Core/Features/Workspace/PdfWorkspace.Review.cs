using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    /// <summary>
    /// Resolves or reopens a comment. A document that takes annotation changes as an appended revision, such as a tagged
    /// PDF or one saved incrementally, keeps the comment's earlier state in the revision before; that is the comment's
    /// history rather than removed content, so the appended revision is allowed, as a reply is.
    /// </summary>
    internal Task SetAnnotationReviewStateAsync(int objectNumber, PdfAnnotationReviewState state,
        CancellationToken cancellationToken, IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateAnnotationBytesAsync(PdfWorkspaceOperationKind.Annotation, "Updated comment review state", [],
            bytes => LoadDocument(bytes).Annotations.SetReviewState(objectNumber, state, PdfMutationExecutionPreference.Automatic,
                allowResidualDataInAppendOnly: true),
            cancellationToken, progress);
}
