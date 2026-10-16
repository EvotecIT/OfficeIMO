using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Features.Workspace;

internal sealed partial class PdfWorkspace {
    internal bool CanNormalizePageRotation => CanPlan(PdfMutationOperation.NormalizePageRotation);

    internal Task NormalizePageRotationAsync(IReadOnlyList<int> pages, CancellationToken cancellationToken,
        IProgress<PdfWorkspaceProgress>? progress = null) =>
        MutateAsync(PdfWorkspaceOperationKind.NormalizeRotation, $"Normalized rotation on {pages.Count} page(s)", pages,
            document => document.Pages.NormalizeRotation(PdfPageSelection.From(pages.ToArray()), cancellationToken),
            cancellationToken, progress);
}
