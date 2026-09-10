using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private async Task<PdfDocument> ReadPrintSnapshotAsync(string path, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        if (_workspace is { } workspace && OfficeIMO.Internal.OfficeStorageIdentity.AreEquivalent(path, workspace.Path)) {
            if (HasFormDrafts) throw new InvalidOperationException(_localizer.GetOrDefault("PrintPreview.FormDrafts", "Apply pending form values before preparing print sheets."));
            return workspace.CreateDocumentSnapshot();
        }
        using PdfWorkspace? opened = await OpenWorkspaceWithPasswordAsync(path, token).ConfigureAwait(true);
        return opened?.CreateDocumentSnapshot() ?? throw new OperationCanceledException(token);
    }
}
