using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Tests;

public sealed partial class PdfWorkspaceTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CanceledDiscardOrUndoPreservesRecoveryAndEditsUntilDeletionCanComplete(bool undo) {
        string root = Path.Combine(Path.GetTempPath(), "studio-workspace-delete-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            CreateEditableSource(source);
            byte[] original = File.ReadAllBytes(source);
            var store = new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery"));
            using var workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None, store);
            await workspace.DuplicateAsync([1], CancellationToken.None);
            byte[] edits = workspace.CopyBytes();
            Task Operation(CancellationToken token) => undo ? workspace.UndoAsync(token) : workspace.DiscardRecoveryAsync(token);
            using (var lease = new FileStream(Path.Combine(root, "recovery", PdfWorkspaceRecoveryStore.LockFileName),
                FileMode.Open, FileAccess.ReadWrite, FileShare.None)) {
                using var cancellation = new CancellationTokenSource();
                Task operation = Operation(cancellation.Token);
                Assert.False(operation.IsCompleted);
                cancellation.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => operation);
                Assert.Equal(edits, workspace.CopyBytes());
                Assert.True(workspace.IsDirty);
                Assert.True(workspace.CanUndo);
                Assert.Equal(edits, store.ReadVerifiedSnapshot(source, workspace.BaseFingerprint));
            }
            await Operation(CancellationToken.None);
            Assert.Null(store.Find(source, workspace.BaseFingerprint));
            Assert.Equal(original, File.ReadAllBytes(source));
            Assert.Equal(undo ? original : edits, workspace.CopyBytes());
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task SaveAsWaitsForRecoveryCleanupAndPreservesOriginalSource() {
        string root = Path.Combine(Path.GetTempPath(), "studio-save-delete-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "source.pdf");
            string destination = Path.Combine(root, "copy.pdf");
            CreateEditableSource(source);
            byte[] original = File.ReadAllBytes(source);
            var store = new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery"));
            using var workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None, store);
            await workspace.DuplicateAsync([1], CancellationToken.None);
            Task save;
            using (var lease = new FileStream(Path.Combine(root, "recovery", PdfWorkspaceRecoveryStore.LockFileName),
                FileMode.Open, FileAccess.ReadWrite, FileShare.None)) {
                save = workspace.SaveAsync(destination, CancellationToken.None);
                Assert.False(save.IsCompleted);
            }
            await save;
            Assert.False(store.HasSnapshotFiles(source));
            Assert.False(workspace.IsDirty);
            Assert.Equal(original, File.ReadAllBytes(source));
            Assert.Equal(2, PdfDocument.Load(destination).Read().Pages.Count);
        } finally { Directory.Delete(root, recursive: true); }
    }
}
