using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Pdf;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfWorkspaceTests {
    [Fact]
    public async Task MutationUndoRedoAndSaveMaintainWorkspaceContract() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-workspace-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        string saved = Path.Combine(root, "saved.pdf");
        CreateEditableSource(source);
        var recovery = new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery"));

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None, recovery);

            await workspace.DuplicateAsync([1], CancellationToken.None);
            Assert.True(workspace.IsDirty);
            Assert.True(workspace.CanUndo);
            Assert.Equal(2, workspace.Pages.Count);
            Assert.False(workspace.HasRecovery);

            await workspace.UndoAsync(CancellationToken.None);
            Assert.Single(workspace.Pages);
            Assert.True(workspace.CanRedo);

            await workspace.RedoAsync(CancellationToken.None);
            Assert.Equal(2, workspace.Pages.Count);

            await workspace.SaveAsync(saved, CancellationToken.None);
            Assert.False(workspace.IsDirty);
            Assert.Equal(saved, workspace.Path);
            Assert.True(File.Exists(saved));
            Assert.False(workspace.HasRecovery);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ReorderRotateCropImportAndBlankProduceReadableArtifacts() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-organizer-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        CreateEditableSource(source);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(
                source,
                CancellationToken.None,
                new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery")));

            await workspace.DuplicateAsync([1], CancellationToken.None);
            await workspace.ReorderAsync([2, 1], CancellationToken.None);
            await workspace.RotateAsync([1], 90, CancellationToken.None);
            await workspace.CropAsync([2], 10, 10, 580, 760, CancellationToken.None);
            await workspace.ImportAsync(source, 2, CancellationToken.None);
            await workspace.InsertBlankAsync(workspace.Pages.Count + 1, 400, 500, CancellationToken.None);

            Assert.Equal(4, workspace.Pages.Count);
            Assert.Equal(90, workspace.Pages[0].RotationDegrees);
            Assert.Equal(6, workspace.Journal.Count);
            Assert.NotEmpty(workspace.CopyBytes());
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task RecoveryCanBeDiscoveredRestoredAndDiscarded() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-recovery-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        CreateEditableSource(source);
        var recovery = new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery"));

        try {
            using (PdfWorkspace edited = await PdfWorkspace.OpenAsync(source, CancellationToken.None, recovery)) {
                await edited.DuplicateAsync([1], CancellationToken.None);
                Assert.False(edited.HasRecovery);
            }

            using PdfWorkspace reopened = await PdfWorkspace.OpenAsync(source, CancellationToken.None, recovery);
            Assert.True(reopened.HasRecovery);
            Assert.Single(reopened.Pages);

            await reopened.RestoreRecoveryAsync(CancellationToken.None);

            Assert.Equal(2, reopened.Pages.Count);
            Assert.True(reopened.IsDirty);
            reopened.DiscardRecovery();
            Assert.False(reopened.HasRecovery);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task TaggedPdfReportsPageMutationAsUnavailableBeforeAnEditStarts() {
        string fixture = Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf");
        using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(fixture, CancellationToken.None);

        Assert.False(workspace.CanMutatePages);
        Assert.NotNull(workspace.SecurityWarning);
        Assert.Contains("tagged", workspace.SecurityWarning, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task FailedRecoveryWriteDoesNotCommitAnInMemoryMutation() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-transaction-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        string blockedRecoveryRoot = Path.Combine(root, "not-a-directory");
        CreateEditableSource(source);
        await File.WriteAllTextAsync(blockedRecoveryRoot, "block directory creation");

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(
                source,
                CancellationToken.None,
                new PdfWorkspaceRecoveryStore(blockedRecoveryRoot));
            byte[] original = workspace.CopyBytes();

            await Assert.ThrowsAnyAsync<IOException>(() => workspace.DuplicateAsync([1], CancellationToken.None));

            Assert.Single(workspace.Pages);
            Assert.False(workspace.IsDirty);
            Assert.False(workspace.CanUndo);
            Assert.Equal(original, workspace.CopyBytes());
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task RecoveryIsRejectedWhenTheSourceAtTheSamePathWasReplaced() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-recovery-identity-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        var recovery = new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery"));
        CreateEditableSource(source);

        try {
            using (PdfWorkspace edited = await PdfWorkspace.OpenAsync(source, CancellationToken.None, recovery)) {
                await edited.DuplicateAsync([1], CancellationToken.None);
            }
            PdfDocument.Create(compose => compose.Page(page => page.Size(300, 400))).Save(source);

            using PdfWorkspace replaced = await PdfWorkspace.OpenAsync(source, CancellationToken.None, recovery);

            Assert.False(replaced.HasRecovery);
            Assert.Single(replaced.Pages);
            Assert.Equal(300D, replaced.Pages[0].Width);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task ExtractRejectsTheOpenDocumentPathWithoutChangingTheSource() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-extract-identity-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        CreateEditableSource(source);
        byte[] original = await File.ReadAllBytesAsync(source);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);

            InvalidOperationException exception = await Assert.ThrowsAsync<InvalidOperationException>(
                () => workspace.ExtractAsync([1], source, CancellationToken.None));

            Assert.Contains("different file", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(original, await File.ReadAllBytesAsync(source));
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task NewEditAfterUndoCannotReuseTheSavedRevisionIdentity() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-revision-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        CreateEditableSource(source);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None);
            await workspace.DuplicateAsync([1], CancellationToken.None);
            await workspace.SaveAsync(path: null, CancellationToken.None);
            await workspace.DuplicateAsync([1], CancellationToken.None);
            await workspace.UndoAsync(CancellationToken.None);
            Assert.False(workspace.IsDirty);

            await workspace.InsertBlankAsync(1, 200, 300, CancellationToken.None);

            Assert.True(workspace.IsDirty);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public async Task FailedRecoveryWriteDuringRedoLeavesHistoryAndDocumentConsistent() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-redo-transaction-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        string source = Path.Combine(root, "source.pdf");
        string recoveryRoot = Path.Combine(root, "recovery");
        CreateEditableSource(source);

        try {
            using PdfWorkspace workspace = await PdfWorkspace.OpenAsync(
                source,
                CancellationToken.None,
                new PdfWorkspaceRecoveryStore(recoveryRoot));
            await workspace.DuplicateAsync([1], CancellationToken.None);
            await workspace.UndoAsync(CancellationToken.None);
            Directory.Delete(recoveryRoot, recursive: true);
            await File.WriteAllTextAsync(recoveryRoot, "block directory creation");

            await Assert.ThrowsAnyAsync<IOException>(() => workspace.RedoAsync(CancellationToken.None));

            Assert.Single(workspace.Pages);
            Assert.False(workspace.IsDirty);
            Assert.True(workspace.CanRedo);
            Assert.False(workspace.CanUndo);
        } finally {
            Directory.Delete(root, recursive: true);
        }
    }

    private static void CreateEditableSource(string path) =>
        PdfDocument.Create(compose => compose.Page(page => page.Size(600, 800))).Save(path);
}
