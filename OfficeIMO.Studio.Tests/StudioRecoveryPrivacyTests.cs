using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioRecoveryPrivacyTests {
    [Fact]
    public async Task DisablingStorageDrainsWritesAndPreservesExistingCopiesUntilExplicitCleanup() {
        string root = Path.Combine(Path.GetTempPath(), "studio-recovery-policy-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            var store = new PdfWorkspaceRecoveryStore(root);
            string source = Path.Combine(root, "source.pdf");
            await store.WriteAsync(source, "base", [1], 1, CancellationToken.None);
            bool persisted = false;
            Task pendingWrite;
            Task disable;
            using (var lease = new FileStream(Path.Combine(root, PdfWorkspaceRecoveryStore.LockFileName),
                FileMode.Open, FileAccess.ReadWrite, FileShare.None)) {
                pendingWrite = store.WriteAsync(source, "base", [2], 2, CancellationToken.None);
                disable = store.SetPersistenceAsync(false, () => persisted = true);
                Assert.False(pendingWrite.IsCompleted);
                Assert.False(disable.IsCompleted);
                Assert.False(persisted);
            }
            await pendingWrite;
            await disable;
            Assert.True(persisted);
            Assert.Null(await store.WriteAsync(source, "base", [3], 3, CancellationToken.None));
            Assert.Equal(new byte[] { 2 }, store.ReadVerifiedSnapshot(source, "base"));
            Assert.Equal(1, (await store.ClearAllAsync()).RemovedFiles);
            Assert.Null(await store.WriteAsync(source, "base", [4], 4, CancellationToken.None));
            Assert.False(store.HasSnapshotFiles(source));
            await store.SetPersistenceAsync(true, () => { });
            Assert.NotNull(await store.WriteAsync(source, "base", [5], 5, CancellationToken.None));
            Assert.Equal(new byte[] { 5 }, store.ReadVerifiedSnapshot(source, "base"));
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task FailedOrCanceledPreferenceChangeRetainsThePreviousPolicy() {
        string root = Path.Combine(Path.GetTempPath(), "studio-recovery-policy-failure-" + Guid.NewGuid().ToString("N"));
        try {
            var store = new PdfWorkspaceRecoveryStore(root, persistenceEnabled: false);
            await Assert.ThrowsAsync<IOException>(() => store.SetPersistenceAsync(true, () => throw new IOException("Unavailable preferences")));
            using var canceled = new CancellationTokenSource();
            canceled.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => store.SetPersistenceAsync(true,
                () => throw new InvalidOperationException("Must not save a canceled preference"), canceled.Token));
            Assert.Null(await store.WriteAsync(Path.Combine(root, "source.pdf"), "base", [1], 1, CancellationToken.None));
            Assert.False(Directory.Exists(root));
        } finally { if (Directory.Exists(root)) Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task OptOutSurvivesRestartAndStillAllowsEditingUndoRedoAndSave() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Avalonia.Application.Current!).Services;
            using var viewModel = new OfficeIMO.Studio.Features.Shell.MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            await viewModel.Settings.ToggleRecoveryPersistenceCommand.ExecuteAsync(null);
            Assert.False(viewModel.Settings.CreateRecoverySnapshots);
            var restarted = StudioApplicationServices.Create(services.Paths);
            Assert.False(restarted.Preferences.Current.CreateRecoverySnapshots);
            string source = Path.Combine(services.Paths.Root, "private.pdf");
            string saved = Path.Combine(services.Paths.Root, "saved.pdf");
            PdfDocument.Create(builder => builder.Page(page => page.Size(600, 800))).Save(source);
            byte[] original = File.ReadAllBytes(source);
            using var workspace = await PdfWorkspace.OpenAsync(source, CancellationToken.None, restarted.Recovery);
            await workspace.DuplicateAsync([1], CancellationToken.None);
            Assert.True(workspace.IsDirty);
            await workspace.UndoAsync(CancellationToken.None);
            Assert.False(workspace.IsDirty);
            await workspace.RedoAsync(CancellationToken.None);
            Assert.True(workspace.IsDirty);
            Assert.False(restarted.Recovery.HasSnapshotFiles(source));
            await workspace.SaveAsync(saved, CancellationToken.None);
            Assert.Equal(2, PdfDocument.Load(saved).Read().Pages.Count);
            Assert.Equal(original, File.ReadAllBytes(source));
            return true;
        }, CancellationToken.None);
    }
}
