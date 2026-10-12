using Avalonia;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

/// <summary>
/// Deliberate failures: a full disk, a read-only destination and damaged recovery records. Each keeps the source and
/// the user's work, states what happened to the output, and leaves no staging files behind.
/// </summary>
public sealed class StudioFailureScenarioTests {
    private static Exception DiskFull() => new IOException("There is not enough space on the disk.", unchecked((int)0x80070070));

    [Fact]
    public async Task Saving_to_a_full_disk_keeps_the_edits_and_a_recovery_copy_and_says_the_disk_is_full() {
        using var storage = new StudioStorageAccess();
        var source = new TestStorageFile("content://documents/full-disk", StudioProviderDocumentTests.CreatePdf()) { FailWrite = true, WriteFailure = DiskFull };
        string location = await storage.RegisterAsync(source.Item, default);
        using var root = new TestDirectory();
        using var workspace = await PdfWorkspace.OpenAsync(location, default, new(root.Path), storage: storage);
        await workspace.DuplicateAsync([1], default);

        IOException error = await Assert.ThrowsAsync<IOException>(() => workspace.SaveAsync(null, default));

        Assert.Equal(StudioLocalization.Current.Get("Errors.DiskFull"), StudioMessages.Describe(error));
        Assert.True(workspace.IsDirty);
        Assert.NotNull(new PdfWorkspaceRecoveryStore(root.Path).ReadVerifiedSnapshot(location, workspace.BaseFingerprint));
        Assert.Equal(2, workspace.CreateDocumentSnapshot().Inspect().Pages.Count);
    }

    [Fact]
    public async Task A_workflow_output_that_meets_a_full_disk_needs_checking_and_keeps_a_recovery_copy() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf");
            File.WriteAllBytes(source, StudioProviderDocumentTests.CreatePdf(2));
            var target = new TestStorageFile("content://documents/full-output", [], "Assembled report.pdf") { FailWrite = true, WriteFailure = DiskFull };
            string destination = await services.Storage.RegisterAsync(target.Item, default);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                confirmProviderWrite: _ => Task.FromResult(true));
            var assembly = model.OutputWorkbench.Assembly;
            assembly.UseDocument(source);
            assembly.OutputPath = destination;

            await assembly.RunCommand.ExecuteAsync(null);

            Assert.False(assembly.HasOutput);
            Assert.True(assembly.HasRecovery, assembly.Summary);
            StudioJobRecord job = Assert.Single(services.Jobs.Entries);
            Assert.True(job.NeedsChecking, job.Status);
            Assert.True(File.Exists(job.Recovery!.FilePath));
            Assert.Equal(2, PdfDocument.Load(job.Recovery.FilePath).Inspect().PageCount);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task Saving_a_copy_over_a_read_only_file_leaves_it_untouched_and_keeps_the_edits() {
        // Windows enforces the read-only attribute on replacement; POSIX renames over read-only files by design.
        if (!OperatingSystem.IsWindows()) return;
        using var root = new TestDirectory();
        string source = Path.Combine(root.Path, "source.pdf");
        File.WriteAllBytes(source, StudioProviderDocumentTests.CreatePdf());
        string locked = Path.Combine(root.Path, "locked.pdf");
        byte[] lockedBytes = StudioProviderDocumentTests.CreatePdf(3);
        File.WriteAllBytes(locked, lockedBytes);
        File.SetAttributes(locked, FileAttributes.ReadOnly);
        try {
            using var workspace = await PdfWorkspace.OpenAsync(source, default, new(Path.Combine(root.Path, "recovery")));
            await workspace.DuplicateAsync([1], default);
            string[] before = Directory.GetFiles(root.Path);

            Exception error = await Assert.ThrowsAnyAsync<Exception>(() => workspace.SaveAsync(locked, default));

            Assert.Equal(StudioLocalization.Current.Get("Errors.ReadOnly"), StudioMessages.Describe(error));
            Assert.True(workspace.IsDirty);
            Assert.Equal(source, workspace.Path);
            Assert.Equal(lockedBytes, File.ReadAllBytes(locked));
            Assert.Equal(before.Order(), Directory.GetFiles(root.Path).Order());
        } finally {
            File.SetAttributes(locked, FileAttributes.Normal);
        }
    }

    [Fact]
    public async Task A_conversion_that_may_not_replace_a_read_only_output_fails_without_touching_it() {
        if (!OperatingSystem.IsWindows()) return;
        using var root = new TestDirectory();
        string input = Path.Combine(root.Path, "notes.txt");
        File.WriteAllText(input, "Quarterly notes");
        string output = Path.Combine(root.Path, "notes.pdf");
        byte[] outputBytes = StudioProviderDocumentTests.CreatePdf();
        File.WriteAllBytes(output, outputBytes);
        File.SetAttributes(output, FileAttributes.ReadOnly);
        try {
            IReadOnlyList<string> files = [input];
            using var model = new ConversionWorkbenchViewModel(_ => Task.FromResult(files), _ => Task.FromResult<string?>(root.Path), new OfficeWorkflowRunner());
            model.SelectedRoute = model.Routes.Single(route => route.Route.Id == "txt-pdf");
            model.SelectedConflict = model.ConflictPolicies.Single(choice => choice.Value == OfficeWorkflowConflictPolicy.Replace);
            await model.AddFilesCommand.ExecuteAsync(null);
            string[] before = Directory.GetFiles(root.Path);

            await model.RunQueueCommand.ExecuteAsync(null);

            ConversionJobViewModel job = Assert.Single(model.Jobs);
            Assert.True(job.IsFailed, job.Status + ": " + job.Summary);
            Assert.False(job.HasOutput);
            Assert.Equal(outputBytes, File.ReadAllBytes(output));
            Assert.Equal(before.Order(), Directory.GetFiles(root.Path).Order());
        } finally {
            File.SetAttributes(output, FileAttributes.Normal);
        }
    }

    [Fact]
    public async Task Damaged_workflow_recovery_records_are_skipped_while_a_real_one_returns_after_restart() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf");
            File.WriteAllBytes(source, StudioProviderDocumentTests.CreatePdf(2));
            var target = new TestStorageFile("content://documents/interrupted", [], "Interrupted report.pdf") { FailWrite = true };
            string destination = await services.Storage.RegisterAsync(target.Item, default);
            using (var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                       confirmProviderWrite: _ => Task.FromResult(true))) {
                model.OutputWorkbench.Assembly.UseDocument(source);
                model.OutputWorkbench.Assembly.OutputPath = destination;
                await model.OutputWorkbench.Assembly.RunCommand.ExecuteAsync(null);
            }
            string realId = Assert.Single(services.Jobs.Entries).Recovery!.Id;

            // Damaged records beside the real one: unreadable JSON, a record pointing at a missing output, an empty folder.
            string recoveryRoot = services.Paths.WorkflowRecoveryRoot;
            string broken = Directory.CreateDirectory(Path.Combine(recoveryRoot, Guid.NewGuid().ToString("N"))).FullName;
            File.WriteAllText(Path.Combine(broken, "record.json"), "{ broken");
            string orphan = Directory.CreateDirectory(Path.Combine(recoveryRoot, Guid.NewGuid().ToString("N"))).FullName;
            File.WriteAllText(Path.Combine(orphan, "record.json"), "{\"Id\":\"" + Path.GetFileName(orphan) + "\",\"FileName\":\"output.pdf\"}");
            Directory.CreateDirectory(Path.Combine(recoveryRoot, Guid.NewGuid().ToString("N")));

            var restarted = StudioApplicationServices.Create(services.Paths);

            StudioJobRecord recovered = Assert.Single(restarted.Jobs.Entries);
            Assert.Equal(realId, recovered.Recovery!.Id);
            Assert.Equal(2, PdfDocument.Load(recovered.Recovery.FilePath).Inspect().PageCount);
            return true;
        }, CancellationToken.None);
    }

    private sealed class TestDirectory : IDisposable {
        internal string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "studio-failure-tests-" + Guid.NewGuid().ToString("N"));
        internal TestDirectory() => Directory.CreateDirectory(Path);
        public void Dispose() => Directory.Delete(Path, true);
    }
}
