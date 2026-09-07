using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioProviderOutputTests {
    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 820, true)]
    public async Task FailedAssemblyPreservesRestartRecoveryAndRequiresExplicitDiscard(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            string source = Path.Combine(services.Paths.Root, "source.pdf");
            File.WriteAllBytes(source, StudioProviderDocumentTests.CreatePdf(2));
            var target = new TestStorageFile("content://documents/output", [], "Assembled provider report.pdf") { FailWrite = true };
            string destination = await services.Storage.RegisterAsync(target.Item, default);
            int confirmations = 0;
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                confirmProviderWrite: location => { Assert.Equal(destination, location); confirmations++; return Task.FromResult(true); });
            var assembly = model.OutputWorkbench.Assembly;
            assembly.UseDocument(source);
            assembly.OutputPath = destination;
            await assembly.RunCommand.ExecuteAsync(null);
            Assert.Equal(1, confirmations);
            Assert.Equal(1, target.Writes);
            Assert.False(assembly.HasOutput);
            Assert.True(assembly.HasRecovery, assembly.Summary);
            var original = Assert.Single(services.Jobs.Entries);
            Assert.Equal("Check output", original.Status);
            services.Jobs.ClearFinished();
            Assert.Same(original, Assert.Single(services.Jobs.Entries));

            var restarted = StudioApplicationServices.Create(services.Paths);
            var recovered = Assert.Single(restarted.Jobs.Entries);
            Assert.Equal(original.Recovery!.Id, recovered.Recovery!.Id);
            Assert.False(recovered.IsActive);
            byte[] retainedBytes = File.ReadAllBytes(recovered.Recovery.FilePath);
            using (var workspace = await PdfWorkspace.OpenAsync(recovered.Recovery.FilePath, default,
                       recoveryStore: restarted.Recovery, storage: restarted.Storage)) {
                await workspace.RotateAsync([1], 90, default);
                await Assert.ThrowsAsync<IOException>(() => workspace.SaveAsync(null, default));
                string savedCopy = Path.Combine(services.Paths.Root, "Recovered document.pdf");
                await workspace.SaveAsync(savedCopy, default);
                Assert.Equal(2, PdfDocument.Load(savedCopy).Inspect().PageCount);
            }
            Assert.Equal(retainedBytes, File.ReadAllBytes(recovered.Recovery.FilePath));
            string? opened = null;
            using var jobs = new StudioJobsViewModel(restarted.Jobs, (path, _) => {
                Assert.Equal(2, PdfDocument.Load(path).Inspect().PageCount);
                opened = path;
                return Task.CompletedTask;
            });
            var window = new Window { Width = width, Height = height, Content = new StudioJobsView { DataContext = jobs } };
            try {
                window.Show();
                window.UpdateLayout();
                await jobs.OpenRecoveryCommand.ExecuteAsync(recovered);
                Assert.Null(jobs.ActionError);
                Assert.Equal(recovered.Recovery.FilePath, opened);
                jobs.ConfirmDiscardRecoveryCommand.Execute(recovered);
                Assert.True(File.Exists(opened));
                jobs.RequestDiscardRecoveryCommand.Execute(recovered);
                window.UpdateLayout();
                var confirm = window.GetVisualDescendants().OfType<Button>()
                    .Single(button => ReferenceEquals(button.Command, jobs.ConfirmDiscardRecoveryCommand));
                Assert.True(confirm.IsEffectivelyVisible);
                Point position = confirm.TranslatePoint(default, window)!.Value;
                Assert.InRange(position.X, 0, window.Bounds.Width - confirm.Bounds.Width);
                Assert.InRange(position.Y, 0, window.Bounds.Height - confirm.Bounds.Height);
                Capture(window, $"provider-output-recovery-{width}-{(dark ? "dark" : "light")}.png");
                jobs.CancelDiscardRecoveryCommand.Execute(recovered);
                Assert.True(File.Exists(opened));
                jobs.RequestDiscardRecoveryCommand.Execute(recovered);
                jobs.ConfirmDiscardRecoveryCommand.Execute(recovered);
                Assert.Null(jobs.ActionError);
                Assert.False(File.Exists(opened));
                Assert.False(recovered.HasRecovery);
                Assert.Empty(restarted.WorkflowRecovery.GetRecoveries());
                Assert.Equal(1, target.Writes);
                jobs.ClearFinishedCommand.Execute(null);
                Assert.Empty(restarted.Jobs.Entries);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task AssemblyWritesOnlyAfterConsentAndReopensVerifiedProviderOutput(bool consent) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf");
            File.WriteAllBytes(source, StudioProviderDocumentTests.CreatePdf(2));
            var target = new TestStorageFile("content://documents/opaque-output", [], "Combined.pdf");
            string destination = await services.Storage.RegisterAsync(target.Item, default);
            string? opened = null;
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                confirmProviderWrite: _ => Task.FromResult(consent),
                openDocumentInTab: (path, _) => { opened = path; return Task.CompletedTask; });
            model.OutputWorkbench.Assembly.UseDocument(source);
            model.OutputWorkbench.Assembly.OutputPath = destination;
            await model.OutputWorkbench.Assembly.RunCommand.ExecuteAsync(null);
            Assert.Equal(consent ? 1 : 0, target.Writes);
            Assert.Empty(services.WorkflowRecovery.GetRecoveries());
            if (consent) {
                var job = Assert.Single(services.Jobs.Entries);
                Assert.True(job.HasOutput, job.Summary);
                Assert.Equal(2, PdfDocument.Load(target.Bytes).Inspect().PageCount);
                await model.Jobs.OpenOutputCommand.ExecuteAsync(job);
                Assert.Null(model.Jobs.ActionError);
                Assert.Equal(destination, opened);
            } else Assert.Empty(services.Jobs.Entries);
            return true;
        }, CancellationToken.None);
    }

    private static void Capture(Window window, string name) {
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        frame.Save(Path.Combine(output, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
