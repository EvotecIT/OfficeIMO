using System.Reflection;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Platform.Storage;
using Avalonia.VisualTree;
using OfficeIMO.Drawing;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioProviderOutputFolderTests {
    [Fact]
    public async Task MultipleRecoveryCopiesStaySelectableAndDiscardConsentTracksTheSelection() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf");
            File.WriteAllBytes(source, StudioProviderDocumentTests.CreatePdf(2));
            var recoveries = new List<OfficeIMO.Workflows.OfficeWorkflowOutputRecovery>();
            foreach (string name in new[] { "First.pdf", "Second.pdf" }) {
                var file = new TestStorageFile("content://recoveries/" + name, [], name) { FailWrite = true };
                var result = await new OfficeIMO.Workflows.OfficeWorkflowRunner().AssemblePdfAsync(new() {
                    Sources = [source], OutputPath = file.Location.AbsoluteUri, ConflictPolicy = OfficeIMO.OfficePublicationPolicy.Replace,
                    OutputStream = new(name, _ => file.Item.OpenReadAsync(), _ => file.Item.OpenWriteAsync(), services.WorkflowRecovery)
                });
                recoveries.Add(Assert.IsType<OfficeIMO.Workflows.OfficeWorkflowOutputRecovery>(result.Recovery));
            }
            var job = services.Jobs.Start("Page image export", source, "content://folder/selected", () => { });

            using var jobs = new StudioJobsViewModel(services.Jobs, (_, _) => Task.CompletedTask);
            var window = new Window { Width = 960, Height = 620, Content = new StudioJobsView { DataContext = jobs } };
            try {
                window.Show(); window.UpdateLayout();
                job.CompleteBatch(OfficeIMO.Workflows.OfficeWorkflowStatus.Unconfirmed, null, "Two recovery copies are available.", recoveries, false, isDirectoryOutput: true);
                window.UpdateLayout();
                Assert.Same(recoveries[0], job.Recovery);
                Assert.True(job.HasMultipleRecoveries);
                Capture(window, "provider-folder-multiple-recoveries.png");
                jobs.RequestDiscardRecoveryCommand.Execute(job);
                job.Recovery = recoveries[1];
                Assert.False(job.RecoveryDiscardRequested);
                jobs.ConfirmDiscardRecoveryCommand.Execute(job);
                Assert.All(recoveries, recovery => Assert.True(File.Exists(recovery.FilePath)));
                jobs.RequestDiscardRecoveryCommand.Execute(job);
                jobs.ConfirmDiscardRecoveryCommand.Execute(job);
                Assert.False(File.Exists(recoveries[1].FilePath));
                Assert.True(File.Exists(recoveries[0].FilePath));
                Assert.Same(recoveries[0], job.Recovery);
                Assert.False(job.HasMultipleRecoveries);
                await jobs.OpenRecoveryCommand.ExecuteAsync(job);
                Assert.Null(jobs.ActionError);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ConversionAndHealthRejectProviderOutputFoldersBeforeWriting(bool health) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, health ? "source.pdf" : "source.html");
            File.WriteAllBytes(source, health ? StudioProviderDocumentTests.CreatePdf(2)
                : System.Text.Encoding.UTF8.GetBytes("<html><body><p>Folder output</p></body></html>"));
            byte[] original = File.ReadAllBytes(source);
            var folder = new OutputFolder();
            string destination = (await services.Storage.RegisterFolderAsync([folder.Item], default))!;
            using var shell = new MainWindowViewModel(_ => Task.FromResult<string?>(source), services: services,
                pickWorkflowFiles: _ => Task.FromResult<IReadOnlyList<string>>([source]));
            string status;
            if (health) {
                var model = shell.DocumentHealth;
                await model.ChooseInputCommand.ExecuteAsync(null);
                model.PrepareWorkflow(OfficeIMO.Workflows.OfficeWorkflowOperation.Optimize);
                model.OutputFolder = destination;
                await model.RunCommand.ExecuteAsync(null);
                status = model.Status;
            } else {
                var model = shell.ConversionWorkbench;
                model.SelectedRoute = model.Routes.Single(route => route.Route.Id == "html-pdf");
                await model.AddFilesCommand.ExecuteAsync(null);
                model.OutputFolder = destination;
                await model.RunQueueCommand.ExecuteAsync(null);
                status = model.Status;
            }
            Assert.Contains(services.Localizer.Get("Output.LocalFolderRequired"), status);
            Assert.Equal(0, folder.Creations);
            Assert.Empty(folder.Files);
            Assert.Empty(services.Jobs.Entries);
            Assert.Equal(original, File.ReadAllBytes(source));
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task PageExportRejectsProviderFoldersWithALocalizedNextAction() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf");
            File.WriteAllBytes(source, StudioProviderDocumentTests.CreatePdf(2));
            var folder = new OutputFolder();
            string destination = (await services.Storage.RegisterFolderAsync([folder.Item], default))!;
            using var shell = new MainWindowViewModel(_ => Task.FromResult<string?>(source), services: services,
                pickOutputFolder: _ => Task.FromResult<string?>(destination));
            shell.OutputWorkbench.SelectedSection = OutputWorkbenchSection.ExportPages;
            var model = shell.OutputWorkbench.PageExport;
            model.UseDocument(source);
            await model.ChooseOutputDirectoryCommand.ExecuteAsync(null);
            await model.ExportCommand.ExecuteAsync(null);
            Assert.Equal(services.Localizer.Get("Output.LocalFolderRequired"), model.Status);
            Assert.Equal(0, folder.Creations);
            Assert.Empty(folder.Files);
            Assert.Empty(services.Jobs.Entries);
            Assert.False(model.HasOutput);
            var window = new Window { Width = 960, Height = 620, Content = new OutputIntakeWorkbenchView { DataContext = shell } };
            try {
                window.Show(); window.UpdateLayout();
                Assert.Contains(window.GetVisualDescendants().OfType<TextBlock>(), text => text.IsEffectivelyVisible && text.Text == model.Status);
                Capture(window, "local-folder-output-required.png");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static void Capture(Window window, string name) {
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? directory = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(directory)) return;
        Directory.CreateDirectory(directory);
        frame.Save(Path.Combine(directory, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }

    internal sealed class OutputFolder {
        internal IStorageFolder Item { get; }
        internal Dictionary<string, TestStorageFile> Files { get; } = new(StringComparer.Ordinal);
        internal int Creations;

        internal OutputFolder() {
            Item = DispatchProxy.Create<IStorageFolder, TestStorageFile.StorageProxy>();
            ((TestStorageFile.StorageProxy)(object)Item).Call = (method, args) => method switch {
                "get_Name" => "Selected output folder", "get_Path" => new Uri("content://folder/selected"), "get_CanBookmark" => false,
                "GetFileAsync" => Task.FromResult<IStorageFile?>(Files.TryGetValue((string)args![0]!, out var file) ? file.Item : null),
                "CreateFileAsync" => Create((string)args![0]!), "Dispose" => null,
                _ => throw new NotSupportedException(method)
            };
        }

        private Task<IStorageFile?> Create(string name) {
            Creations++;
            throw new IOException("Provider child creation is unavailable.");
        }
    }
}
