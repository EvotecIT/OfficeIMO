using System.Reflection;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Platform.Storage;
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
                    Sources = [source], OutputPath = file.Location.AbsoluteUri, ConflictPolicy = OfficeIMO.Workflows.OfficeWorkflowConflictPolicy.Replace,
                    OutputStream = new(name, _ => file.Item.OpenReadAsync(), _ => file.Item.OpenWriteAsync(), services.WorkflowRecovery)
                });
                recoveries.Add(Assert.IsType<OfficeIMO.Workflows.OfficeWorkflowOutputRecovery>(result.Recovery));
            }
            var job = services.Jobs.Start("Page image export", source, "content://folder/selected", () => { });
            job.CompleteBatch(OfficeIMO.Workflows.OfficeWorkflowStatus.Unconfirmed, null, "Two recovery copies are available.", recoveries, false);
            using var jobs = new StudioJobsViewModel(services.Jobs, (_, _) => Task.CompletedTask);
            var window = new Window { Width = 960, Height = 620, Content = new StudioJobsView { DataContext = jobs } };
            try {
                window.Show(); window.UpdateLayout();
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
    public async Task ConversionAndHealthUseTheSameProviderFolderAccess(bool health) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, health ? "source.pdf" : "source.html");
            File.WriteAllBytes(source, health ? StudioProviderDocumentTests.CreatePdf(2)
                : System.Text.Encoding.UTF8.GetBytes("<html><body><p>Provider folder conversion</p></body></html>"));
            var folder = new OutputFolder();
            string? destination = await services.Storage.RegisterFolderAsync([folder.Item], default);
            using var shell = new MainWindowViewModel(_ => Task.FromResult<string?>(source), services: services,
                pickWorkflowFiles: _ => Task.FromResult<IReadOnlyList<string>>([source]),
                pickOutputFolder: _ => Task.FromResult(destination), confirmProviderWrite: _ => Task.FromResult(true));
            if (health) {
                var model = shell.DocumentHealth;
                await model.ChooseInputCommand.ExecuteAsync(null);
                model.PrepareWorkflow(OfficeIMO.Workflows.OfficeWorkflowOperation.Optimize);
                model.OutputFolder = destination!;
                await model.RunCommand.ExecuteAsync(null);
                Assert.Equal(OfficeIMO.Workflows.OfficeWorkflowStatus.Completed, model.ResultStatus);
            } else {
                var model = shell.ConversionWorkbench;
                model.SelectedRoute = model.Routes.Single(route => route.Route.Id == "html-pdf");
                await model.AddFilesCommand.ExecuteAsync(null);
                model.OutputFolder = destination!;
                Assert.False(model.CanChooseConflictPolicy);
                await model.RunQueueCommand.ExecuteAsync(null);
                Assert.Equal(ConversionJobState.Completed, Assert.Single(model.Jobs).State);
            }
            var job = Assert.Single(services.Jobs.Entries);
            var file = Assert.Single(folder.Files).Value;
            Assert.Equal(file.Location.AbsoluteUri, job.OutputPath);
            Assert.True(job.HasOutput, job.Summary);
            var reopened = await services.Storage.ReadSnapshotAsync(job.OutputPath!, default);
            Assert.NotEmpty(OfficeIMO.Pdf.PdfDocument.Load(reopened.Bytes).Inspect().Pages);
            var window = new Window { Width = 1100, Height = 740, Content = health
                ? new DocumentHealthView { DataContext = shell } : new ConversionWorkbenchView { DataContext = shell } };
            try { window.Show(); window.UpdateLayout(); Capture(window, health ? "provider-folder-health.png" : "provider-folder-conversion.png"); }
            finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(960, 620, false, true, true)]
    [InlineData(1280, 820, true, false, true)]
    [InlineData(960, 620, false, false, false)]
    public async Task PageExportCreatesOpaqueChildrenAndShowsVerifiedPartialResults(int width, int height, bool dark, bool failSecond, bool consent) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "source.pdf");
            File.WriteAllBytes(source, StudioProviderDocumentTests.CreatePdf(2));
            var folder = new OutputFolder {
                FailSecondCreation = failSecond,
                BeforeCreate = () => Assert.NotEmpty(Directory.GetFiles(services.WorkflowRecovery.DirectoryPath, "record.json", SearchOption.AllDirectories))
            };
            string? destination = await services.Storage.RegisterFolderAsync([folder.Item], default);
            using var shell = new MainWindowViewModel(_ => Task.FromResult<string?>(source), services: services,
                pickOutputFolder: _ => Task.FromResult(destination), confirmProviderWrite: _ => Task.FromResult(consent));
            shell.OutputWorkbench.SelectedSection = OutputWorkbenchSection.ExportPages;
            var model = shell.OutputWorkbench.PageExport;
            model.UseDocument(source);
            model.MaximumDimension = 64;
            await model.ChooseOutputDirectoryCommand.ExecuteAsync(null);
            await model.ExportCommand.ExecuteAsync(null);
            if (!consent) {
                Assert.Equal(0, folder.Creations);
                Assert.Empty(services.Jobs.Entries);
                Assert.False(model.HasOutput);
                var dialog = new ProviderSaveDialog("Selected output folder", services.Localizer, workflowOutput: true, folderOutput: true);
                try { dialog.Show(); dialog.UpdateLayout(); Capture(dialog, "provider-folder-consent.png"); }
                finally { dialog.Close(); }
                return true;
            }
            Assert.Equal(2, folder.Creations);
            Assert.True(model.HasOutput, model.Summary);
            Assert.Equal(failSecond, model.HasRecovery);
            var job = Assert.Single(services.Jobs.Entries);
            Assert.True(job.HasOutput, job.Summary);
            Assert.Equal(failSecond, job.HasRecovery);
            var first = folder.Files.Values.First();
            Assert.True(OfficeImageReader.TryValidateContent(first.Bytes, first.Name, default, out _));
            Assert.Equal(1, first.Disposals);
            var reopened = await services.Storage.ReadSnapshotAsync(first.Location.AbsoluteUri, default);
            Assert.Equal(first.Bytes, reopened.Bytes);
            Assert.Equal(first.Name, services.Storage.Describe(first.Location.AbsoluteUri).Name);
            if (failSecond) {
                await services.WorkflowRecovery.VerifyAsync(job.Recovery!);
                Assert.True(OfficeImageReader.TryValidateContent(File.ReadAllBytes(job.Recovery!.FilePath), job.Recovery.Name, default, out _));
                Assert.Contains("1 of 2", model.Summary);
            } else Assert.Empty(services.WorkflowRecovery.GetRecoveries());
            var window = new Window { Width = width, Height = height, Content = new OutputIntakeWorkbenchView { DataContext = shell } };
            try {
                window.Show(); window.UpdateLayout();
                Capture(window, $"provider-folder-output-{width}-{(dark ? "dark" : "light")}.png");
                if (failSecond) {
                    window.Content = new StudioJobsView { DataContext = shell.Jobs };
                    window.UpdateLayout();
                    Capture(window, "provider-folder-output-partial-jobs.png");
                }
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
        internal bool FailSecondCreation;
        internal Action? BeforeCreate;

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
            BeforeCreate?.Invoke();
            Creations++;
            var file = new TestStorageFile("content://provider/assigned/" + Creations, [], name);
            Files[name] = file;
            if (FailSecondCreation && Creations == 2) throw new IOException("Provider creation failed after creating the file.");
            return Task.FromResult<IStorageFile?>(file.Item);
        }
    }
}
