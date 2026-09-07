using System.Text;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioProviderWorkflowTests {
    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 820, true)]
    public async Task HealthInspectsProviderWithoutOutputAndRequiresFolderForOptimization(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            byte[] original = StudioProviderDocumentTests.CreatePdf(2);
            var provider = new TestStorageFile("content://documents/health", original, "Provider document selected for inspection and compression.pdf");
            string location = await services.Storage.RegisterAsync(provider.Item, default);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(location), services: services);
            var health = model.DocumentHealth;
            await health.ChooseInputCommand.ExecuteAsync(null);
            Assert.Equal(provider.Name, health.InputFileName);
            Assert.Equal("Storage provider", health.InputDirectory);
            Assert.True(health.CanRun);
            await health.RunCommand.ExecuteAsync(null);
            Assert.Equal(OfficeWorkflowStatus.Completed, health.ResultStatus);
            Assert.True(health.HasHealthReport);
            Assert.False(health.HasOutput);
            health.PrepareWorkflow(OfficeWorkflowOperation.Optimize);
            health.ContinueCommand.Execute(null);
            Assert.True(health.IsOptionsStep);
            Assert.False(health.CanContinue);
            Assert.False(health.CanRun);
            var window = new Window { Width = width, Height = height, Content = new DocumentHealthView { DataContext = model } };
            try {
                window.Show();
                Capture(window, $"provider-health-{width}-{(dark ? "dark" : "light")}-choose-output.png");
                health.OutputFolder = services.Paths.Root;
                Assert.True(health.CanContinue);
                Assert.True(health.CanRun);
                health.ContinueCommand.Execute(null);
                await health.RunCommand.ExecuteAsync(null);
                Assert.Equal(OfficeWorkflowStatus.Completed, health.ResultStatus);
                Assert.Equal(2, PdfDocument.Load(health.OutputPath!).Inspect().PageCount);
                Assert.Equal(provider.Name.Replace(".pdf", ".optimized.pdf"), Path.GetFileName(health.OutputPath));
                Assert.Equal(original, provider.Bytes);
                Assert.Equal(0, provider.Writes);
                Assert.Equal(provider.Reads, provider.ClosedReads);
                Capture(window, $"provider-health-{width}-{(dark ? "dark" : "light")}-completed.png");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HealthComparisonUsesBothProviderStreamsAndHonorsRevokedAccess(bool denied) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            var first = new TestStorageFile("content://documents/first", StudioProviderDocumentTests.CreatePdf(), "First.pdf");
            var second = new TestStorageFile("content://documents/second", StudioProviderDocumentTests.CreatePdf(2), "Second.pdf") { DenyRead = denied };
            var locations = await services.Storage.RegisterManyAsync([first.Item, second.Item], default);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            var health = model.DocumentHealth;
            health.InputPath = locations[0];
            health.ComparisonPath = locations[1];
            health.OutputFolder = services.Paths.Root;
            health.PrepareWorkflow(OfficeWorkflowOperation.Compare);
            await health.RunCommand.ExecuteAsync(null);
            Assert.Equal(denied ? OfficeWorkflowStatus.Failed : OfficeWorkflowStatus.Completed, health.ResultStatus);
            Assert.Equal(!denied, health.HasOutput);
            if (!denied) Assert.True(File.Exists(health.OutputPath));
            Assert.Equal(0, first.Writes + second.Writes);
            Assert.Equal(first.Reads, first.ClosedReads);
            Assert.Equal(denied ? 0 : second.Reads, second.ClosedReads);
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 820, true)]
    public async Task ConversionKeepsProviderNamesAndRetriesOnlyFailedInput(int width, int height, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            byte[] html = Encoding.UTF8.GetBytes("<html><body><p>Provider input conversion result</p></body></html>");
            var first = new TestStorageFile("content://documents/opaque-first", html, "Quarterly review from the external document provider.html");
            var second = new TestStorageFile("content://documents/opaque-second", html, "Restored permission.html") { DenyRead = true };
            var locations = await services.Storage.RegisterManyAsync([first.Item, second.Item], default);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickWorkflowFiles: _ => Task.FromResult(locations), services: services);
            var queue = model.ConversionWorkbench;
            queue.SelectedRoute = queue.Routes.Single(route => route.Route.Id == "html-pdf");
            await queue.AddFilesCommand.ExecuteAsync(null);
            await queue.RunQueueCommand.ExecuteAsync(null);
            Assert.All(queue.Jobs, job => Assert.Equal(ConversionJobState.Queued, job.State));
            Assert.Equal(0, first.Reads + second.Reads);
            queue.OutputFolder = services.Paths.Root;
            var window = new Window { Width = width, Height = height, Content = new ConversionWorkbenchView { DataContext = model } };
            try {
                window.Show();
                await queue.RunQueueCommand.ExecuteAsync(null);
                Assert.Equal(ConversionJobState.Completed, queue.Jobs[0].State);
                Assert.Equal(ConversionJobState.Failed, queue.Jobs[1].State);
                Assert.Equal(first.Name, queue.Jobs[0].FileName);
                Assert.Equal(locations[0], queue.Jobs[0].InputPath);
                byte[] completed = File.ReadAllBytes(queue.Jobs[0].OutputPath!);
                int completedReads = first.Reads;
                Capture(window, $"provider-conversion-{width}-{(dark ? "dark" : "light")}-failed.png");
                second.DenyRead = false;
                await queue.RetryFailedCommand.ExecuteAsync(null);
                Assert.All(queue.Jobs, job => Assert.Equal(ConversionJobState.Completed, job.State));
                Assert.Equal(completedReads, first.Reads);
                Assert.Equal(completed, File.ReadAllBytes(queue.Jobs[0].OutputPath!));
                Assert.Equal(2, Directory.GetFiles(services.Paths.Root, "*.pdf").Length);
                Assert.All(queue.Jobs, job => Assert.NotEmpty(PdfDocument.Load(job.OutputPath!).Inspect().Pages));
                Capture(window, $"provider-conversion-{width}-{(dark ? "dark" : "light")}-completed.png");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task AssemblyAcceptsOpaqueProviderInputsAndPublishesAReopenablePdf() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            var first = new TestStorageFile("content://documents/first", StudioProviderDocumentTests.CreatePdf(2), "First selected document.pdf");
            var second = new TestStorageFile("content://documents/second", StudioProviderDocumentTests.CreatePdf(3), "Second selected document.pdf");
            var locations = await services.Storage.RegisterManyAsync([first.Item, second.Item], default);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickWorkflowFiles: _ => Task.FromResult(locations), services: services);
            var assembly = model.OutputWorkbench.Assembly;
            await assembly.AddFilesCommand.ExecuteAsync(null);
            Assert.Equal(first.Name, assembly.Sources[0].Name);
            Assert.Equal(second.Name, assembly.Sources[1].Name);
            Assert.False(assembly.RunCommand.CanExecute(null));
            assembly.OutputPath = Path.Combine(services.Paths.Root, "assembled.pdf");
            await assembly.RunCommand.ExecuteAsync(null);
            Assert.True(assembly.HasOutput, assembly.Status);
            Assert.Equal(5, PdfDocument.Load(assembly.PublishedPath!).Inspect().PageCount);
            Assert.Equal(first.Reads, first.ClosedReads);
            Assert.Equal(second.Reads, second.ClosedReads);
            Assert.Equal(0, first.Writes + second.Writes);
            return true;
        }, CancellationToken.None);
    }

    private static void Capture(Window window, string name) {
        window.UpdateLayout();
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        frame.Save(Path.Combine(output, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
