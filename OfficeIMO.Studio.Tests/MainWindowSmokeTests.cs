using Avalonia.Controls;
using Avalonia.Headless;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class MainWindowSmokeTests {
    [Fact]
    public async Task JobToastOffersOpenOnlyForBrowsableOutputs() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var window = new MainWindow();
            try {
                window.Show();
                var providerJob = window.ViewModel.Jobs.History.Start("Page image export", "source.pdf", "content://folder/selected", () => { });
                providerJob.CompleteBatch(OfficeIMO.Workflows.OfficeWorkflowStatus.Completed, "content://folder/selected",
                    "Images ready", [], hasVerifiedOutput: true, isDirectoryOutput: true);
                await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => window.UpdateLayout(), Avalonia.Threading.DispatcherPriority.Background);
                var open = window.FindControl<Avalonia.Controls.Button>("JobToastOpenButton")!;
                Assert.False(open.IsEffectivelyVisible);
                Capture(window, "provider-folder-toast.png");

                var localJob = window.ViewModel.Jobs.History.Start("Page image export", "source.pdf", Path.GetTempPath(), () => { });
                localJob.CompleteBatch(OfficeIMO.Workflows.OfficeWorkflowStatus.Completed, Path.GetTempPath(),
                    "Images ready", [], hasVerifiedOutput: true, isDirectoryOutput: true);
                await Avalonia.Threading.Dispatcher.UIThread.InvokeAsync(() => window.UpdateLayout(), Avalonia.Threading.DispatcherPriority.Background);
                Assert.True(open.IsEffectivelyVisible);
                Capture(window, "local-folder-toast.png");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task CreatesAndLaysOutResponsiveStudioSurfaces() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var window = new MainWindow();
            try {
                window.Show();
                window.Measure(new Avalonia.Size(1280, 820));
                window.Arrange(new Avalonia.Rect(0, 0, 1280, 820));

                Assert.NotNull(window.DataContext);
                Assert.IsType<MainWindowViewModel>(window.DataContext);
                Assert.False(window.ViewModel.HasDocument);
                Assert.True(window.ViewModel.IsEmpty);
                Assert.True(window.ViewModel.IsHomeMode);

                window.ViewModel.ShowPdfWorkspaceCommand.Execute(null);
                window.Width = 1050;
                window.Height = 560;
                window.Measure(new Avalonia.Size(1050, 560));
                window.Arrange(new Avalonia.Rect(0, 0, 1050, 560));
                window.ApplyResponsiveLayout(1050);
                Assert.True(window.ViewModel.IsPdfWorkspaceMode);
                Assert.True(window.IsCompactLayout);
                Assert.True(window.AreFitShortcutsVisible);

                window.ViewModel.ShowConversionWorkbenchCommand.Execute(null);
                window.Measure(new Avalonia.Size(1050, 560));
                window.Arrange(new Avalonia.Rect(0, 0, 1050, 560));
                Assert.True(window.ViewModel.IsConversionMode);
                Assert.True(window.IsConversionCompact);

                window.ViewModel.ShowPrintPreviewCommand.Execute(null);
                window.Measure(new Avalonia.Size(1050, 560));
                window.Arrange(new Avalonia.Rect(0, 0, 1050, 560));
                Assert.True(window.ViewModel.IsOutputMode);
                Assert.True(window.ViewModel.OutputWorkbench.IsPrintPreview);

                window.ViewModel.ShowPageExportCommand.Execute(null);
                window.Measure(new Avalonia.Size(1600, 1000));
                window.Arrange(new Avalonia.Rect(0, 0, 1600, 1000));
                Assert.True(window.ViewModel.OutputWorkbench.IsPageExport);

                window.ViewModel.ShowAssemblyCommand.Execute(null);
                window.Measure(new Avalonia.Size(1280, 820));
                window.Arrange(new Avalonia.Rect(0, 0, 1280, 820));
                Assert.True(window.ViewModel.OutputWorkbench.IsAssembly);

                window.ViewModel.ShowOcrCommand.Execute(null);
                window.Measure(new Avalonia.Size(1280, 820));
                window.Arrange(new Avalonia.Rect(0, 0, 1280, 820));
                Assert.True(window.ViewModel.IsOcrMode);

                window.ViewModel.ShowDocumentHealthCommand.Execute(null);
                window.Measure(new Avalonia.Size(1050, 560));
                window.Arrange(new Avalonia.Rect(0, 0, 1050, 560));
                Assert.True(window.IsDocumentHealthCompact);
                window.Width = 1600;
                window.Height = 1000;
                window.Measure(new Avalonia.Size(1600, 1000));
                window.Arrange(new Avalonia.Rect(0, 0, 1600, 1000));
                window.ApplyResponsiveLayout(1600);
                Assert.True(window.ViewModel.IsDocumentHealthMode);
                Assert.False(window.IsCompactLayout);
                Assert.False(window.IsDocumentHealthCompact);
                Assert.True(window.AreFitShortcutsVisible);
            } finally {
                window.Close();
            }

            return true;
        }, CancellationToken.None);
    }

    private static void Capture(MainWindow window, string name) {
        string? path = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(path)) return;
        Directory.CreateDirectory(path);
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        frame.Save(Path.Combine(path, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
