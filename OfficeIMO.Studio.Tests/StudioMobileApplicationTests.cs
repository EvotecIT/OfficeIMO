using System.Reflection;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Mobile;
using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Settings;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Tests;

public sealed partial class StudioMobileApplicationTests {
    [Theory]
    [InlineData(1366, 1024)]
    [InlineData(390, 844)]
    public async Task ApplicationNavigationRunsConversionAndKeepsTheDocument(int width, int height) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-mobile-app-" + Guid.NewGuid().ToString("N"));
        try {
            using var session = TestAppBuilder.StartSession();
            await session.Dispatch(async () => {
                var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "data")),
                    new StudioLocalDocumentRoot(Path.Combine(root, "documents")));
                var input = new TestStorageFile("content://mobile/convert.pdf",
                    PdfDocument.Create(pdf => pdf.Page(page => page.Content(content => content.Text("Shared conversion on iPad")))).ToBytes(), "convert.pdf");
                var provider = DispatchProxy.Create<IStorageProvider, TestStorageFile.StorageProxy>();
                ((TestStorageFile.StorageProxy)(object)provider).Call = (method, _) => method switch {
                    "OpenFilePickerAsync" => Task.FromResult<IReadOnlyList<IStorageFile>>([input.Item]),
                    _ => throw new NotSupportedException(method)
                };
                var view = new MobileWorkspaceView();
                var host = new MobileDocumentHost(services, view, _ => Task.CompletedTask, () => provider);
                using var controller = new MobileDocumentController(services, _ => Task.FromResult<IStorageFile?>(null), _ => Task.CompletedTask, host);
                view.Connect(controller);
                var window = new Window { Content = view, Width = width, Height = height };
                try {
                    window.Show(); Layout(window);
                    await controller.OpenSampleAsync();
                    var document = controller.Document;
                    string original = document.DocumentPath!;
                    await document.Commands["Home"].ExecuteAsync(); Layout(window); Capture(window, $"application-home-{width}");
                    await document.Commands["Tools"].ExecuteAsync(); Layout(window); Capture(window, $"application-tools-{width}");
                    await document.Commands["Convert"].ExecuteAsync(); Layout(window);
                    var conversion = document.ConversionWorkbench;
                    conversion.SelectedRoute = conversion.Routes.Single(route => route.Route.Id == "pdf-html");
                    conversion.OutputFolder = Path.Combine(root, "converted");
                    Directory.CreateDirectory(conversion.OutputFolder);
                    var conversionView = Assert.Single(view.GetVisualDescendants().OfType<ConversionWorkbenchView>());
                    if (width < 760) conversionView.FindControl<TabControl>("CompactTabs")!.SelectedItem = conversionView.FindControl<TabItem>("QueueTab");
                    Layout(window);
                    Click(window, view, conversion.AddFilesCommand);
                    await conversion.AddFilesCommand.ExecutionTask!;
                    Assert.Single(conversion.Jobs);
                    Layout(window); Capture(window, $"application-convert-{width}");
                    Click(window, view, conversion.RunQueueCommand);
                    await conversion.RunQueueCommand.ExecutionTask!;
                    var job = Assert.Single(conversion.Jobs);
                    Assert.Equal(ConversionJobState.Completed, job.State);
                    Assert.Contains("Shared conversion on iPad", File.ReadAllText(job.OutputPath!));
                    Assert.Equal(input.Reads, input.ClosedReads);
                    var finished = Assert.Single(services.Jobs.Entries);
                    Assert.False(document.Jobs.RevealOutputCommand.CanExecute(finished));
                    Assert.False(document.OutputActions.RevealCommand.CanExecute(job.OutputPath));
                    await document.Commands["Jobs"].ExecuteAsync(); Layout(window);
                    Assert.Same(document.Jobs, Assert.Single(view.GetVisualDescendants().OfType<StudioJobsView>()).DataContext);
                    Capture(window, $"application-jobs-{width}");
                    await document.Commands["Settings"].ExecuteAsync(); Layout(window);
                    Assert.Same(document.Settings, Assert.Single(view.GetVisualDescendants().OfType<SettingsView>()).DataContext);
                    Capture(window, $"application-settings-{width}");
                    await document.Commands["Assemble"].ExecuteAsync(); Layout(window);
                    var output = Assert.Single(view.GetVisualDescendants().OfType<OutputIntakeWorkbenchView>());
                    Capture(window, $"application-assembly-{width}");
                    if (width < 760) {
                        var tasks = output.GetVisualDescendants().OfType<TabControl>().Single(tabs => tabs.IsEffectivelyVisible);
                        tasks.SelectedIndex = 1;
                        Layout(window); Capture(window, $"application-assembly-options-{width}");
                    }
                    document.OutputWorkbench.ShowPageExportCommand.Execute(null); Layout(window);
                    Capture(window, $"application-page-export-{width}");
                    document.OutputWorkbench.ShowPrintPreviewCommand.Execute(null); Layout(window);
                    Capture(window, $"application-print-setup-{width}");
                    document.ShowPdfWorkspaceCommand.Execute(null); Layout(window);
                    Assert.Equal(original, controller.Document.DocumentPath);
                    Assert.Single(controller.Tabs.Tabs);
                    Assert.True(view.FindControl<Border>("ReaderSurface")!.IsEffectivelyVisible);
                } finally { window.Close(); services.Storage.Dispose(); }
                return true;
            }, CancellationToken.None);
        } finally { if (Directory.Exists(root)) Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(1366, 1024)]
    [InlineData(390, 844)]
    [InlineData(844, 390)]
    public async Task SharedReviewSheetAppliesAndCancelsWithoutAWindow(int width, int height) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            using var document = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services);
            var view = new MobileWorkspaceView { DataContext = document };
            var window = new Window { Content = view, Width = width, Height = height };
            try {
                window.Show(); Layout(window);
                var content = new PageMoveDialogContent(new PageMovePreviewViewModel(3, [1], services.Localizer));
                Task<bool> decision = view.ShowDialogAsync<bool>(content);
                Layout(window);
                Assert.False(view.FindControl<SplitView>("ApplicationNavigation")!.IsEnabled);
                Assert.Empty(window.OwnedWindows);
                Capture(window, $"application-review-{width}");
                var apply = content.GetVisualDescendants().OfType<Button>().Single(button => Equals(button.Content, services.Localizer.Get("Organizer.ApplyMove")));
                apply.BringIntoView(); Layout(window);
                Click(window, apply);
                Assert.True(await decision);
                Assert.True(view.FindControl<SplitView>("ApplicationNavigation")!.IsEnabled);
                decision = view.ShowDialogAsync<bool>(new PageMoveDialogContent(new PageMovePreviewViewModel(3, [1], services.Localizer)));
                Layout(window);
                window.Close();
                Assert.False(await decision);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static void Layout(Window window) {
        window.UpdateLayout();
        Dispatcher.UIThread.RunJobs();
    }

    private static void Click(Window window, Control scope, System.Windows.Input.ICommand command, bool requireHitTarget = false) =>
        Click(window, scope.GetVisualDescendants().OfType<Button>().First(button => ReferenceEquals(button.Command, command) && button.IsEffectivelyVisible), requireHitTarget);

    private static void Click(Window window, Button button, bool requireHitTarget = false) {
        Assert.True(button.IsEffectivelyEnabled);
        Point point = button.TranslatePoint(new Point(button.Bounds.Width / 2, button.Bounds.Height / 2), window)!.Value;
        Assert.InRange(point.X, 0, window.Bounds.Width);
        Assert.InRange(point.Y, 0, window.Bounds.Height);
        if (requireHitTarget) {
            var target = window.InputHitTest(point) as Visual;
            Assert.True(ReferenceEquals(target, button) || target?.GetVisualAncestors().Contains(button) == true,
                $"Button {Avalonia.Automation.AutomationProperties.GetName(button)} at {point} is obscured by {target}.");
        }
        window.MouseDown(point, MouseButton.Left);
        window.MouseUp(point, MouseButton.Left);
    }

    private static void Capture(Window window, string name) {
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        using var frame = window.CaptureRenderedFrame();
        frame!.Save(Path.Combine(output, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
