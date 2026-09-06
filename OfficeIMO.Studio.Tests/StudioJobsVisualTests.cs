using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Threading;
using Avalonia.VisualTree;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioJobsVisualTests {
    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 820, true)]
    [InlineData(1600, 900, false)]
    public async Task SharedJobsNavigationShowsOutcomesAndReopensARealOutput(int width, int height, bool dark) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            string source = Path.Combine(services.Paths.Root, "Job source.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(400, 500)
                .Content(content => content.Text("A real workflow output opened from Jobs.")))).Save(source);
            var window = new MainWindow(services) { Width = width, Height = height };
            StudioJobRecord? waiting = null;
            try {
                window.Show();
                var assembly = window.ViewModel.OutputWorkbench.Assembly;
                assembly.UseDocument(source);
                assembly.OutputPath = Path.Combine(services.Paths.Root, "Assembled report.pdf");
                await assembly.RunCommand.ExecuteAsync(null);
                StudioJobRecord assemblyEntry = Assert.Single(services.Jobs.Entries);
                Assert.True(assemblyEntry.HasOutput);
                var export = window.ViewModel.OutputWorkbench.PageExport;
                export.UseDocument(source);
                export.SelectedFormat = export.Formats.Single(format => format.Value == OfficeImageExportFormat.Svg);
                await export.ExportCommand.ExecuteAsync(null);
                Assert.Equal(2, services.Jobs.Entries.Count);
                Assert.True(services.Jobs.Entries[0].HasOutput);
                var conversion = window.ViewModel.ConversionWorkbench;
                var route = conversion.Routes.Single(candidate => candidate.Route.Id == "html-pdf");
                conversion.Jobs.Add(new ConversionJobViewModel(Path.Combine(services.Paths.Root, "Unavailable input.html"), route));
                await conversion.RunQueueCommand.ExecuteAsync(null);
                Assert.Equal("Failed", services.Jobs.Entries[0].Status);
                waiting = services.Jobs.Start("Waiting workflow", source, "Requested output.pdf",
                    () => waiting!.Complete(OfficeWorkflowStatus.Cancelled, null, "Cancelled before execution"));

                await window.ViewModel.Commands["Jobs"].ExecuteAsync();
                Assert.True(window.ViewModel.IsJobsMode);
                Assert.False(window.ViewModel.IsConversionMode);
                Layout(window, width, height);
                var jobsView = window.GetVisualDescendants().OfType<StudioJobsView>().Single();
                Assert.True(jobsView.IsEffectivelyVisible);
                var jobs = window.ViewModel.Jobs;
                Button clear = jobsView.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, jobs.ClearFinishedCommand));
                CheckBounds(window, clear);
                Button cancel = jobsView.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, waiting.CancelCommand));
                CheckBounds(window, cancel);
                Capture(window, width, dark, "outcomes");
                cancel.Command!.Execute(null);
                Assert.False(waiting.IsActive);

                ListBox list = jobsView.GetVisualDescendants().OfType<ListBox>().Single();
                list.ScrollIntoView(assemblyEntry);
                await Dispatcher.UIThread.InvokeAsync(() => { }, DispatcherPriority.Background);
                Layout(window, width, height);
                Button open = jobsView.GetVisualDescendants().OfType<Button>()
                    .Single(button => ReferenceEquals(button.CommandParameter, assemblyEntry) && ReferenceEquals(button.Command, jobs.OpenOutputCommand));
                Assert.True(open.IsEffectivelyEnabled);
                CheckBounds(window, open);
                Capture(window, width, dark, "output-action");
                open.Command!.Execute(assemblyEntry);
                await jobs.OpenOutputCommand.ExecutionTask!;
                Assert.Null(jobs.ActionError);
                Assert.Single(window.TabHost.Tabs);
                Assert.Equal(assemblyEntry.OutputPath, window.ViewModel.DocumentPath);
                Assert.Single(window.ViewModel.Pages);
                Assert.Same(services.Jobs, window.ViewModel.Jobs.History);
                await window.ViewModel.Commands["Jobs"].ExecuteAsync();
                Assert.Equal(4, window.ViewModel.Jobs.History.Entries.Count);
            } finally {
                if (waiting?.IsActive == true) waiting.Complete(OfficeWorkflowStatus.Cancelled, null, "Test ended");
                window.Close();
            }
            return true;
        }, CancellationToken.None);
    }

    private static void CheckBounds(Window window, Control control) {
        Assert.True(control.IsEffectivelyVisible);
        Point point = control.TranslatePoint(default, window)!.Value;
        Assert.InRange(point.X, 0, window.Bounds.Width - control.Bounds.Width);
        Assert.InRange(point.Y, 0, window.Bounds.Height - control.Bounds.Height);
    }
    private static void Layout(Window window, int width, int height) {
        window.Measure(new Size(width, height));
        window.Arrange(new Rect(0, 0, width, height));
        window.UpdateLayout();
    }
    private static void Capture(Window window, int width, bool dark, string state) {
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        frame.Save(Path.Combine(output, $"jobs-{width}-{(dark ? "dark" : "light")}-{state}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
