using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class ConversionQueueVisualTests {
    [Theory]
    [InlineData(840, 500, false)]
    [InlineData(1160, 700, true)]
    [InlineData(1600, 800, false)]
    public async Task PendingAndRetryActionsStayReachableAndPreserveCompletedOutputs(int width, int height, bool dark) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            string good = Path.Combine(services.Paths.Root, "Quarterly report.html");
            string missing = Path.Combine(services.Paths.Root, "Report requiring a restored source file.html");
            File.WriteAllText(good, "<html><body><p>Completed output must remain unchanged on retry.</p></body></html>");
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickWorkflowFiles: _ => Task.FromResult<IReadOnlyList<string>>([good, missing]), services: services);
            var queue = model.ConversionWorkbench;
            queue.SelectedRoute = queue.Routes.Single(route => route.Route.Id == "html-pdf");
            await queue.AddFilesCommand.ExecuteAsync(null);
            var view = new ConversionWorkbenchView { DataContext = model };
            var window = new Window { Width = width, Height = height, Content = view };
            try {
                window.Show();
                Layout(window, width, height);
                var run = ButtonFor(view, queue.RunQueueCommand);
                var retry = ButtonFor(view, queue.RetryFailedCommand);
                CheckBounds(window, run);
                CheckBounds(window, retry);
                Assert.True(run.IsEffectivelyEnabled);
                Assert.False(retry.IsEffectivelyEnabled);
                await queue.RunQueueCommand.ExecuteAsync(null);
                Layout(window, width, height);
                Assert.False(run.IsEffectivelyEnabled);
                Assert.True(retry.IsEffectivelyEnabled);
                Assert.Equal(ConversionJobState.Failed, queue.Jobs[1].State);
                foreach (ProgressBar progress in view.GetVisualDescendants().OfType<ProgressBar>()) {
                    if (progress.Parent is not Grid row || Grid.GetColumn(progress) != 3) continue;
                    double columnStart = row.ColumnDefinitions.Take(3).Sum(column => column.ActualWidth);
                    Assert.True(progress.Bounds.X >= columnStart - 0.5D, "Progress must stay inside its own column.");
                    Assert.True(progress.Bounds.Right <= row.Bounds.Width + 0.5D);
                }
                Capture(window, width, dark, "failed");
                File.WriteAllText(missing, "<html><body><p>The restored input can now be converted.</p></body></html>");
                retry.Command!.Execute(null);
                await queue.RetryFailedCommand.ExecutionTask!;
                Layout(window, width, height);
                Assert.All(queue.Jobs, job => Assert.Equal(ConversionJobState.Completed, job.State));
                Assert.False(retry.IsEffectivelyEnabled);
                Assert.Equal(2, Directory.GetFiles(services.Paths.Root, "*.pdf").Length);
                CheckBounds(window, run);
                CheckBounds(window, retry);
                Capture(window, width, dark, "retried");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static Button ButtonFor(Control view, System.Windows.Input.ICommand command) =>
        view.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, command));

    private static void CheckBounds(Window window, Button button) {
        Assert.True(button.IsEffectivelyVisible);
        Point point = button.TranslatePoint(default, window)!.Value;
        Assert.InRange(point.X, 0, window.Bounds.Width - button.Bounds.Width);
        Assert.InRange(point.Y, 0, window.Bounds.Height - button.Bounds.Height);
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
        frame.Save(Path.Combine(output, $"conversion-queue-{width}-{(dark ? "dark" : "light")}-{state}.png"),
            Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
