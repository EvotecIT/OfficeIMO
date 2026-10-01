using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfArchiveVisualTests {
    [Theory]
    [InlineData(900, 650)]
    [InlineData(1400, 850)]
    public async Task FolderControlsRunAndResumeTheSharedArchive(int width, int height) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string input = Path.Combine(services.Paths.Root, "source"), output = Path.Combine(services.Paths.Root, "PDF"), state = Path.Combine(services.Paths.Root, "checkpoints");
            Directory.CreateDirectory(input); File.WriteAllText(Path.Combine(input, "one.txt"), "  Studio archive evidence");
            var folders = new Queue<string>([input, output, state]);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickOutputFolder: _ => Task.FromResult<string?>(folders.Dequeue()));
            var view = new ConversionWorkbenchView { DataContext = model };
            var window = new Window { Width = width, Height = height, Content = view };
            try {
                window.Show();
                var archive = model.ConversionWorkbench.Archive;
                var panel = view.FindControl<Expander>("ArchivePanel")!; panel.IsExpanded = true;
                foreach (string target in new[] { "input", "output", "state" }) await archive.ChooseFolderCommand.ExecuteAsync(target);
                Assert.Equal(input, archive.InputDirectory); Assert.Equal(output, archive.OutputDirectory); Assert.Equal(state, archive.CheckpointDirectory);
                window.UpdateLayout();
                var run = view.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, archive.RunCommand));
                panel.FindAncestorOfType<ScrollViewer>()!.ScrollToEnd(); window.UpdateLayout();
                Assert.True(run.IsEffectivelyVisible); Assert.True(run.IsEnabled);
                Point location = run.TranslatePoint(new Point(run.Bounds.Width / 2, run.Bounds.Height / 2), window)!.Value;
                Assert.InRange(location.X, 0, width); Assert.InRange(location.Y, 0, height);
                Capture(window, $"archive-ready-{width}.png");
                window.MouseDown(location, MouseButton.Left); window.MouseUp(location, MouseButton.Left);
                if (archive.RunCommand.ExecutionTask is { } execution) await execution;
                Assert.True(File.Exists(Path.Combine(output, "one.txt.pdf")), archive.Status);
                Assert.Contains("1 completed", archive.Status);
                await archive.RunCommand.ExecuteAsync(null);
                Assert.Contains("1 reused", archive.Status);
                Assert.Equal(archive.Status, model.ConversionWorkbench.Status);
                Assert.True(archive.CanEdit); Assert.False(archive.CancelCommand.CanExecute(null));
                window.UpdateLayout(); panel.FindAncestorOfType<ScrollViewer>()!.ScrollToEnd(); window.UpdateLayout();
                Capture(window, $"archive-resumed-{width}.png");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }
    private static void Capture(Window window, string name) {
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
        frame.Save(Path.Combine(output, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
