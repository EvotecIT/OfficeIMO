using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class PdfBatchExportVisualTests {
    [Theory]
    [InlineData(900, 650, false)]
    [InlineData(1400, 850, false)]
    [InlineData(900, 650, true)]
    [InlineData(1400, 850, true)]
    public async Task FolderControlsExportMixedFormatsWithOptionalCheckpoints(int width, int height, bool checkpointed) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string input = Path.Combine(services.Paths.Root, "source"), output = Path.Combine(services.Paths.Root, "PDF"), state = Path.Combine(services.Paths.Root, "checkpoints");
            Directory.CreateDirectory(input);
            File.WriteAllText(Path.Combine(input, "one.txt"), "<h1>Studio literal text</h1>");
            File.WriteAllText(Path.Combine(input, "two.md"), "# Studio Markdown");
            File.WriteAllText(Path.Combine(input, "three.html"), "<h1>Studio HTML</h1>");
            File.WriteAllText(Path.Combine(input, "four.bin"), "Skipped input");
            var folders = new Queue<string>([input, output, state]);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickOutputFolder: _ => Task.FromResult<string?>(folders.Dequeue()));
            var view = new ConversionWorkbenchView { DataContext = model };
            var window = new Window { Width = width, Height = height, Content = view };
            try {
                window.Show();
                var archive = model.ConversionWorkbench.BatchExport;
                var panel = view.FindControl<Expander>("BatchExportPanel")!; panel.IsExpanded = true;
                foreach (string target in checkpointed ? new[] { "input", "output", "state" } : new[] { "input", "output" }) await archive.ChooseFolderCommand.ExecuteAsync(target);
                Assert.Equal(input, archive.InputDirectory); Assert.Equal(output, archive.OutputDirectory);
                Assert.Equal(checkpointed ? state : string.Empty, archive.CheckpointDirectory);
                window.UpdateLayout();
                var run = view.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, archive.RunCommand));
                panel.FindAncestorOfType<ScrollViewer>()!.ScrollToEnd(); window.UpdateLayout();
                Assert.True(run.IsEffectivelyVisible); Assert.True(run.IsEnabled);
                Point location = run.TranslatePoint(new Point(run.Bounds.Width / 2, run.Bounds.Height / 2), window)!.Value;
                Assert.InRange(location.X, 0, width); Assert.InRange(location.Y, 0, height);
                Capture(window, $"batch-ready-{width}-{checkpointed}.png");
                window.MouseDown(location, MouseButton.Left); window.MouseUp(location, MouseButton.Left);
                if (archive.RunCommand.ExecutionTask is { } execution) await execution;
                Assert.True(File.Exists(Path.Combine(output, "one.txt.pdf")), archive.Status);
                Assert.Contains("3 completed", archive.Status);
                Assert.Contains("1 skipped", archive.Status);
                string text = OfficeIMO.Pdf.PdfDocument.Load(Path.Combine(output, "one.txt.pdf")).Read().Text;
                Assert.Contains("<h1>Studio literal text</h1>", text);
                if (checkpointed) {
                    await archive.RunCommand.ExecuteAsync(null);
                    Assert.Contains("3 reused", archive.Status);
                } else Assert.False(Directory.Exists(state));
                Assert.Equal(archive.Status, model.ConversionWorkbench.Status);
                Assert.True(archive.CanEdit); Assert.False(archive.CancelCommand.CanExecute(null));
                window.UpdateLayout(); panel.FindAncestorOfType<ScrollViewer>()!.ScrollToEnd(); window.UpdateLayout();
                Capture(window, $"batch-finished-{width}-{checkpointed}.png");
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
