using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;

namespace OfficeIMO.Studio.Tests;

public sealed class ConversionIntakeVisualTests {
    [Theory]
    [InlineData(960)]
    [InlineData(1400)]
    public async Task MismatchedPdfCanBeQueuedFromItsVisibleRouteChoiceAndConverted(int width) {
        using var files = new TextEditingReviewTests.Files();
        PdfDocument.Create(document => document.Page(page => page.Content(content => content.Text("Conversion evidence")))).Save(files.Source);
        byte[] original = File.ReadAllBytes(files.Source);
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null), services: services,
                pickWorkflowFiles: _ => Task.FromResult<IReadOnlyList<string>>([files.Source]));
            var view = new ConversionWorkbenchView { DataContext = model };
            var window = new Window { Content = view, Width = width, Height = 720 };
            try {
                window.Show();
                var conversion = model.ConversionWorkbench;
                Assert.Equal("docx-pdf", conversion.SelectedRoute.Route.Id);
                await conversion.AddFilesCommand.ExecuteAsync(null);
                Assert.True(conversion.HasUnmatchedInputs);
                Assert.Empty(conversion.Jobs);
                window.UpdateLayout();
                var choice = view.GetVisualDescendants().OfType<Button>().Single(button =>
                    button.IsEffectivelyVisible && button.CommandParameter is ConversionRouteChoice route && route.Route.Id == "pdf-html");
                Assert.Same(conversion.UseInputRouteCommand, choice.Command);
                Point point = choice.TranslatePoint(new Point(choice.Bounds.Width / 2, choice.Bounds.Height / 2), window)!.Value;
                Assert.InRange(point.X, 0, window.Bounds.Width);
                Assert.InRange(point.Y, 0, window.Bounds.Height);
                Capture(window, $"conversion-mismatch-{width}.png");
                window.MouseDown(point, MouseButton.Left);
                window.MouseUp(point, MouseButton.Left);
                var job = Assert.Single(conversion.Jobs);
                Assert.Equal("pdf-html", job.Route.Route.Id);
                Assert.False(conversion.HasUnmatchedInputs);
                await conversion.RunQueueCommand.ExecuteAsync(null);
                Assert.Equal(ConversionJobState.Completed, job.State);
                Assert.True(job.HasOutput, conversion.Status);
                string html = File.ReadAllText(job.OutputPath!);
                string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
                if (!string.IsNullOrWhiteSpace(output)) File.WriteAllText(Path.Combine(output, $"conversion-output-{width}.html"), html);
                int svgStart = html.IndexOf("<svg", StringComparison.Ordinal);
                Assert.True(svgStart >= 0, "Positioned HTML must contain its text layer.");
                int svgEnd = html.IndexOf("</svg>", svgStart, StringComparison.Ordinal);
                Assert.True(svgEnd > svgStart, "The text layer must be complete.");
                var textLayer = System.Xml.Linq.XElement.Parse(html[svgStart..(svgEnd + 6)]);
                string visibleText = string.Join(" ", textLayer.Descendants()
                    .Where(element => element.Name.LocalName == "text").Select(element => element.Value));
                Assert.Equal("Conversion evidence", visibleText);
                Assert.Equal(original, File.ReadAllBytes(files.Source));
                window.UpdateLayout();
                foreach (var command in new System.Windows.Input.ICommand[] { conversion.PreviewOutputCommand, conversion.OpenOutputCommand }) {
                    var action = Assert.Single(view.GetVisualDescendants().OfType<Button>(), button => ReferenceEquals(button.Command, command));
                    Point actionPoint = action.TranslatePoint(new Point(action.Bounds.Width / 2, action.Bounds.Height / 2), window)!.Value;
                    Assert.True(action.IsEffectivelyVisible);
                    Assert.InRange(actionPoint.X, 0, window.Bounds.Width);
                    Assert.InRange(actionPoint.Y, 0, window.Bounds.Height);
                    Assert.True(action.IsEnabled);
                }
                Capture(window, $"conversion-completed-{width}.png");
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static void Capture(Window window, string name) {
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        frame.Save(Path.Combine(output, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
