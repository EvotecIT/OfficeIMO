using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using PdfPageCanvas = OfficeIMO.Studio.Features.Reader.PdfPageCanvas;
using OfficeIMO.Studio.Features.Shell;

namespace OfficeIMO.Studio.Tests;

public sealed class WatermarkSelectionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ClickingManagedContentOpensThatWatermarkInsteadOfInlineEditing(bool image) {
        using var files = new TextEditingReviewTests.Files();
        var settings = new PdfWatermarkOptions {
            Text = "CHOOSE THIS", X = 80, Y = 180, Width = 200, Height = 70, FontSize = 24, RotationDegrees = 0,
            ImageBytes = image ? Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=") : null,
            BehindContent = true
        };
        var document = PdfDocument.Create(builder => builder.Page(page => page.Size(600, 800)
            .Canvas(canvas => canvas.Text("FOREGROUND CONTENT", 80, 180, 200, 70, fontSize: 24))))
            .Stamp.Watermark(settings)
            .Stamp.Watermark(new PdfWatermarkOptions { Text = "OTHER MARK", X = 80, Y = 320, Width = 200, Height = 70, FontSize = 24, RotationDegrees = 0 });
        document.Save(files.Source);
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var window = new MainWindow(((App)Application.Current!).Services) { Width = 1280, Height = 900 };
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(files.Source);
                window.ViewModel.ShowEditModeCommand.Execute(null);
                window.UpdateLayout();
                await TextEditingReviewTests.WaitUntilAsync(() => window.ViewModel.Pages[0].Scene is not null);
                window.UpdateLayout();
                if (Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT") is { Length: > 0 } output) {
                    Directory.CreateDirectory(output);
                    using var frame = window.CaptureRenderedFrame();
                    Assert.NotNull(frame);
                    frame.Save(Path.Combine(output, $"watermark-inspector-{image}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
                }
                var canvas = window.GetVisualDescendants().OfType<PdfPageCanvas>().First(control =>
                    control.Scene?.PageNumber == 1 && control.SelectionMode == PdfEditorSelectionMode.PageContent);
                var region = canvas.Scene!.Interactions!.Regions.First(item => item.WatermarkId == settings.Id);
                var foreground = canvas.Scene.Interactions.Regions.First(item =>
                    item.WatermarkId is null && item.Kind == PdfInteractionKind.Text &&
                    item.Quad.Left < region.Quad.Right && item.Quad.Right > region.Quad.Left &&
                    item.Quad.Top < region.Quad.Bottom && item.Quad.Bottom > region.Quad.Top);
                double hitLeft = Math.Max(region.Quad.Left, foreground.Quad.Left);
                double hitRight = Math.Min(region.Quad.Right, foreground.Quad.Right);
                double hitTop = Math.Max(region.Quad.Top, foreground.Quad.Top);
                double hitBottom = Math.Min(region.Quad.Bottom, foreground.Quad.Bottom);
                var point = canvas.TranslatePoint(new Point(
                    (hitLeft + hitRight) / 2 * canvas.Bounds.Width / canvas.Scene.Drawing.Width,
                    (hitTop + hitBottom) / 2 * canvas.Bounds.Height / canvas.Scene.Drawing.Height), window)!.Value;
                window.MouseDown(point, MouseButton.Left);
                window.MouseUp(point, MouseButton.Left);
                await TextEditingReviewTests.WaitUntilAsync(() => window.OwnedWindows.OfType<WatermarkDialog>().Any());
                var dialog = Assert.Single(window.OwnedWindows.OfType<WatermarkDialog>());
                var model = Assert.IsType<WatermarkPreviewViewModel>(dialog.DataContext);
                await TextEditingReviewTests.WaitUntilAsync(() => model.CanApply);
                Assert.Equal(settings.Id, model.SelectedWatermark.Options!.Id);
                Assert.Equal(image, model.UseImage);
                Assert.Null(window.ViewModel.TextEditDraft);
                Assert.Null(window.ViewModel.SelectedObject);
                Assert.False(window.ViewModel.IsDirty);
            } finally {
                foreach (var dialog in window.OwnedWindows.ToArray()) dialog.Close(false);
                foreach (var tab in window.TabHost.Tabs.ToArray()) tab.Document.CompletePreparedClose();
                window.Close(); window.TabHost.Dispose();
            }
            return true;
        }, CancellationToken.None);
    }
}
