using System.Globalization;
using System.Text;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using PdfPageCanvas = OfficeIMO.Studio.Features.Reader.PdfPageCanvas;

namespace OfficeIMO.Studio.Tests;

public sealed class DirectEditingCoordinateTests {
    [Theory]
    [InlineData(0)] [InlineData(90)] [InlineData(270)]
    public async Task InlineEditorAndPreviewFollowScaledCroppedRotatedCanvas(int rotation) {
        using var files = new TextEditingReviewTests.Files();
        byte[] source = Source(rotation);
        File.WriteAllBytes(files.Source, source);
        var document = PdfDocument.Load(source);
        PdfTextMatch match = Assert.Single(document.Text.Find("Account"));
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            var window = new MainWindow(services) { Width = 1280, Height = 900 };
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(files.Source);
                var model = window.ViewModel;
                model.ShowEditModeCommand.Execute(null);
                model.ActualSizeCommand.Execute(null);
                window.UpdateLayout();
                await TextEditingReviewTests.WaitUntilAsync(() => model.Pages[0].Scene != null);
                window.UpdateLayout();
                var canvas = window.GetVisualDescendants().OfType<PdfPageCanvas>().First(c => c.Scene?.PageNumber == 1 && c.SelectionMode == PdfEditorSelectionMode.PageContent);
                double scaleX = canvas.Bounds.Width / canvas.Scene!.Drawing.Width;
                double scaleY = canvas.Bounds.Height / canvas.Scene.Drawing.Height;
                Point click = canvas.TranslatePoint(new Point((match.VisualBounds.Left + match.VisualBounds.Width / 2) * scaleX,
                    (match.VisualBounds.Top + match.VisualBounds.Height / 2) * scaleY), window)!.Value;
                window.MouseDown(click, MouseButton.Left); window.MouseUp(click, MouseButton.Left);
                await TextEditingReviewTests.WaitUntilAsync(() => model.TextEditDraft?.IsReady == true);
                window.UpdateLayout();
                var page = model.Pages[0];
                var selected = Assert.IsType<PdfEditorSelection>(model.SelectedObject);
                scaleX = canvas.Bounds.Width / canvas.Scene.Drawing.Width;
                scaleY = canvas.Bounds.Height / canvas.Scene.Drawing.Height;
                Assert.Equal(Math.Clamp(selected.Bounds.Left * scaleX, 0, Math.Max(0, canvas.Bounds.Width - page.InlineEditorWidth - 4)), page.InlineEditorLeft, 5);
                Assert.Equal(Math.Clamp(selected.Bounds.Top * scaleY, 0, Math.Max(0, canvas.Bounds.Height - 160)), page.InlineEditorTop, 5);
                var editor = Assert.Single(window.GetVisualDescendants().OfType<InlineTextEditorView>(), c => c.IsEffectivelyVisible);
                Point origin = canvas.TranslatePoint(default, window)!.Value;
                Point editorOrigin = editor.TranslatePoint(default, window)!.Value;
                Assert.InRange(Math.Abs(editorOrigin.X - origin.X - page.InlineEditorLeft), 0, 1.1);
                Assert.InRange(Math.Abs(editorOrigin.Y - origin.Y - page.InlineEditorTop), 0, 1.1);
                Capture(window, rotation);
                model.SelectedObjectText = "Record";
                await model.PreviewTextEditCommand.ExecuteAsync(null);
                Assert.True(model.HasTextPreview, model.ErrorMessage);
                AssertPreviewContainsMatch(model, match, document);
            } finally {
                foreach (var tab in window.TabHost.Tabs.ToArray()) tab.Document.CompletePreparedClose();
                window.Close(); window.TabHost.Dispose();
            }
            return true;
        }, default);
    }

    [Fact]
    public async Task BatchPreviewUsesPhysicalGeometryBeforeTheReaderSceneLoads() {
        using var files = new TextEditingReviewTests.Files();
        byte[] source = Source(0); File.WriteAllBytes(files.Source, source);
        var document = PdfDocument.Load(source);
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await model.OpenDocumentAsync(files.Source);
            model.ShowEditModeCommand.Execute(null);
            model.ReplaceAllFindText = "Account"; model.ReplaceAllReplacementText = "Record";
            await model.FindTextReplacementsCommand.ExecuteAsync(null);
            await model.PreviewTextEditCommand.ExecuteAsync(null);
            Assert.Null(model.Pages[0].Scene);
            Assert.True(model.HasTextPreview, model.ErrorMessage);
            AssertPreviewContainsMatch(model, document.Text.Find("Account")[0], document);
            return true;
        }, default);
    }

    private static void AssertPreviewContainsMatch(MainWindowViewModel model, PdfTextMatch match, PdfDocument document) {
        Rect region = Assert.IsType<Rect>(model.TextPreviewRegion);
        var drawing = document.Render.Drawing(1);
        Assert.InRange((match.VisualBounds.Left + match.VisualBounds.Width / 2) / drawing.Width, region.Left, region.Right);
        Assert.InRange((match.VisualBounds.Top + match.VisualBounds.Height / 2) / drawing.Height, region.Top, region.Bottom);
    }

    private static byte[] Source(int rotation) {
        const string content = "BT /F1 12 Tf 70 480 Td (Account) Tj ET";
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", $"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 500 600] /CropBox [20 40 480 580] /UserUnit 2 /Rotate {rotation} /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length.ToString(CultureInfo.InvariantCulture) + " >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>", "endobj",
            "trailer", "<< /Root 1 0 R /Size 6 >>", "%%EOF"
        }));
    }

    private static void Capture(Window window, int rotation) {
        using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return;
        Directory.CreateDirectory(output);
        frame.Save(Path.Combine(output, $"inline-userunit-rotation-{rotation}.png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
