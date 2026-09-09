using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure.Preferences;
using PdfPageCanvas = OfficeIMO.Studio.Features.Reader.PdfPageCanvas;

namespace OfficeIMO.Studio.Tests;

public sealed class DirectEditingVisualTests {
    [Theory]
    [InlineData(960, 620, false)]
    [InlineData(1280, 900, true)]
    public async Task InlineEditorAndPreparedResultAreUsableAtBothSizes(int width, int height, bool dark) {
        using var files = new TextEditingReviewTests.Files();
        PdfDocument document = TextEditingReviewTests.CreateDocument();
        document.Save(files.Source);
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(current => current with { Theme = dark ? StudioThemePreference.Dark : StudioThemePreference.Light });
            var window = new MainWindow(services) { Width = width, Height = height };
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(files.Source);
                var model = window.ViewModel;
                model.ShowEditModeCommand.Execute(null);
                window.UpdateLayout();
                await TextEditingReviewTests.WaitUntilAsync(() => model.Pages[0].Scene is not null);
                window.UpdateLayout();
                var canvas = window.GetVisualDescendants().OfType<PdfPageCanvas>().First(control => control.Scene?.PageNumber == 1 && control.SelectionMode == PdfEditorSelectionMode.PageContent);
                PdfTextMatch match = document.Text.Find("Account")[0];
                Point position = PagePoint(canvas, (match.VisualBounds.TopLeft.X + match.VisualBounds.BottomRight.X) / 2,
                    (match.VisualBounds.TopLeft.Y + match.VisualBounds.BottomRight.Y) / 2, window);
                window.MouseDown(position, MouseButton.Left); window.MouseUp(position, MouseButton.Left);
                Capture(window, $"text-selection-{width}-{dark}.png");
                Assert.True(model.HasSelectedText, $"Click {position}; canvas {canvas.Bounds}; mode {canvas.SelectionMode}; error {model.ErrorMessage}");
                await TextEditingReviewTests.WaitUntilAsync(() => model.TextEditDraft?.IsReady == true);
                window.UpdateLayout();
                var editor = Assert.Single(window.GetVisualDescendants().OfType<InlineTextEditorView>(), control => control.IsEffectivelyVisible);
                var input = editor.FindControl<TextBox>("ReplacementText")!;
                input.Focus();
                window.KeyPress(Key.A, RawInputModifiers.Control, PhysicalKey.A, "a");
                window.KeyRelease(Key.A, RawInputModifiers.Control, PhysicalKey.A, "a");
                window.KeyTextInput("Reviewed account details");
                Assert.Equal("Reviewed account details", model.SelectedObjectText);
                var fitChoice = window.GetVisualDescendants().OfType<ComboBox>().Single(control => control.IsEffectivelyVisible && ReferenceEquals(control.ItemsSource, model.TextEditDraft!.Fits));
                fitChoice.SelectedItem = model.TextEditDraft!.Fits.Single(choice => choice.Value == PdfTextRegionWidthPolicy.PreserveFontSize);
                AssertContained(window, input);
                Capture(window, $"inline-text-{width}-{dark}.png");
                var preview = editor.GetVisualDescendants().OfType<Button>().Single(button => ReferenceEquals(button.Command, model.PreviewTextEditCommand));
                Click(window, preview);
                if (model.PreviewTextEditCommand.ExecutionTask is { } previewTask) await previewTask;
                Assert.True(model.HasTextPreview, model.ErrorMessage);
                window.UpdateLayout();
                var apply = window.GetVisualDescendants().OfType<Button>().Single(button => button.IsEffectivelyVisible && ReferenceEquals(button.Command, model.ApplyReviewedTextEditCommand));
                apply.BringIntoView(); window.UpdateLayout();
                AssertContained(window, apply);
                Capture(window, $"text-preview-{width}-{dark}.png");
                Click(window, apply);
                if (model.ApplyReviewedTextEditCommand.ExecutionTask is { } applyTask) await applyTask;
                Assert.True(model.IsDirty, model.ErrorMessage);
                Assert.Null(model.TextEditDraft);
            } finally { Close(window); }
            return true;
        }, default);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SelectedObjectClickDoesNotMoveButHandleDragResizesAndSupportsUndo(bool annotation) {
        using var files = new TextEditingReviewTests.Files();
        byte[] image = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        PdfDocument document = annotation
            ? PdfDocument.Load(TextEditingReviewTests.CreateDocument().Annotations.Add(new PdfAnnotationCreateOptions {
                Subtype = "Square", Rectangle = [2, 450, 122, 510], Contents = "Review area"
            }).Bytes)
            : TextEditingReviewTests.CreateDocument().Images.Add(new PdfPageRegion(1, 2, 450, 120, 60), image).Document;
        document.Save(files.Source);
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            var window = new MainWindow(services) { Width = 1280, Height = 900 };
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(files.Source);
                var model = window.ViewModel;
                if (annotation) model.ShowAnnotateModeCommand.Execute(null); else model.ShowEditModeCommand.Execute(null);
                window.UpdateLayout();
                await TextEditingReviewTests.WaitUntilAsync(() => model.Pages[0].Scene is not null);
                window.UpdateLayout();
                var mode = annotation ? PdfEditorSelectionMode.Annotations : PdfEditorSelectionMode.PageContent;
                var canvas = window.GetVisualDescendants().OfType<PdfPageCanvas>().First(control => control.Scene?.PageNumber == 1 && control.SelectionMode == mode);
                PdfPageInteractionRegion region = document.Render.Interactions(1).Regions.Single(item => item.Kind == (annotation ? PdfInteractionKind.Annotation : PdfInteractionKind.Image));
                Point center = PagePoint(canvas, (region.Quad.Left + region.Quad.Right) / 2, (region.Quad.Top + region.Quad.Bottom) / 2, window);
                window.MouseDown(center, MouseButton.Left); window.MouseUp(center, MouseButton.Left);
                Capture(window, annotation ? "annotation-selection.png" : "image-selection.png");
                Assert.True(annotation ? model.HasSelectedAnnotation : model.HasSelectedImage, $"Click {center}; canvas {canvas.Bounds}; mode {canvas.SelectionMode}; error {model.ErrorMessage}");
                var selected = model.SelectedObject;
                window.MouseDown(center, MouseButton.Left); window.MouseUp(center, MouseButton.Left);
                Assert.False(model.IsWorkspaceBusy || model.IsDirty);
                Assert.Same(selected, model.SelectedObject);
                window.MouseDown(center, MouseButton.Left);
                window.MouseMove(center + new Vector(1, 1), RawInputModifiers.LeftMouseButton);
                window.MouseUp(center + new Vector(1, 1), MouseButton.Left);
                Assert.False(model.IsWorkspaceBusy || model.IsDirty);
                Assert.Same(selected, model.SelectedObject);
                Point corner = PagePoint(canvas, region.Quad.Right, region.Quad.Bottom, window);
                window.MouseDown(corner, MouseButton.Left);
                Point target = corner + new Vector(45, 22.5);
                window.MouseMove(target, RawInputModifiers.LeftMouseButton);
                Capture(window, annotation ? "annotation-resize-handles.png" : "image-resize-handles.png");
                window.MouseUp(target, MouseButton.Left);
                await TextEditingReviewTests.WaitUntilAsync(() => !model.IsWorkspaceBusy && model.IsDirty);
                Assert.False(model.HasError, model.ErrorMessage);
                await model.UndoCommand.ExecuteAsync(null);
                Assert.False(model.IsDirty);
            } finally { Close(window); }
            return true;
        }, default);
    }

    private static Point PagePoint(PdfPageCanvas canvas, double x, double y, Window window) =>
        canvas.TranslatePoint(new Point(x * canvas.Bounds.Width / canvas.Scene!.Drawing.Width,
            y * canvas.Bounds.Height / canvas.Scene.Drawing.Height), window)!.Value;

    private static void Click(Window window, Button button) {
        button.BringIntoView(); window.UpdateLayout();
        Point point = button.TranslatePoint(new Point(button.Bounds.Width / 2, button.Bounds.Height / 2), window)!.Value;
        window.MouseDown(point, MouseButton.Left); window.MouseUp(point, MouseButton.Left);
    }

    private static void AssertContained(Window window, Control control) {
        Point point = control.TranslatePoint(default, window)!.Value;
        Assert.InRange(point.X, 0, window.Bounds.Width - control.Bounds.Width + 1);
        Assert.InRange(point.Y, 0, window.Bounds.Height - control.Bounds.Height + 1);
    }

    private static void Close(MainWindow window) {
        foreach (var tab in window.TabHost.Tabs.ToArray()) tab.Document.CompletePreparedClose();
        window.Close(); window.TabHost.Dispose();
    }

    private static void Capture(Window window, string name) {
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame);
        string? root = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(root)) return;
        Directory.CreateDirectory(root);
        frame.Save(Path.Combine(root, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
