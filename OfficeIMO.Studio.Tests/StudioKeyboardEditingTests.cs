using Avalonia.Input;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Tests;

/// <summary>Everything the pointer does to a page object can also be done from the keyboard.</summary>
public sealed class StudioKeyboardEditingTests {
    [Theory]
    [InlineData(PdfEditorSelectionKind.Image, PdfEditorSelectionMode.PageContent)]
    [InlineData(PdfEditorSelectionKind.Annotation, PdfEditorSelectionMode.Annotations)]
    public async Task Arrow_keys_move_and_resize_the_selected_object(PdfEditorSelectionKind kind, PdfEditorSelectionMode mode) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var selection = new PdfEditorSelection(kind, 1, new PdfEditorVisualBounds(100, 100, 200, 150), ObjectNumber: 7);
            using var canvas = new PdfPageCanvas { Scene = TestPdfPageScenes.Create(), SelectionMode = mode, SelectedObject = selection };
            var transforms = new List<PdfEditorVisualBounds>();
            canvas.ObjectTransformCompleted += gesture => {
                Assert.True(gesture.KeepSelection);
                transforms.Add(gesture.Target);
            };
            KeyModifiers primary = Features.Shell.StudioShortcuts.PrimaryModifier;

            Press(canvas, Key.Right, primary);
            Press(canvas, Key.Down, primary | KeyModifiers.Shift);
            Press(canvas, Key.Right, primary | KeyModifiers.Alt);
            Press(canvas, Key.Up, primary | KeyModifiers.Alt | KeyModifiers.Shift);
            Press(canvas, Key.Right, KeyModifiers.None);

            bool image = kind == PdfEditorSelectionKind.Image;
            Assert.Equal(new[] {
                new PdfEditorVisualBounds(101, 100, 201, 150), // moved one point
                new PdfEditorVisualBounds(100, 110, 200, 160), // moved ten points with Shift
                // Wider by one point and shorter by ten; an image keeps its proportions.
                new PdfEditorVisualBounds(100, 100, 201, image ? 150.5 : 150),
                new PdfEditorVisualBounds(100, 100, image ? 180 : 200, 140)
            }, transforms);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public void Keyboard_resizing_keeps_an_image_in_proportion() {
        // The engine resizes images only with one scale factor, as the pointer handles do.
        var bounds = new PdfEditorVisualBounds(100, 100, 200, 150);
        Assert.Equal(new PdfEditorVisualBounds(100, 100, 210, 155), PdfPageCanvas.KeyboardResize(bounds, 10, 0, keepAspect: true));
        Assert.Equal(new PdfEditorVisualBounds(100, 100, 180, 140), PdfPageCanvas.KeyboardResize(bounds, 0, -10, keepAspect: true));
        Assert.Equal(new PdfEditorVisualBounds(100, 100, 210, 150), PdfPageCanvas.KeyboardResize(bounds, 10, 0, keepAspect: false));
    }

    [Fact]
    public async Task Enter_places_the_active_tool_in_the_middle_of_the_page() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            using var canvas = new PdfPageCanvas { Scene = TestPdfPageScenes.Create(), EditorTool = PdfEditorTool.Note };
            PdfEditorGesture? placed = null;
            canvas.EditorGestureCompleted += gesture => placed = gesture;

            Press(canvas, Key.Enter, KeyModifiers.None);

            Assert.NotNull(placed);
            double width = canvas.Scene!.Drawing.Width, height = canvas.Scene.Drawing.Height;
            Assert.InRange((placed.Left + placed.Right) / 2D, width / 2D - 1D, width / 2D + 1D);
            Assert.InRange((placed.Top + placed.Bottom) / 2D, height / 2D - 1D, height / 2D + 1D);

            // With the select tool Enter keeps its meaning of activating the focused link or field.
            placed = null;
            canvas.EditorTool = PdfEditorTool.Select;
            Press(canvas, Key.Enter, KeyModifiers.None);
            Assert.Null(placed);
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task Cancel_in_the_inspector_shows_the_saved_annotation_values_again() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-inspector-cancel-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "annotated.pdf");
            byte[] annotated = Pdf.PdfDocument.Load(new System.IO.MemoryStream(Pdf.PdfDocument.Create(compose => compose.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(600D), global::OfficeIMO.OfficeLength.Points(800D)))).ToBytes()))
                .Annotations.Add(new Pdf.PdfAnnotationCreateOptions { Subtype = "Text", Rectangle = [40D, 50D, 60D, 70D], Contents = "Original", Title = "Reviewer" }).Bytes;
            await File.WriteAllBytesAsync(path, annotated);
            using var document = new Features.Shell.MainWindowViewModel(_ => Task.FromResult<string?>(null));
            await document.OpenDocumentAsync(path);
            int objectNumber = Assert.Single(Pdf.PdfDocument.Load(new System.IO.MemoryStream(annotated)).Inspect().GetAnnotationsBySubtype("Text")).ObjectNumber!.Value;
            document.Pages[0].SelectObject(new PdfEditorSelection(PdfEditorSelectionKind.Annotation, 1,
                new PdfEditorVisualBounds(40D, 730D, 60D, 750D), ObjectNumber: objectNumber, Subtype: "Text"));
            double width = document.SelectedAnnotationWidth;

            document.SelectedAnnotationContents = "Draft that should be discarded";
            document.SelectedAnnotationAuthor = "Someone else";
            document.SelectedAnnotationWidth = width + 50;
            document.RevertSelectedAnnotationCommand.Execute(null);

            Assert.Equal("Original", document.SelectedAnnotationContents);
            Assert.Equal("Reviewer", document.SelectedAnnotationAuthor);
            Assert.Equal(width, document.SelectedAnnotationWidth);
            Assert.False(document.IsDirty);
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public async Task Escape_abandons_a_signature_or_date_waiting_to_be_placed() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var window = new Features.Shell.MainWindow();
            try {
                window.Show();
                await window.TabHost.OpenDocumentAsync(Path.Combine(AppContext.BaseDirectory, "Fixtures", "openpreserve-pdfa1b-text.pdf"));
                window.ViewModel.ShowFormsModeCommand.Execute(null);
                window.ViewModel.PlaceDateCommand.Execute(null);
                Assert.True(window.ViewModel.IsPlacingFillSignItem);

                // The page canvas uses Escape to clear its selection; the pending placement is abandoned first.
                window.Width = 1280; window.Height = 820;
                Avalonia.Threading.Dispatcher.UIThread.RunJobs();
                window.UpdateLayout();
                Avalonia.Threading.Dispatcher.UIThread.RunJobs();
                PdfPageCanvas canvas = Avalonia.VisualTree.VisualExtensions.GetVisualDescendants(window).OfType<PdfPageCanvas>().First();
                canvas.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = Key.Escape });

                Assert.False(window.ViewModel.IsPlacingFillSignItem);
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static void Press(PdfPageCanvas canvas, Key key, KeyModifiers modifiers) =>
        canvas.RaiseEvent(new KeyEventArgs { RoutedEvent = InputElement.KeyDownEvent, Key = key, KeyModifiers = modifiers });
}
