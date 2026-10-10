using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Styling;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using PdfPageCanvas = OfficeIMO.Studio.Features.Reader.PdfPageCanvas;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioAnnotationSelectionTests {
    [Theory]
    [InlineData(960, false)]
    [InlineData(1280, true)]
    public async Task PointerMultiSelectionGroupMoveAndUndoSaveAsOneAction(int width, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            Application.Current!.RequestedThemeVariant = dark ? ThemeVariant.Dark : ThemeVariant.Light;
            var services = ((App)Application.Current).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "annotations.pdf");
            string saved = Path.Combine(services.Paths.Root, "grouped.pdf");
            File.WriteAllBytes(source, CreateSource());
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickSavePdf: _ => Task.FromResult<string?>(saved), services: services);
            await model.OpenDocumentAsync(source);
            model.ShowAnnotateModeCommand.Execute(null);
            var window = new Window { Width = width, Height = 760, Content = new DocumentWorkspaceView { DataContext = model } };
            try {
                window.Show(); window.UpdateLayout();
                await model.SelectedPage!.EnsureRenderedAsync();
                window.UpdateLayout();
                var canvas = window.GetVisualDescendants().OfType<PdfPageCanvas>().First(control =>
                    control.Scene?.PageNumber == 1 && control.SelectionMode == PdfEditorSelectionMode.Annotations);
                Point At(double x, double y) => canvas.TranslatePoint(new Point(x * canvas.Bounds.Width / canvas.Scene!.Drawing.Width,
                    y * canvas.Bounds.Height / canvas.Scene.Drawing.Height), window)!.Value;
                void Click(Point point, RawInputModifiers modifiers = RawInputModifiers.None) {
                    window.MouseDown(point, MouseButton.Left, modifiers); window.MouseUp(point, MouseButton.Left, modifiers);
                }
                Click(At(100, 170));
                Assert.Single(model.SelectedAnnotations);
                Click(At(200, 170), RawInputModifiers.Shift);
                Assert.Equal(2, model.SelectedAnnotations.Count);
                Assert.False(model.HasSelectedAnnotation); // Individual property edits must not silently affect one item.
                Assert.True(model.CanGroupAnnotations);
                Click(At(280, 240), RawInputModifiers.Shift);
                Assert.Equal(2, model.SelectedAnnotations.Count);
                Capture(window, $"annotation-selection-{width}-{(dark ? "dark" : "light")}.png");

                await model.GroupAnnotationsCommand.ExecuteAsync(null);
                Assert.Null(model.ErrorMessage);
                await model.SaveAsCommand.ExecuteAsync(null);
                var grouped = PdfDocument.Load(File.ReadAllBytes(saved));
                Assert.Single(grouped.Inspect().Annotations, annotation => annotation.Review?.IsGroup == true);
                var annotation = grouped.Inspect().Annotations.First(item => item.Name == "square");
                model.Pages[0].SelectObject(new(PdfEditorSelectionKind.Annotation, 1, new(60, 140, 140, 200), ObjectNumber: annotation.ObjectNumber, Subtype: "Square"));
                Assert.Equal(2, model.SelectedAnnotations.Count);
                Assert.True(model.CanUngroupAnnotations);
                window.UpdateLayout(); await model.SelectedPage!.EnsureRenderedAsync(); window.UpdateLayout();
                canvas = window.GetVisualDescendants().OfType<PdfPageCanvas>().First(control =>
                    control.Scene?.PageNumber == 1 && control.SelectionMode == PdfEditorSelectionMode.Annotations);
                window.MouseDown(At(100, 170), MouseButton.Left);
                window.MouseMove(At(125, 160), RawInputModifiers.LeftMouseButton);
                window.MouseUp(At(125, 160), MouseButton.Left);
                using (var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10))) {
                    while (model.IsWorkspaceBusy) await Task.Delay(10, timeout.Token);
                }
                Assert.Null(model.ErrorMessage);
                await model.SaveAsCommand.ExecuteAsync(null);
                var moved = PdfDocument.Load(File.ReadAllBytes(saved)).Inspect().Annotations.Where(item => item.Name is "square" or "circle").ToArray();
                // Headless pointer delivery rounds device pixels; both members must receive the same translation.
                Assert.InRange(moved.Single(item => item.Name == "square").X1, 84, 86);
                Assert.Equal(100, moved.Single(item => item.Name == "circle").X1 - moved.Single(item => item.Name == "square").X1, 7);
                await model.UndoCommand.ExecuteAsync(null);
                await model.SaveAsCommand.ExecuteAsync(null);
                var undone = PdfDocument.Load(File.ReadAllBytes(saved)).Inspect().Annotations;
                Assert.Equal(60, undone.Single(item => item.Name == "square").X1);
                Assert.Equal(160, undone.Single(item => item.Name == "circle").X1);
                Assert.Single(undone, item => item.Review?.IsGroup == true);

                // Rubber-band selects annotation rectangles; it does not select page text or organizer pages.
                window.UpdateLayout(); await model.SelectedPage!.EnsureRenderedAsync(); window.UpdateLayout();
                canvas = window.GetVisualDescendants().OfType<PdfPageCanvas>().First(control =>
                    control.Scene?.PageNumber == 1 && control.SelectionMode == PdfEditorSelectionMode.Annotations);
                Point start = At(45, 125), end = At(255, 220);
                window.MouseDown(start, MouseButton.Left); window.MouseMove(end, RawInputModifiers.LeftMouseButton); window.MouseUp(end, MouseButton.Left);
                Assert.Equal(2, model.SelectedAnnotations.Count);
                Assert.False(model.HasSelectedText);
                Click(At(100, 170), RawInputModifiers.Shift);
                Assert.Empty(model.SelectedAnnotations); // Toggling one member toggles the entire durable group.
                canvas.Focus();
                window.KeyPress(Key.Home, RawInputModifiers.None, PhysicalKey.None, null);
                window.KeyPress(Key.Space, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.Equal(2, model.SelectedAnnotations.Count);
                window.KeyPress(Key.End, RawInputModifiers.None, PhysicalKey.None, null);
                window.KeyPress(Key.Space, RawInputModifiers.Shift, PhysicalKey.None, null);
                Assert.Equal(3, model.SelectedAnnotations.Count);
                window.KeyPress(Key.Space, RawInputModifiers.Shift, PhysicalKey.None, null);
                Assert.Equal(2, model.SelectedAnnotations.Count);
                window.KeyPress(Key.Right, RawInputModifiers.Control, PhysicalKey.None, null);
                using (var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(10))) {
                    while (model.IsWorkspaceBusy) await Task.Delay(10, timeout.Token);
                }
                Assert.Null(model.ErrorMessage);
                await model.SaveAsCommand.ExecuteAsync(null);
                var nudged = PdfDocument.Load(File.ReadAllBytes(saved)).Inspect().Annotations;
                Assert.Equal(61, nudged.Single(item => item.Name == "square").X1);
                Assert.Equal(161, nudged.Single(item => item.Name == "circle").X1);
                await model.UndoCommand.ExecuteAsync(null);
                await model.SaveAsCommand.ExecuteAsync(null);
                window.UpdateLayout(); await model.SelectedPage!.EnsureRenderedAsync(); window.UpdateLayout();
                canvas = window.GetVisualDescendants().OfType<PdfPageCanvas>().First(control =>
                    control.Scene?.PageNumber == 1 && control.SelectionMode == PdfEditorSelectionMode.Annotations);
                Click(At(100, 170));
                Assert.Equal(2, model.SelectedAnnotations.Count);
                Assert.True(model.CanUngroupAnnotations);
                Capture(window, $"annotation-group-{width}-{(dark ? "dark" : "light")}.png");
                if (Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT") is { Length: > 0 } output) {
                    File.Copy(source, Path.Combine(output, "annotations.pdf"), true);
                    File.Copy(saved, Path.Combine(output, "grouped.pdf"), true);
                }
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Fact]
    public async Task WorkspaceBatchMutationIsAtomicAndStaleSelectionDoesNotChangeDocument() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-annotation-batch-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "annotations.pdf"); File.WriteAllBytes(path, CreateSource());
            using var workspace = await PdfWorkspace.OpenAsync(path, CancellationToken.None, new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery")));
            var selections = workspace.DocumentInfo!.Annotations.Take(2).Select(annotation => new PdfEditorSelection(PdfEditorSelectionKind.Annotation,
                1, new(0, 0, 10, 10), ObjectNumber: annotation.ObjectNumber, Subtype: annotation.Subtype)).ToArray();
            byte[] original = workspace.CopyBytes();
            await workspace.EditAnnotationsAsync(selections, workspace.Revision, (editor, numbers) => editor.CopyMany(numbers), "Duplicate", CancellationToken.None);
            Assert.Equal(5, workspace.DocumentInfo!.Annotations.Count);
            Assert.Single(workspace.Journal);
            await workspace.UndoAsync(CancellationToken.None);
            Assert.Equal(original, workspace.CopyBytes());
            Assert.False(workspace.CanUndo);
            await Assert.ThrowsAsync<InvalidOperationException>(() => workspace.EditAnnotationsAsync(selections, revision: 99,
                (editor, numbers) => editor.RemoveMany(numbers), "Delete", CancellationToken.None));
            Assert.Equal(original, workspace.CopyBytes());
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Theory]
    [InlineData(0, 13, -7)]
    [InlineData(90, -7, -13)]
    [InlineData(180, -13, 7)]
    [InlineData(270, 7, 13)]
    public async Task MultiSelectionGestureMapsVisualMovementOnCroppedRotatedPages(int rotation, double expectedX, double expectedY) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-annotation-geometry-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string path = Path.Combine(root, "annotations.pdf");
            PdfDocument.Load(CreateSource()).Pages.SetCropBox(20, 30, 580, 770, 1).Pages.Rotate(rotation, 1).Save(path);
            using var workspace = await PdfWorkspace.OpenAsync(path, CancellationToken.None, new PdfWorkspaceRecoveryStore(Path.Combine(root, "recovery")));
            var before = workspace.DocumentInfo!.Annotations.Take(2).ToArray();
            var page = workspace.CreateDocumentSnapshot().Read(new PdfReadOptions { Profile = PdfReadProfile.Fast }).Pages[0];
            var selections = before.Select(annotation => {
                var quad = page.MapUserSpaceRectangleToVisual(annotation.X1, annotation.Y1, annotation.X2, annotation.Y2);
                return new PdfEditorSelection(PdfEditorSelectionKind.Annotation, 1, new(quad.Left, quad.Top, quad.Right, quad.Bottom),
                    ObjectNumber: annotation.ObjectNumber, Subtype: annotation.Subtype);
            }).ToArray();
            var bounds = new PdfEditorVisualBounds(selections.Min(item => item.Bounds.Left), selections.Min(item => item.Bounds.Top),
                selections.Max(item => item.Bounds.Right), selections.Max(item => item.Bounds.Bottom));
            var gesture = new PdfObjectTransformGesture(selections[0] with { Bounds = bounds },
                new(bounds.Left + 13, bounds.Top + 7, bounds.Right + 13, bounds.Bottom + 7), selections);
            await workspace.TransformSelectedObjectAsync(gesture, workspace.Revision, PdfImageEditLayer.AboveExistingContent, CancellationToken.None, null);
            Assert.Single(workspace.Journal);
            foreach (var original in before) {
                var moved = workspace.DocumentInfo!.Annotations.Single(item => item.Name == original.Name);
                Assert.Equal(original.X1 + expectedX, moved.X1, 7);
                Assert.Equal(original.Y1 + expectedY, moved.Y1, 7);
                Assert.Equal(original.Width, moved.Width, 7);
                Assert.Equal(original.Height, moved.Height, 7);
            }
            await workspace.UndoAsync(CancellationToken.None);
            Assert.Equal(File.ReadAllBytes(path), workspace.CopyBytes());
        } finally { Directory.Delete(root, recursive: true); }
    }

    private static byte[] CreateSource() {
        var document = PdfDocument.Create(compose => compose.Page(page => page.Size(600, 800).Content(content =>
            content.Item(item => item.Paragraph(paragraph => paragraph.Text("Annotation selection workspace"))))));
        document = document.Annotations.Add(new PdfAnnotationCreateOptions { Subtype = "Square", Name = "square", Contents = "Review box",
            Rectangle = new[] { 60D, 600D, 140D, 660D }, Color = new[] { 0.9D, 0.2D, 0.2D }, GenerateAppearance = true }).ToDocument();
        document = document.Annotations.Add(new PdfAnnotationCreateOptions { Subtype = "Circle", Name = "circle", Contents = "Review circle",
            Rectangle = new[] { 160D, 600D, 240D, 660D }, Color = new[] { 0.2D, 0.4D, 0.9D }, GenerateAppearance = true }).ToDocument();
        return document.Annotations.Add(new PdfAnnotationCreateOptions { Subtype = "Line", Name = "line", Contents = "Separate line",
            Rectangle = new[] { 120D, 480D, 220D, 500D }, Line = new[] { 125D, 485D, 215D, 495D }, GenerateAppearance = true }).Bytes;
    }

    private static void Capture(Window window, string name) {
        if (Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT") is not { Length: > 0 } output) return;
        Directory.CreateDirectory(output);
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame); frame.Save(Path.Combine(output, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
