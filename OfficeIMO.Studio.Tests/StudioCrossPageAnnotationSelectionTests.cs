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

public sealed class StudioCrossPageAnnotationSelectionTests {
    [Fact]
    public async Task AnnotationOnlyUserPasswordAllowsCrossPageSelectionMoveCopyAndUndoWithoutTextExtraction() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "annotation-only.pdf");
            string saved = Path.Combine(services.Paths.Root, "annotation-only-edited.pdf");
            var options = new PdfLoadOptions { Password = "open" };
            var ownerOptions = new PdfLoadOptions { Password = "owner" }; // Independent artifact oracle; Studio receives only the user password.
            byte[] encrypted = PdfDocument.Load(CreateSource()).Security.Encrypt(new PdfStandardEncryptionOptions("open") {
                OwnerPassword = "owner", AllowedPermissions = PdfStandardPermissions.ModifyAnnotations
            }).Pdf;
            File.WriteAllBytes(source, encrypted);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickSavePdf: _ => Task.FromResult<string?>(saved), promptPdfPassword: (_, _, _) => Task.FromResult<string?>("open"), services: services);
            await model.OpenDocumentAsync(source);
            model.ShowAnnotateModeCommand.Execute(null);
            Assert.True(model.CanEditAnnotations);
            var snapshot = PdfDocument.Load(encrypted, options);
            var inspected = PdfDocument.Load(encrypted, ownerOptions).Inspect();
            var layouts = snapshot.GetPageLayouts();
            model.SelectedReaderLayoutChoice = model.ReaderLayoutChoices.Single(choice => choice.Mode == ReaderLayoutMode.Grid);
            model.SetTouchZoom(0.5);
            var window = new Window { Width = 1280, Height = 900, Content = new DocumentWorkspaceView { DataContext = model } };
            try {
                window.Show();
                async Task Render() {
                    window.UpdateLayout();
                    foreach (var page in model.Pages) await page.EnsureRenderedAsync();
                    window.UpdateLayout();
                }
                await Render();
                PdfPageCanvas Canvas(int page) => window.GetVisualDescendants().OfType<PdfPageCanvas>().Single(canvas =>
                    canvas.IsEffectivelyVisible && canvas.Scene?.PageNumber == page && canvas.SelectionMode == PdfEditorSelectionMode.Annotations);
                void SelectBoth() {
                    foreach (var annotation in inspected.Annotations.Where(item => item.Name!.StartsWith("selected-", StringComparison.Ordinal))) {
                        int page = annotation.PageNumber!.Value;
                        var quad = layouts[page - 1].MapUserSpaceRectangleToVisual(annotation.X1, annotation.Y1, annotation.X2, annotation.Y2);
                        var canvas = Canvas(page);
                        Assert.Empty(canvas.Scene!.Interactions!.TextRegions);
                        Assert.All(canvas.Scene.Interactions.Regions, region => Assert.Null(region.Text));
                        Point local = new((quad.Left + quad.Right) / 2 * canvas.Bounds.Width / canvas.Scene.Drawing.Width,
                            (quad.Top + quad.Bottom) / 2 * canvas.Bounds.Height / canvas.Scene.Drawing.Height);
                        Point point = canvas.TranslatePoint(local, window)!.Value;
                        RawInputModifiers modifiers = page > 1 ? RawInputModifiers.Shift : RawInputModifiers.None;
                        window.MouseDown(point, MouseButton.Left, modifiers); window.MouseUp(point, MouseButton.Left, modifiers);
                    }
                }
                SelectBoth();
                Assert.Null(model.ErrorMessage);
                Assert.Equal(2, model.SelectedAnnotations.Count);
                Capture(window, "cross-page-encrypted-selection.png");
                Assert.False(model.HasSelectedText);
                var selected = model.Pages[1].SelectedObject!;
                var bounds = selected.Bounds;
                model.Pages[1].TransformObject(new(selected, new(bounds.Left + 11, bounds.Top + 7, bounds.Right + 11, bounds.Bottom + 7), model.SelectedAnnotations));
                await WaitForMutation(model);
                await model.SaveAsCommand.ExecuteAsync(null);
                var moved = PdfDocument.Load(File.ReadAllBytes(saved), ownerOptions).Inspect();
                foreach (var original in inspected.Annotations.Where(item => item.Name!.StartsWith("selected-", StringComparison.Ordinal))) {
                    var changed = moved.Annotations.Single(item => item.Name == original.Name);
                    var page = layouts[original.PageNumber!.Value - 1];
                    var a = page.MapUserSpaceRectangleToVisual(original.X1, original.Y1, original.X2, original.Y2);
                    var b = page.MapUserSpaceRectangleToVisual(changed.X1, changed.Y1, changed.X2, changed.Y2);
                    Assert.Equal(a.Left + 11, b.Left, 7); Assert.Equal(a.Top + 7, b.Top, 7);
                }
                Assert.Throws<PdfPermissionDeniedException>(() => PdfDocument.Load(File.ReadAllBytes(saved), options).Read());
                await model.UndoCommand.ExecuteAsync(null);
                await model.SaveAsCommand.ExecuteAsync(null);
                Assert.Equal(encrypted, File.ReadAllBytes(saved));
                await Render();
                SelectBoth();
                await model.CopyAnnotationsCommand.ExecuteAsync(null);
                Assert.Null(model.ErrorMessage);
                await model.SaveAsCommand.ExecuteAsync(null);
                Assert.Equal(6, PdfDocument.Load(File.ReadAllBytes(saved), ownerOptions).Inspect().Annotations.Count);
                Assert.Throws<PdfPermissionDeniedException>(() => PdfDocument.Load(File.ReadAllBytes(saved), options).Read());
                await model.UndoCommand.ExecuteAsync(null);
                await model.SaveAsCommand.ExecuteAsync(null);
                Assert.Equal(encrypted, File.ReadAllBytes(saved));
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    [Theory]
    [InlineData(ReaderLayoutMode.Continuous, 960, false)]
    [InlineData(ReaderLayoutMode.Grid, 1280, true)]
    public async Task VisiblePageMarqueeAndAdditiveSelectionEditBothRotatedPagesAsOneAction(ReaderLayoutMode layout, int width, bool dark) {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(async () => {
            Application.Current!.RequestedThemeVariant = dark ? ThemeVariant.Dark : ThemeVariant.Light;
            var services = ((App)Application.Current).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "cross-page-annotations.pdf");
            string saved = Path.Combine(services.Paths.Root, "cross-page-edited.pdf");
            File.WriteAllBytes(source, CreateSource());
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickSavePdf: _ => Task.FromResult<string?>(saved), services: services);
            await model.OpenDocumentAsync(source);
            model.ShowAnnotateModeCommand.Execute(null);
            model.SelectedReaderLayoutChoice = model.ReaderLayoutChoices.Single(choice => choice.Mode == layout);
            model.SetTouchZoom(0.5); // Both complete pages fit in the visible reader viewport.
            var view = new DocumentWorkspaceView { DataContext = model };
            var window = new Window { Width = width, Height = 900, Content = view };
            try {
                window.Show();
                async Task Render() {
                    window.UpdateLayout();
                    foreach (var page in model.Pages) await page.EnsureRenderedAsync();
                    window.UpdateLayout();
                }
                await Render();
                PdfPageCanvas Canvas(int pageNumber) => window.GetVisualDescendants().OfType<PdfPageCanvas>().Single(canvas =>
                    canvas.IsEffectivelyVisible && canvas.Scene?.PageNumber == pageNumber && canvas.SelectionMode == PdfEditorSelectionMode.Annotations);
                Point At(int page, double x, double y) {
                    var canvas = Canvas(page);
                    return canvas.TranslatePoint(new Point(x * canvas.Bounds.Width / canvas.Scene!.Drawing.Width,
                        y * canvas.Bounds.Height / canvas.Scene.Drawing.Height), window)!.Value;
                }
                void Click(int page, double x, double y, RawInputModifiers modifiers = RawInputModifiers.None) {
                    Point point = At(page, x, y);
                    window.MouseDown(point, MouseButton.Left, modifiers); window.MouseUp(point, MouseButton.Left, modifiers);
                }
                void SelectBoth() { Click(1, 70, 90); Click(2, 260, 70, RawInputModifiers.Shift); }
                SelectBoth();
                Assert.Equal(new[] { 1, 2 }, model.SelectedAnnotations.Select(item => item.PageNumber).Order().ToArray());
                Assert.True(model.HasCrossPageAnnotationSelection);
                Assert.False(model.CanGroupAnnotations);
                Assert.False(model.CanResizeSelectedAnnotation);
                Assert.False(model.HasSelectedText);
                Assert.NotNull(Canvas(1).SelectedObject);
                Assert.NotNull(Canvas(2).SelectedObject);
                Assert.Equal(40, Canvas(1).SelectedObject!.Bounds.Left);
                Assert.Equal(230, Canvas(2).SelectedObject!.Bounds.Left);
                await model.GroupAnnotationsCommand.ExecuteAsync(null);
                Assert.Equal(2, model.SelectedAnnotations.Count); // Disabled Group cannot mutate a cross-page set.
                Capture(window, $"cross-page-selected-{layout}-{width}.png");

                // A key sent to the second page moves the whole selection, using each page's own rotation.
                Canvas(2).Focus();
                window.KeyPress(Key.Right, RawInputModifiers.Control, PhysicalKey.None, null);
                await WaitForMutation(model);
                await model.SaveAsCommand.ExecuteAsync(null);
                AssertVisualMovement(source, saved, 1, 0, tolerance: 0.000001);
                await model.UndoCommand.ExecuteAsync(null);
                await model.SaveAsCommand.ExecuteAsync(null);
                AssertVisualMovement(source, saved, 0, 0, tolerance: 0.000001);
                await Render();

                SelectBoth();
                Point start = At(2, 260, 70), end = At(2, 277, 79);
                window.MouseDown(start, MouseButton.Left);
                window.MouseMove(end, RawInputModifiers.LeftMouseButton);
                window.MouseUp(end, MouseButton.Left);
                await WaitForMutation(model);
                await model.SaveAsCommand.ExecuteAsync(null);
                AssertVisualMovement(source, saved, 17, 9, tolerance: 2);
                await model.UndoCommand.ExecuteAsync(null);
                await model.SaveAsCommand.ExecuteAsync(null);
                AssertVisualMovement(source, saved, 0, 0, tolerance: 0.000001);
                await Render();

                SelectBoth();
                model.ObjectMoveX = 6; model.ObjectMoveY = 4;
                await model.MoveSelectedObjectCommand.ExecuteAsync(null);
                await model.SaveAsCommand.ExecuteAsync(null);
                AssertVisualMovement(source, saved, 6, 4, tolerance: 0.000001);
                await model.UndoCommand.ExecuteAsync(null);
                await model.SaveAsCommand.ExecuteAsync(null);
                AssertVisualMovement(source, saved, 0, 0, tolerance: 0.000001);
                await Render();

                void Marquee() {
                    Click(1, 10, 10);
                    Point a = At(1, 20, 20), b = At(2, 310, 150);
                    window.MouseDown(a, MouseButton.Left);
                    window.MouseMove(b, RawInputModifiers.LeftMouseButton);
                    Assert.NotNull(Canvas(1).AnnotationMarqueePreview);
                    Assert.NotNull(Canvas(2).AnnotationMarqueePreview);
                    Capture(window, $"cross-page-marquee-{layout}-{width}.png");
                    window.MouseUp(b, MouseButton.Left);
                }
                Marquee();
                Assert.Equal(2, model.SelectedAnnotations.Count);
                Assert.True(model.HasCrossPageAnnotationSelection);
                Assert.All(new[] { Canvas(1), Canvas(2) }, canvas => Assert.Null(canvas.AnnotationMarqueePreview));
                Assert.False(model.HasSelectedText);

                // Add/remove with keyboard focus on another page preserves the other page's member.
                Canvas(2).Focus();
                window.KeyPress(Key.Home, RawInputModifiers.None, PhysicalKey.None, null);
                window.KeyPress(Key.Space, RawInputModifiers.Shift, PhysicalKey.None, null);
                Assert.Single(model.SelectedAnnotations);
                Assert.Equal(1, model.SelectedAnnotations[0].PageNumber);
                window.KeyPress(Key.Space, RawInputModifiers.Shift, PhysicalKey.None, null);
                Assert.Equal(2, model.SelectedAnnotations.Count);
                await model.CopyAnnotationsCommand.ExecuteAsync(null);
                await model.SaveAsCommand.ExecuteAsync(null);
                var copies = PdfDocument.Load(File.ReadAllBytes(saved));
                Assert.Equal(6, copies.Inspect().Annotations.Count);
                Assert.All(copies.Inspect().Annotations.GroupBy(item => item.PageNumber), group => Assert.Equal(3, group.Count()));
                var pages = copies.Read(new PdfReadOptions { Profile = PdfReadProfile.Fast }).Pages;
                foreach (int page in new[] { 1, 2 }) {
                    var original = copies.Inspect().Annotations.Single(item => item.Name == $"selected-{page}");
                    var copy = copies.Inspect().Annotations.Single(item => item.PageNumber == page && item.Name != $"selected-{page}" && item.Name != $"other-{page}");
                    var a = pages[page - 1].MapUserSpaceRectangleToVisual(original.X1, original.Y1, original.X2, original.Y2);
                    var b = pages[page - 1].MapUserSpaceRectangleToVisual(copy.X1, copy.Y1, copy.X2, copy.Y2);
                    Assert.Equal(a.Left + 10, b.Left, 7); Assert.Equal(a.Top + 10, b.Top, 7);
                }
                await model.UndoCommand.ExecuteAsync(null);
                await Render();
                Marquee();
                await model.RaiseAnnotationsCommand.ExecuteAsync(null);
                await model.SaveAsCommand.ExecuteAsync(null);
                Assert.All(PdfDocument.Load(File.ReadAllBytes(saved)).Inspect().Annotations.GroupBy(item => item.PageNumber),
                    group => Assert.StartsWith("selected-", group.Last().Name));
                await Render();
                Marquee();
                await model.LowerAnnotationsCommand.ExecuteAsync(null);
                await model.SaveAsCommand.ExecuteAsync(null);
                Assert.All(PdfDocument.Load(File.ReadAllBytes(saved)).Inspect().Annotations.GroupBy(item => item.PageNumber),
                    group => Assert.StartsWith("selected-", group.First().Name));
                await model.UndoCommand.ExecuteAsync(null);
                await model.SaveAsCommand.ExecuteAsync(null);
                Assert.All(PdfDocument.Load(File.ReadAllBytes(saved)).Inspect().Annotations.GroupBy(item => item.PageNumber),
                    group => Assert.StartsWith("selected-", group.Last().Name));
                await model.UndoCommand.ExecuteAsync(null);
                await Render();
                Marquee();
                Canvas(2).Focus(); window.KeyPress(Key.Delete, RawInputModifiers.None, PhysicalKey.None, null);
                await WaitForMutation(model);
                await model.SaveAsCommand.ExecuteAsync(null);
                Assert.All(PdfDocument.Load(File.ReadAllBytes(saved)).Inspect().Annotations, item => Assert.StartsWith("other-", item.Name));
                await model.UndoCommand.ExecuteAsync(null);
                await model.SaveAsCommand.ExecuteAsync(null);
                Assert.Equal(4, PdfDocument.Load(File.ReadAllBytes(saved)).Inspect().Annotations.Count);
                await Render();
                Marquee();
                Capture(window, $"cross-page-final-{layout}-{width}.png");
                // Escape cancels the shared preview and releases capture for the next click.
                Click(1, 10, 10);
                Point cancelStart = At(1, 20, 20), cancelEnd = At(2, 310, 150);
                window.MouseDown(cancelStart, MouseButton.Left);
                window.MouseMove(cancelEnd, RawInputModifiers.LeftMouseButton);
                window.KeyPress(Key.Escape, RawInputModifiers.None, PhysicalKey.None, null);
                window.MouseUp(cancelEnd, MouseButton.Left);
                Assert.Empty(model.SelectedAnnotations);
                Assert.All(new[] { Canvas(1), Canvas(2) }, canvas => Assert.Null(canvas.AnnotationMarqueePreview));
                SelectBoth();
                Assert.Equal(2, model.SelectedAnnotations.Count);
                if (Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT") is { Length: > 0 } output) {
                    File.Copy(source, Path.Combine(output, "cross-page-annotations.pdf"), true);
                    File.Copy(saved, Path.Combine(output, "cross-page-edited.pdf"), true);
                }
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }

    private static async Task WaitForMutation(MainWindowViewModel model) {
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(15));
        while (model.IsWorkspaceBusy) await Task.Delay(10, timeout.Token);
        Assert.Null(model.ErrorMessage);
    }

    private static void AssertVisualMovement(string source, string saved, double deltaX, double deltaY, double tolerance) {
        var original = PdfDocument.Load(File.ReadAllBytes(source));
        var changed = PdfDocument.Load(File.ReadAllBytes(saved));
        var pages = original.Read(new PdfReadOptions { Profile = PdfReadProfile.Fast }).Pages;
        foreach (var annotation in original.Inspect().Annotations) {
            var moved = changed.Inspect().Annotations.Single(item => item.Name == annotation.Name);
            var page = pages[annotation.PageNumber!.Value - 1];
            var a = page.MapUserSpaceRectangleToVisual(annotation.X1, annotation.Y1, annotation.X2, annotation.Y2);
            var b = page.MapUserSpaceRectangleToVisual(moved.X1, moved.Y1, moved.X2, moved.Y2);
            bool selected = annotation.Name!.StartsWith("selected-", StringComparison.Ordinal);
            Assert.InRange(b.Left - a.Left, (selected ? deltaX : 0) - tolerance, (selected ? deltaX : 0) + tolerance);
            Assert.InRange(b.Top - a.Top, (selected ? deltaY : 0) - tolerance, (selected ? deltaY : 0) + tolerance);
            Assert.Equal(a.Width, b.Width, 7); Assert.Equal(a.Height, b.Height, 7);
        }
    }

    private static byte[] CreateSource() {
        var document = PdfDocument.Create(compose => {
            for (int page = 1; page <= 2; page++) compose.Page(builder => builder.Size(400, 400).Content(content =>
                content.Item(item => item.Paragraph(paragraph => paragraph.Text("Cross-page annotation workspace")))));
        });
        document = document.Pages.SetCropBox(20, 30, 380, 390, 1).Pages.SetCropBox(30, 40, 390, 400, 2).Pages.Rotate(90, 2);
        var pages = document.Read(new PdfReadOptions { Profile = PdfReadProfile.Fast }).Pages;
        foreach (int page in new[] { 1, 2 }) {
            var selected = page == 1 ? new PdfEditorVisualBounds(40, 60, 100, 120) : new PdfEditorVisualBounds(230, 40, 290, 100);
            foreach (var (name, bounds, subtype) in new[] { ($"selected-{page}", selected, "Square"), ($"other-{page}", new PdfEditorVisualBounds(320, 250, 340, 270), "Circle") }) {
                var rectangle = pages[page - 1].MapVisualRectangleToUserSpace(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom);
                document = document.Annotations.Add(new PdfAnnotationCreateOptions { PageNumber = page, Subtype = subtype, Name = name,
                    Rectangle = new[] { rectangle.Left, rectangle.Bottom, rectangle.Right, rectangle.Top },
                    Color = new[] { 0.2D, 0.4D, 0.9D }, GenerateAppearance = true }).ToDocument();
            }
        }
        return document.ToBytes();
    }

    private static void Capture(Window window, string name) {
        if (Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT") is not { Length: > 0 } output) return;
        Directory.CreateDirectory(output);
        using var frame = window.CaptureRenderedFrame();
        Assert.NotNull(frame); frame.Save(Path.Combine(output, name), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
