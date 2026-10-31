using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using PdfPageCanvas = OfficeIMO.Studio.Features.Reader.PdfPageCanvas;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioViewRotationInteractionTests {
    [Theory]
    [InlineData(90, false)] [InlineData(180, false)] [InlineData(270, false)]
    [InlineData(90, true)] [InlineData(180, true)] [InlineData(270, true)]
    public async Task MarqueeKeyboardMoveAndResizeUseDisplayedAxes(int rotation, bool independent) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "view-rotation-annotations.pdf");
            string saved = Path.Combine(services.Paths.Root, "view-rotation-edited.pdf");
            PdfDocument.Create(compose => compose.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(400), global::OfficeIMO.OfficeLength.Points(600))
                .Content(content => content.Text("Rotate the view and select this annotation."))))
                .Annotations.Add(new PdfAnnotationCreateOptions { PageNumber = 1, Subtype = "Square", Name = "target",
                    Rectangle = new[] { 170D, 270D, 230D, 330D }, GenerateAppearance = true }).ToDocument().Save(source);
            byte[] original = File.ReadAllBytes(source);
            var view = new DocumentWorkspaceView();
            using var host = new StudioDocumentTabHost(open => new MainWindowViewModel(_ => Task.FromResult<string?>(null),
                pickSavePdf: _ => Task.FromResult<string?>(saved), openDocumentInTab: open, services: services),
                document => view.DataContext = document);
            view.Panes = host.Panes;
            var window = new Window { Width = 1600, Height = 900, Content = view };
            try {
                window.Show();
                await host.OpenDocumentAsync(source);
                var document = host.ActiveDocument;
                document.ShowAnnotateModeCommand.Execute(null);
                document.SelectedReaderLayoutChoice = document.ReaderLayoutChoices.Single(choice => choice.Mode == ReaderLayoutMode.SinglePage);
                document.SetTouchZoom(1);
                for (int i = 0; i < rotation / 90; i++) document.RotateViewRightCommand.Execute(null);
                PdfPageViewModel presentation = document.Pages[0];
                if (independent) {
                    host.Panes.OpenSecondPane(host.SelectedTab);
                    var pane = host.Panes.Right!;
                    pane.Activate();
                    pane.SelectedLayout = pane.LayoutChoices.Single(choice => choice.Mode == ReaderLayoutMode.SinglePage);
                    pane.Zoom = 1;
                    presentation = pane.Pages[0];
                }
                async Task<PdfPageCanvas> Canvas() {
                    window.UpdateLayout();
                    await TextEditingReviewTests.WaitUntilAsync(() => window.GetVisualDescendants().OfType<PdfPageCanvas>()
                        .Any(canvas => canvas.IsEffectivelyVisible && ReferenceEquals(canvas.DataContext, presentation) && canvas.Scene is not null));
                    window.UpdateLayout();
                    return window.GetVisualDescendants().OfType<PdfPageCanvas>().Single(canvas =>
                        canvas.IsEffectivelyVisible && ReferenceEquals(canvas.DataContext, presentation) && canvas.Scene is not null);
                }
                PdfPageCanvas canvas = await Canvas();
                Point At(double x, double y) => canvas.TranslatePoint(new Point(x * canvas.Bounds.Width / canvas.Scene!.Drawing.Width,
                    y * canvas.Bounds.Height / canvas.Scene.Drawing.Height), window)!.Value;
                Point start = At(150, 250), end = At(250, 350);
                window.MouseDown(start, MouseButton.Left);
                window.MouseMove(end, RawInputModifiers.LeftMouseButton);
                Assert.NotNull(canvas.AnnotationMarqueePreview);
                window.MouseUp(end, MouseButton.Left);
                Assert.Single(document.SelectedAnnotations);
                PdfEditorVisualBounds bounds = canvas.SelectedObject!.Bounds;
                Point centerBefore = At(bounds.Left + bounds.Width / 2, bounds.Top + bounds.Height / 2);
                canvas.Focus();
                window.KeyPress(Key.Right, RawInputModifiers.Control, PhysicalKey.None, null);
                await TextEditingReviewTests.WaitUntilAsync(() => !document.IsWorkspaceBusy);
                Assert.False(document.HasError, document.ErrorMessage);
                presentation = independent ? host.Panes.Right!.Pages[0] : document.Pages[0];
                canvas = await Canvas();
                var annotation = canvas.Scene!.Interactions!.Regions.Single(region => region.Kind == PdfInteractionKind.Annotation);
                Point movedCenter = At((annotation.Quad.Left + annotation.Quad.Right) / 2, (annotation.Quad.Top + annotation.Quad.Bottom) / 2);
                Assert.NotNull(canvas.SelectedObject);
                bounds = canvas.SelectedObject!.Bounds;
                Point centerAfter = At(bounds.Left + bounds.Width / 2, bounds.Top + bounds.Height / 2);
                double pointScale = canvas.Bounds.Width / canvas.Scene!.Drawing.Width;
                Assert.InRange(centerAfter.X - centerBefore.X, pointScale - 0.02, pointScale + 0.02);
                Assert.InRange(centerAfter.Y - centerBefore.Y, -0.02, 0.02);
                double displayedWidth = rotation % 180 == 90 ? bounds.Height : bounds.Width;
                canvas.Focus();
                window.KeyPress(Key.Right, RawInputModifiers.Control | RawInputModifiers.Alt, PhysicalKey.None, null);
                await TextEditingReviewTests.WaitUntilAsync(() => !document.IsWorkspaceBusy);
                Assert.False(document.HasError, document.ErrorMessage);
                presentation = independent ? host.Panes.Right!.Pages[0] : document.Pages[0];
                canvas = await Canvas();
                annotation = canvas.Scene!.Interactions!.Regions.Single(region => region.Kind == PdfInteractionKind.Annotation);
                Assert.InRange((rotation % 180 == 90 ? annotation.Quad.Height : annotation.Quad.Width) - displayedWidth, 0.99, 1.01);
                await document.SaveAsCommand.ExecuteAsync(null);
                Assert.Equal(0, PdfDocument.Load(saved).Inspect().Pages[0].RotationDegrees);
                Assert.Equal(original, File.ReadAllBytes(source));
                await document.UndoCommand.ExecuteAsync(null);
                await document.UndoCommand.ExecuteAsync(null);
                await document.SaveAsCommand.ExecuteAsync(null);
                Assert.Equal(original, File.ReadAllBytes(saved));
            } finally { window.Close(); }
            return true;
        }, CancellationToken.None);
    }
}
