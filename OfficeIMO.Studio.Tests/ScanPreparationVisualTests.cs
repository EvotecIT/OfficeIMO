using Avalonia;
using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.VisualTree;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class ScanPreparationVisualTests {
    [Fact]
    public async Task PageAndDocumentChangesClearVisualSelectionsAndDisabledPerspectiveStaysDisabled() {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            string source = Path.Combine(services.Paths.Root, "distinct-scan-pages.pdf");
            PdfDocument.Create(c => {
                c.Page(p => p.Size(240, 320).Content(content => content.Item(i => i.Paragraph(t => t.Text("FIRST PORTRAIT PAGE")))));
                c.Page(p => p.Size(320, 240).Content(content => content.Item(i => i.Paragraph(t => t.Text("SECOND LANDSCAPE PAGE")))));
            }).Save(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(source), services: services);
            model.OcrWorkbench.UseDocument(source);
            var scan = model.OcrWorkbench.Scan;
            scan.Dpi = 72;
            var view = new SearchablePdfOcrView { DataContext = model };
            var window = new Window { Width = 960, Height = 640, Content = view };
            try {
                window.Show(); window.UpdateLayout();
                Assert.All(view.GetVisualDescendants().OfType<NumericUpDown>(), input => Assert.True(input.IsEffectivelyEnabled));
                await scan.PreviewCommand.ExecuteAsync(null);
                Assert.True(scan.IsCurrent, scan.Status);
                Assert.Equal(240, scan.SourcePreview!.PixelSize.Width);
                scan.Region = new Rect(.1, .1, .5, .5); scan.UseRegion = true;
                scan.UsePerspective = true; scan.EditCorners = true;
                scan.TopLeft = new Point(.05, .05);
                scan.PageNumber = 2;
                Assert.Null(scan.SourcePreview); Assert.Null(scan.PreparedPreview);
                Assert.False(scan.CanSelectRegion); Assert.False(scan.IsCurrent);
                window.UpdateLayout();
                Assert.All(view.GetVisualDescendants().OfType<NumericUpDown>(), input => Assert.True(input.IsEffectivelyEnabled));
                Assert.False(scan.UseRegion); Assert.False(scan.UsePerspective); Assert.False(scan.EditCorners);
                Assert.Empty(scan.ApplyTo(new()).Regions);
                await scan.PreviewCommand.ExecuteAsync(null);
                Assert.True(scan.IsCurrent, scan.Status);
                Assert.Equal(320, scan.SourcePreview!.PixelSize.Width);
                Assert.Equal(240, scan.SourcePreview.PixelSize.Height);
                window.UpdateLayout();
                var canvas = view.GetVisualDescendants().OfType<ScanSelectionCanvas>().Single();
                canvas.BringIntoView(); window.UpdateLayout();
                Capture(window, "scan-page-two-selection-reset");
                scan.UsePerspective = true; scan.EditCorners = true;
                scan.UsePerspective = false;
                Assert.False(scan.EditCorners);
                canvas.Focus();
                window.KeyPress(Key.Right, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.False(scan.UsePerspective); Assert.True(scan.UseRegion);
                double scale = Math.Min(canvas.Bounds.Width / 320, canvas.Bounds.Height / 240);
                Point At(double x, double y) => canvas.TranslatePoint(new Point(
                    (canvas.Bounds.Width - 320 * scale) / 2 + 320 * scale * x,
                    (canvas.Bounds.Height - 240 * scale) / 2 + 240 * scale * y), window)!.Value;
                window.MouseDown(At(.2, .2), MouseButton.Left);
                window.MouseMove(At(.7, .7), RawInputModifiers.LeftMouseButton);
                window.MouseUp(At(.7, .7), MouseButton.Left);
                Assert.False(scan.UsePerspective);
                Assert.Equal(.5, scan.Region.Width, 4);
                Assert.Equal(2, Assert.Single(scan.ApplyTo(new()).Regions).PageNumber);
                Capture(window, "scan-perspective-disabled-region");
                scan.UsePerspective = true;
                Assert.False(scan.EditCorners);
                scan.EditCorners = true;
                scan.Invalidate(clearSource: true);
                Assert.Equal(1, scan.PageNumber);
                Assert.False(scan.EditCorners); Assert.False(scan.UsePerspective); Assert.False(scan.CanSelectRegion);
                Assert.True(scan.CanEdit);
                Assert.Null(scan.SourcePreview);
            } finally { window.Close(); }
            return true;
        }, default);
    }

    [Theory]
    [InlineData(960, 640)]
    [InlineData(1280, 800)]
    public async Task DrawRegionPreviewCorrectionsAndSaveReviewedRasterCopy(int width, int height) {
        using var app = TestAppBuilder.StartSession();
        await app.Dispatch(async () => {
            var services = ((App)Application.Current!).Services;
            services.Preferences.Update(value => value with { Theme = width >= 1000 ? StudioThemePreference.Dark : StudioThemePreference.Light });
            string source = Path.Combine(services.Paths.Root, "scan.pdf"), output = Path.Combine(services.Paths.Root, "prepared.pdf");
            PdfDocument.Create(c => c.Page(p => p.Size(240, 320).Content(content => content.Item(i => i.Paragraph(t => t.Text("Scanned page for review")))))).Save(source);
            using var model = new MainWindowViewModel(_ => Task.FromResult<string?>(source), pickSavePdf: _ => Task.FromResult<string?>(output), services: services);
            model.OcrWorkbench.UseDocument(source);
            var scan = model.OcrWorkbench.Scan; scan.Dpi = 72;
            var view = new SearchablePdfOcrView { DataContext = model };
            var window = new Window { Width = width, Height = height, Content = view };
            try {
                window.Show(); window.UpdateLayout();
                await scan.PreviewCommand.ExecuteAsync(null);
                Assert.True(scan.IsCurrent, scan.Status); Assert.NotNull(scan.SourcePreview);
                window.UpdateLayout();
                var canvas = view.GetVisualDescendants().OfType<ScanSelectionCanvas>().Single();
                canvas.BringIntoView(); window.UpdateLayout();
                Capture(window, "scan-before-selection-" + width);
                double scale = Math.Min(canvas.Bounds.Width / scan.SourcePreview.Size.Width, canvas.Bounds.Height / scan.SourcePreview.Size.Height);
                double imageWidth = scan.SourcePreview.Size.Width * scale, imageHeight = scan.SourcePreview.Size.Height * scale;
                Point start = canvas.TranslatePoint(new Point((canvas.Bounds.Width - imageWidth) / 2 + imageWidth * .25, (canvas.Bounds.Height - imageHeight) / 2 + imageHeight * .25), window)!.Value;
                Point end = canvas.TranslatePoint(new Point((canvas.Bounds.Width - imageWidth) / 2 + imageWidth * .75, (canvas.Bounds.Height - imageHeight) / 2 + imageHeight * .75), window)!.Value;
                window.MouseDown(start, MouseButton.Left); window.MouseMove(end, RawInputModifiers.LeftMouseButton); window.MouseUp(end, MouseButton.Left);
                Assert.True(scan.UseRegion, $"Canvas {canvas.Bounds}; start {start}; end {end}; editable {scan.CanEdit}"); Assert.False(scan.IsCurrent); Assert.False(scan.SaveCommand.CanExecute(null));
                Assert.InRange(scan.Region.Width, .1, .9); Assert.InRange(scan.Region.Height, .4, .6);
                Capture(window, "scan-region-" + width);
                Rect drawn = scan.Region;
                canvas.Focus();
                window.KeyPress(Key.Right, RawInputModifiers.None, PhysicalKey.None, null);
                window.KeyPress(Key.Left, RawInputModifiers.Shift, PhysicalKey.None, null);
                Assert.Equal(drawn.X + .01, scan.Region.X, 5);
                Assert.Equal(drawn.Width - .01, scan.Region.Width, 5);
                scan.UsePerspective = true; scan.EditCorners = true;
                window.KeyPress(Key.D1, RawInputModifiers.None, PhysicalKey.None, null);
                window.KeyPress(Key.Right, RawInputModifiers.None, PhysicalKey.None, null);
                window.KeyPress(Key.Down, RawInputModifiers.None, PhysicalKey.None, null);
                Assert.Equal(new Point(.01, .01), scan.TopLeft);
                Point RegionPoint(double x, double y) => canvas.TranslatePoint(new Point(
                    (canvas.Bounds.Width - imageWidth) / 2 + imageWidth * (scan.Region.X + scan.Region.Width * x),
                    (canvas.Bounds.Height - imageHeight) / 2 + imageHeight * (scan.Region.Y + scan.Region.Height * y)), window)!.Value;
                Point cornerStart = RegionPoint(.99, .99), cornerEnd = RegionPoint(.94, .96);
                window.MouseDown(cornerStart, MouseButton.Left);
                window.MouseMove(cornerEnd, RawInputModifiers.LeftMouseButton);
                window.MouseUp(cornerEnd, MouseButton.Left);
                Assert.InRange(scan.BottomRight.X, .93, .95);
                Assert.InRange(scan.BottomRight.Y, .95, .97);
                Capture(window, "scan-perspective-" + width);
                scan.EnableCleanup = true; scan.Deskew = false; scan.NormalizeBackground = false; scan.StraightenDegrees = 3;
                await scan.PreviewCommand.ExecuteAsync(null); Assert.True(scan.IsCurrent, scan.Status);
                var prepared = view.GetVisualDescendants().OfType<Image>().Single(image => ReferenceEquals(image.Source, scan.PreparedPreview));
                prepared.BringIntoView(); window.UpdateLayout(); Capture(window, "scan-prepared-" + width);
                await scan.SaveCommand.ExecuteAsync(null);
                Assert.True(File.Exists(output), model.OcrWorkbench.ErrorMessage ?? scan.Status);
                Assert.Equal(1, PdfDocument.Load(File.ReadAllBytes(output)).Inspect().PageCount);
                scan.Gamma = 1.5; Assert.False(scan.IsCurrent); Assert.False(scan.SaveCommand.CanExecute(null));
            } finally { window.Close(); }
            return true;
        }, default);
    }
    private static void Capture(Window window, string name) {
        using var frame = window.CaptureRenderedFrame(); Assert.NotNull(frame);
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_STUDIO_VISUAL_OUTPUT");
        if (string.IsNullOrWhiteSpace(output)) return; Directory.CreateDirectory(output);
        frame.Save(Path.Combine(output, name + ".png"), Avalonia.Media.Imaging.PngBitmapEncoderOptions.Default);
    }
}
