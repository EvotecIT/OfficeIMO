using System.Xml.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfOutputIntentCompositionTests {
    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory,
        "Pdf", "Fixtures", "Rendering", "OutputIntent", name);

    [Theory]
    [InlineData(2, "Opacity proof")]
    [InlineData(3, "Multiply proof")]
    [InlineData(4, "Image mask proof")]
    public void ProducerPages_ProofCompositedPixelsAndKeepTextGeometry(int pageNumber, string label) {
        byte[] pdf = File.ReadAllBytes(Fixture("reportlab-output-intent.pdf"));
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[pageNumber - 1];
        var before = page.GetTextSpans().Select(span => (span.Text, span.X, span.Y, span.Advance)).ToArray();
        Assert.Equal(label, string.Concat(before.Select(span => span.Text)));
        Assert.Equal(6D, before[0].X, 4);
        Assert.Equal(82D, before[0].Y, 4);

        OfficeDrawing drawing = page.ToDrawing();
        Assert.True(OfficePngReader.TryDecode(File.ReadAllBytes(Fixture($"page-{pageNumber}-reference.png")), out OfficeRasterImage? reference));
        Assert.NotNull(reference);
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);
        Assert.Equal(128, raster.Width);
        Assert.Equal(96, raster.Height);
        foreach (int x in new[] { 20, 56, 99 }) AssertColor(reference!.GetPixel(x, 56), raster.GetPixel(x, 56));

        PdfRenderCapabilityDiagnostic diagnostic = Assert.Single(page.GetRenderCapabilityDiagnostics(),
            item => item.Code == "render.colorspace.icc-output-intent-post-composition");
        Assert.Equal(PdfRenderSupportLevel.Supported, diagnostic.SupportLevel);
        Assert.Equal(OfficeConversionLossKind.None, diagnostic.LossKind);
        Assert.DoesNotContain(page.GetRenderCapabilityDiagnostics(), item => item.Code == "render.colorspace.icc-output-intent-transparency-simplified");
        Assert.Equal(before, page.GetTextSpans().Select(span => (span.Text, span.X, span.Y, span.Advance)).ToArray());
        Assert.Equal(raster.GetPixels(), OfficeDrawingRasterRenderer.Render(drawing.Clone(), background: OfficeColor.White).GetPixels());

        PdfPageRenderResult svgResult = Assert.Single(PdfDocument.Load(Fixture("reportlab-output-intent.pdf"))
            .Render.Pages(PdfPageSelection.From(pageNumber), new PdfPageRenderOptions {
                Format = OfficeImageExportFormat.Svg, Scale = 2D
            }));
        Assert.True(svgResult.Succeeded, string.Join(" ", svgResult.Diagnostics));
        string svg = Encoding.UTF8.GetString(svgResult.Bytes!);
        XElement image = Assert.Single(XDocument.Parse(svg).Descendants(), item => item.Name.LocalName == "image");
        byte[] png = Convert.FromBase64String(image.Attribute("href")!.Value.Split(',')[1]);
        Assert.True(OfficePngReader.TryDecode(png, out OfficeRasterImage? svgRaster));
        Assert.Equal(256, svgRaster!.Width);
        Assert.Equal(192, svgRaster.Height);
        foreach (int x in new[] { 20, 56, 99 }) AssertColor(reference!.GetPixel(x, 56), svgRaster.GetPixel(x * 2, 112));
    }

    [Fact]
    public void ProducerBatch_ReportsOnlyPagesThatRequirePostComposition() {
        byte[] pdf = File.ReadAllBytes(Fixture("reportlab-output-intent.pdf"));
        IReadOnlyList<PdfPageRenderResult> results = PdfPageImageRenderer.RenderPages(pdf);
        Assert.Equal(4, results.Count);
        Assert.All(results, result => Assert.True(result.Succeeded, string.Join(" ", result.Diagnostics)));
        Assert.DoesNotContain(results[0].CapabilityDiagnostics, item => item.Code.Contains("post-composition"));
        Assert.All(results.Skip(1), result => Assert.Contains(result.CapabilityDiagnostics,
            item => item.Code == "render.colorspace.icc-output-intent-post-composition"));
    }

    [Fact]
    public void SvgExport_PreservesRequestedTransparentBackdrop() {
        PdfReadPage page = PdfReadDocument.Open(File.ReadAllBytes(Fixture("reportlab-output-intent.pdf"))).Pages[1];
        OfficeImageExportResult result = page.ToImage(OfficeImageExportFormat.Svg,
            new PdfImageExportOptions { BackgroundColor = OfficeColor.Transparent });
        XElement image = Assert.Single(XDocument.Parse(Encoding.UTF8.GetString(result.Bytes)).Descendants(),
            item => item.Name.LocalName == "image");
        Assert.True(OfficePngReader.TryDecode(Convert.FromBase64String(image.Attribute("href")!.Value.Split(',')[1]),
            out OfficeRasterImage? raster));
        Assert.Equal((byte)0, raster!.GetPixel(0, 0).A);
        Assert.InRange(raster.GetPixel(99, 56).A, (byte)127, (byte)129);
    }

    [Fact]
    public void SvgExport_RejectsProofSurfaceBeyondTheCallerRasterBudget() {
        PdfReadPage page = PdfReadDocument.Open(File.ReadAllBytes(Fixture("reportlab-output-intent.pdf"))).Pages[1];
        OfficeImageExportLimitException error = Assert.Throws<OfficeImageExportLimitException>(() =>
            page.ToImage(OfficeImageExportFormat.Svg, new PdfImageExportOptions {
                MaximumRasterPixels = 100, RasterOverflowBehavior = OfficeRasterOverflowBehavior.Throw
            }));
        Assert.Equal(100, error.MaximumPixels);
        Assert.Equal(128 * 96, error.RequestedPixels);
    }

    [Theory]
    [InlineData(1D)]
    [InlineData(1.7D)]
    public void ProducerTiles_MatchWholeProofWithoutAllocatingTheWholePage(double scale) {
        byte[] pdf = File.ReadAllBytes(Fixture("reportlab-output-intent.pdf"));
        OfficeDrawing drawing = PdfReadDocument.Open(pdf).Pages[3].ToDrawing();
        OfficeRasterImage whole = OfficeDrawingRasterRenderer.Render(drawing, scale, OfficeColor.White);
        var renderer = new PdfPageDisplayRenderer(PdfDocument.Load(pdf), 1, 1_000_000);
        var limits = new PdfPageRegionLimits { MaximumRegionPixels = 16, MaximumWorkingPixels = 1024 };
        OfficeRasterImage tile = renderer.RenderRegion(4, scale, 54, 54, 4, 4, limits, default);
        for (int y = 0; y < 4; y++)
            for (int x = 0; x < 4; x++) Assert.Equal(whole.GetPixel(54 + x, 54 + y), tile.GetPixel(x, y));
    }

    [Theory]
    [InlineData("drawing")]
    [InlineData("clip")]
    [InlineData("effect")]
    [InlineData("pattern")]
    public void NestedProducerPage_KeepsItsOwnProofedSurface(string composition) {
        byte[] pdf = File.ReadAllBytes(Fixture("reportlab-output-intent.pdf"));
        OfficeDrawing page = PdfReadDocument.Open(pdf).Pages[1].ToDrawing();
        var drawing = new OfficeDrawing(150, 120);
        switch (composition) {
            case "drawing": drawing.AddDrawing(page, 5, 7); break;
            case "clip": drawing.AddClippedDrawing(page, 5, 7, OfficeClipPath.Rectangle(128, 96)); break;
            case "effect": drawing.AddEffectDrawing(page, OfficeTransform.Translate(5, 7)); break;
            case "pattern":
                drawing.AddTilingPattern(page, new OfficeImagePlacement(5, 7, 128, 96), 128, 96, originX: 5, originY: 7);
                break;
        }
        OfficeRasterImage expected = OfficeDrawingRasterRenderer.Render(page);
        OfficeRasterImage actual = OfficeDrawingRasterRenderer.Render(drawing);
        foreach (int x in new[] { 20, 56, 99 }) Assert.Equal(expected.GetPixel(x, 56), actual.GetPixel(x + 5, 63));
        Assert.Equal(expected.GetPixel(5, 70).A, actual.GetPixel(10, 77).A);
        Assert.Equal((byte)0, actual.GetPixel(0, 0).A);
        XElement image = Assert.Single(XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing)).Descendants(),
            item => item.Name.LocalName == "image");
        Assert.True(OfficePngReader.TryDecode(Convert.FromBase64String(image.Attribute("href")!.Value.Split(',')[1]),
            out OfficeRasterImage? svgRaster));
        Assert.Equal(expected.GetPixels(), svgRaster!.GetPixels());
    }

    private static void AssertColor(OfficeColor expected, OfficeColor actual) {
        Assert.InRange(Math.Abs(actual.R - expected.R), 0, 4);
        Assert.InRange(Math.Abs(actual.G - expected.G), 0, 4);
        Assert.InRange(Math.Abs(actual.B - expected.B), 0, 4);
        Assert.Equal((byte)255, actual.A);
    }
}
