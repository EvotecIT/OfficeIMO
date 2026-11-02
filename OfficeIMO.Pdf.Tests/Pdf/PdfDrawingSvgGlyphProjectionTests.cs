using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingSvgGlyphProjectionTests {
    [Theory]
    [InlineData("OfficeIMO.Drawing.Tests/TestAssets/SourceSansPro-Regular.otf", "Marks stay attached")]
    [InlineData("OfficeIMO.Drawing.Tests/TestAssets/SourceSansPro-Regular.otf", "Case B-76: accepted.")]
    [InlineData("OfficeIMO.Drawing.Tests/TestAssets/OpenSans-Bold.ttf", " ")]
    public void SavedPdfGlyphsExportAsSvgContoursWithoutBrowserFontSubstitution(string relativeFont, string text) {
        byte[] fontData = File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), relativeFont));
        var options = new PdfCreateOptions().UseManagedTextShaping()
            .EmbedStandardFont(PdfStandardFont.Helvetica, fontData, "Svg Projection");
        byte[] bytes = PdfDocument.Create(options).Canvas(canvas => canvas.PositionedText(
            new[] { new PdfTextRun(text, fontSize: global::OfficeIMO.OfficeLength.Points(20D)) }, PdfCanvasTextStructureRole.Paragraph,
            20D, 20D, 260D, 40D, fontSize: 20D)).ToBytes();
        OfficeDrawing drawing = PdfReadDocument.Open(bytes).Pages[0].ToDrawing();
        OfficeDrawingText[] runs = drawing.Elements.OfType<OfficeDrawingText>().ToArray();

        Assert.NotEmpty(runs);
        Assert.Equal(text, string.Concat(runs.Select(run => run.Text)));
        drawing.TextShapingProvider = new RejectReshaping();
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing);
        XElement root = XElement.Parse(svg);
        Assert.Equal(text, string.Concat(root.Descendants().Attributes("aria-label").Select(attribute => attribute.Value)));
        Assert.DoesNotContain(root.Descendants(), element => element.Name.LocalName is "text" or "tspan");
        Assert.DoesNotContain("@font-face", svg, StringComparison.Ordinal);
        string[] paths = root.Descendants().Where(element => element.Name.LocalName == "path")
            .Select(element => element.Attribute("d")?.Value ?? string.Empty).ToArray();
        Assert.NotEmpty(paths);
        Assert.Equal(!string.IsNullOrWhiteSpace(text), paths.Any(path => path.Contains('M')));

        Assert.NotEmpty(OfficeDrawingRasterRenderer.Render(drawing).GetPixels());
        var replayOptions = new PdfCreateOptions {
            PageWidth = global::OfficeIMO.OfficeLength.Points(drawing.Width),
            PageHeight = global::OfficeIMO.OfficeLength.Points(drawing.Height),
            MarginLeft = global::OfficeIMO.OfficeLength.Points(0D),
            MarginRight = global::OfficeIMO.OfficeLength.Points(0D),
            MarginTop = global::OfficeIMO.OfficeLength.Points(0D),
            MarginBottom = global::OfficeIMO.OfficeLength.Points(0D)
        }.SetTextShapingProvider(new RejectReshaping());
        byte[] replay = PdfDocument.Create(replayOptions)
            .Drawing(drawing).ToBytes();
        Assert.Equal(text.Trim(), PdfReadDocument.Open(replay).ExtractText().Trim());
    }

    private sealed class RejectReshaping : IOfficeTextShapingProvider {
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) =>
            throw new InvalidOperationException("Resolved PDF glyphs must keep their painted forms.");
    }
}
