using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfWriterResourceProgressTests {
    [Theory]
    [InlineData("artifact", 0)]
    [InlineData("artifact", 90)]
    [InlineData("artifact", -90)]
    [InlineData("actual-text", 0)]
    [InlineData("actual-text", 90)]
    [InlineData("actual-text", -90)]
    [InlineData("figure", 0)]
    [InlineData("figure", 90)]
    [InlineData("figure", -90)]
    [InlineData("ordinary", 0)]
    [InlineData("ordinary", 90)]
    [InlineData("ordinary", -90)]
    public void CanvasTableInlineImagesRespectEnclosingAccessibilityOwnership(string group, int rotation) {
        byte[] image = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.Red));
        var runs = new[] { PdfTextRun.Inline(new PdfInlineImage(image, OfficeLength.Points(12), OfficeLength.Points(12), alternativeText: "Child image")) };
        // Imported Office cells use the same native orientation contract.
        var cell = new PdfTableCell(runs, paragraphs: null, textRotation: rotation);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = OfficeLength.Points(200), PageHeight = OfficeLength.Points(200), TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers
        }).Canvas(canvas => {
            void Paint(PdfPageCanvas child) => child.Table(new[] { new[] { cell } }, OfficeLength.Points(20), OfficeLength.Points(20), OfficeLength.Points(100), OfficeLength.Points(60));
            if (group == "actual-text") canvas.ActualText("Logical image", Paint);
            else if (group == "figure") canvas.Figure("Composite diagram", Paint);
            else if (group == "artifact") canvas.Artifact(Paint);
            else Paint(canvas);
        }).ToBytes();

        var raster = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(bytes).Pages[0].ToDrawing());
        Assert.Contains(Enumerable.Range(0, raster.Height), y => Enumerable.Range(0, raster.Width).Any(x => raster.GetPixel(x, y) == OfficeColor.Red));
        var figures = PdfSyntax.ParseObjects(bytes).Map.Values.Select(item => item.Value).OfType<PdfDictionary>()
            .Where(element => element.Get<PdfName>("Type")?.Name == "StructElem" && element.Get<PdfName>("S")?.Name == "Figure").ToArray();
        Assert.Equal(group is "figure" or "ordinary" ? 1 : 0, figures.Length);
    }

    [Theory]
    [InlineData(0.25D)]
    [InlineData(0.5D)]
    [InlineData(0.75D)]
    public void RunningHeadersLeavingSubPointBodyFramesStillMakeSpacerProgress(double bodyHeight) {
        var document = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = OfficeLength.Points(240), PageHeight = OfficeLength.Points(320), MarginTop = OfficeLength.Points(32), MarginBottom = OfficeLength.Points(32),
            MarginLeft = OfficeLength.Points(24), MarginRight = OfficeLength.Points(24), MaxGeneratedPages = 8
        });
        var table = PdfTableStyles.Minimal();
        table.HeaderRowCount = 0;
        table.FixedRowHeights = new List<OfficeLength?> { OfficeLength.Points(270D - bodyHeight) };
        table.ColumnWidths = new List<OfficeLength?> { OfficeLength.Points(100D) };
        document.Header(header => header.Content(content => content.Table(new[] { new[] { "HEADER" } }, style: table), OfficeLength.Points(18)));
        document.Spacer(1);

        byte[] bytes = document.ToBytes();
        var result = PdfReadDocument.Open(bytes);

        Assert.InRange(result.Pages.Count, 1, 4);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Contains(independent.GetPages(), page => page.Text.Contains("HEADER"));
    }

    [Theory]
    [InlineData("artifact", false)]
    [InlineData("artifact", true)]
    [InlineData("actual-text", false)]
    [InlineData("actual-text", true)]
    [InlineData("figure", false)]
    [InlineData("figure", true)]
    public void CanvasGroupsRetainAnActiveImagePageWithoutRegisteringChildStructure(string group, bool textBox) {
        byte[] image = OfficePngWriter.Encode(new OfficeRasterImage(2, 2, OfficeColor.Red));
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = OfficeLength.Points(200), PageHeight = OfficeLength.Points(200),
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers
        }).Canvas(canvas => {
            void Paint(PdfPageCanvas child) {
                var runs = new[] { PdfTextRun.Inline(new PdfInlineImage(image, OfficeLength.Points(12), OfficeLength.Points(12), alternativeText: "Child image")) };
                if (textBox) child.TextBox(runs, OfficeLength.Points(20), OfficeLength.Points(20), OfficeLength.Points(80), OfficeLength.Points(40));
                else child.Text(runs, OfficeLength.Points(20), OfficeLength.Points(20), OfficeLength.Points(80), OfficeLength.Points(40));
            }
            if (group == "actual-text") canvas.ActualText("Logical image", Paint);
            else if (group == "figure") canvas.Figure("Composite diagram", Paint);
            else canvas.Artifact(Paint);
        }).ToBytes();

        var page = PdfReadDocument.Open(bytes).Pages[0];
        var raster = OfficeDrawingRasterRenderer.Render(page.ToDrawing());
        Assert.Contains(Enumerable.Range(0, raster.Height), y => Enumerable.Range(0, raster.Width).Any(x => raster.GetPixel(x, y) == OfficeColor.Red));
        var figures = PdfSyntax.ParseObjects(bytes).Map.Values.Select(item => item.Value).OfType<PdfDictionary>()
            .Where(element => element.Get<PdfName>("Type")?.Name == "StructElem" && element.Get<PdfName>("S")?.Name == "Figure").ToArray();
        Assert.Equal(group == "figure" ? 1 : 0, figures.Length);
    }
}
