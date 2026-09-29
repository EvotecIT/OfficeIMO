using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class ExcelChartRadialLayoutTests {
    [Theory]
    [InlineData("361", "50")]
    [InlineData("0", "9")]
    [InlineData("invalid", "50")]
    [InlineData("9999999999999", "50")]
    public void RadialGeometry_MalformedNativeValuesFailTrySnapshot(string angle, string hole) {
        using ExcelDocument document = ExcelDocument.Create();
        var chart = document.AddWorksheet("Results").AddChart(OfficeChartKind.Doughnut,
            new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Status", new[] { 1d }) }), 1, 1);
        var part = document.OpenXmlDocument.WorkbookPart!.WorksheetParts.Single(p => p.DrawingsPart != null).DrawingsPart!.ChartParts.Single();
        chart.SetRadialLayout(OfficeChartRadialLayout.Default);
        var native = part.ChartSpace.Descendants<C.DoughnutChart>().Single();
        native.GetFirstChild<C.FirstSliceAngle>()!.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", "val", "", angle));
        native.GetFirstChild<C.HoleSize>()!.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("", "val", "", hole));
        Assert.False(chart.TryGetSnapshot(out _));
    }

    [Fact]
    public void RadialGeometry_WorkbookPdfPreservesHoleSize() {
        int ColoredPixels(int hole) {
            using ExcelDocument document = ExcelDocument.Create();
            document.AddWorksheet("Results").AddChart(OfficeChartKind.Doughnut,
                new OfficeChartData(new[] { "A" }, new[] { new OfficeChartSeries("Status", new[] { 1d }, null, color: OfficeColor.Parse("#FF0000")) }), 1, 1)
                .SetRadialLayout(new OfficeChartRadialLayout(doughnutHolePercent: hole));
            byte[] pdf = document.ToPdfBytes(new ExcelToPdfOptions { IncludeSheetHeadings = false });
            var pages = OfficeIMO.Pdf.PdfDocument.Load(pdf).Render.Pages(options:
                new OfficeIMO.Pdf.PdfPageRenderOptions { Dpi = 72, MaxPages = 8, ContinueOnError = false });
            int count = 0;
            foreach (var page in pages) {
                Assert.True(OfficePngReader.TryDecode(page.Bytes!, out OfficeRasterImage? raster));
                for (int y = 0; y < raster!.Height; y++)
                    for (int x = 0; x < raster.Width; x++) {
                        OfficeColor pixel = raster.GetPixel(x, y);
                        if (pixel.R > 200 && pixel.G < 80 && pixel.B < 80) count++;
                    }
            }
            return count;
        }
        int thin = ColoredPixels(75), thick = ColoredPixels(25);
        Assert.True(thin > 100);
        Assert.InRange((double)thin / thick, 0.35, 0.6);
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void RadialGeometry_PersistsThroughReopenUpdateAndSnapshot(OfficeChartKind kind) {
        var data = new OfficeChartData(new[] { "A", "B" }, new[] { new OfficeChartSeries("Status", new[] { 7d, 3d }) });
        using ExcelDocument authored = ExcelDocument.Create();
        authored.AddWorksheet("Results").AddChart(kind, data, 1, 1).SetRadialLayout(new OfficeChartRadialLayout(135, 75));
        using var bytes = new MemoryStream(authored.ToBytes());
        using ExcelDocument reopened = ExcelDocument.Load(bytes);
        ExcelChart chart = reopened.Sheets.Single(sheet => sheet.Name == "Results").Charts.Single();
        chart.UpdateData(new ExcelChartData(new[] { "A", "B" }, new[] { new ExcelChartSeries("Status", new[] { 8d, 2d }) }));
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Equal(135, snapshot.RadialLayout.FirstSliceAngleDegrees);
        Assert.Equal(kind == OfficeChartKind.Doughnut ? 75 : 50, snapshot.RadialLayout.DoughnutHolePercent);
        Assert.Empty(reopened.ValidateDocument());
    }
}
