using System.IO;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class ExcelChartExplodedSlicesTests {
    [Fact]
    public void PieWithoutThemePartUsesSpreadsheetDefaultAccentPalette() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddWorksheet("Results").AddChart(OfficeChartKind.Pie,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Status", new[] { 7d, 3d })
            }), 1, 1);
        WorkbookPart workbook = document.OpenXmlDocument.WorkbookPart!;
        if (workbook.ThemePart is ThemePart theme) workbook.DeletePart(theme);

        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Equal(OfficeColor.Parse("#4F81BD"), snapshot.Style!.Palette[0]);
        Assert.Equal(OfficeColor.Parse("#C0504D"), snapshot.Style.Palette[1]);
    }

    [Fact]
    public void CustomModernColorStyleControlsRadialPaletteWithoutAuthoringPointFills() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddWorksheet("Results").AddChart(OfficeChartKind.Pie,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Status", new[] { 7d, 3d })
            }), 1, 1);
        chart.ApplyStylePreset();
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot baseline));
        ChartPart chartPart = document.OpenXmlDocument.WorkbookPart!.WorksheetParts
            .Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single();
        string styleXml;
        using (var reader = new StreamReader(Assert.Single(chartPart.GetPartsOfType<ChartStylePart>()).GetStream()))
            styleXml = reader.ReadToEnd();
        string colorXml;
        using (var reader = new StreamReader(Assert.Single(chartPart.GetPartsOfType<ChartColorStylePart>()).GetStream()))
            colorXml = reader.ReadToEnd();
        string customColors = colorXml.Replace(
            "<a:schemeClr val=\"accent1\"/><a:schemeClr val=\"accent2\"/>",
            "<a:schemeClr val=\"accent2\"/><a:schemeClr val=\"accent1\"/>");
        Assert.NotEqual(colorXml, customColors);
        chart.ApplyStylePreset(new ExcelChartStylePreset(styleXml, customColors));

        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Equal(baseline.Style!.Palette[1], snapshot.Style!.Palette[0]);
        Assert.Equal(baseline.Style.Palette[0], snapshot.Style.Palette[1]);
        Assert.Null(snapshot.Data.Series.Single().PointColorArgb);
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void MultiRingDoughnutRejectsTransformedModernPalette() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddWorksheet("Results").AddChart(OfficeChartKind.Doughnut,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Inner", new[] { 7d, 3d }),
                new OfficeChartSeries("Outer", new[] { 2d, 8d })
            }), 1, 1);
        chart.ApplyStylePreset();
        Assert.False(chart.TryGetSnapshot(out _));
    }

    [Fact]
    public void MultiRingDoughnutWithExplicitSliceFillsDoesNotNeedTheTransformedPalette() {
        OfficeColor?[] colors = { OfficeColor.Parse("#234567"), OfficeColor.Parse("#89ABCD") };
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddWorksheet("Results").AddChart(OfficeChartKind.Doughnut,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Inner", new[] { 7d, 3d }, null, null, colors),
                new OfficeChartSeries("Outer", new[] { 2d, 8d }, null, null, colors)
            }), 1, 1);
        chart.ApplyStylePreset();
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.All(snapshot.Data.Series, series =>
            Assert.Equal(new[] { "234567", "89ABCD" }, series.PointColorArgb));
    }

    [Fact]
    public void ModernPaletteRejectsNonAutomaticDataPointFill() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddWorksheet("Results").AddChart(OfficeChartKind.Pie,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Status", new[] { 7d, 3d })
            }), 1, 1);
        chart.ApplyStylePreset();
        ChartPart chartPart = document.OpenXmlDocument.WorkbookPart!.WorksheetParts
            .Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single();
        string styleXml;
        using (var reader = new StreamReader(Assert.Single(chartPart.GetPartsOfType<ChartStylePart>()).GetStream()))
            styleXml = reader.ReadToEnd();
        string colorXml;
        using (var reader = new StreamReader(Assert.Single(chartPart.GetPartsOfType<ChartColorStylePart>()).GetStream()))
            colorXml = reader.ReadToEnd();
        string customStyle = styleXml.Replace("<cs:fillRef idx=\"1\"><cs:styleClr val=\"auto\"/></cs:fillRef>",
            "<cs:fillRef idx=\"1\"><cs:styleClr val=\"accent1\"/></cs:fillRef>");
        Assert.NotEqual(styleXml, customStyle);
        chart.ApplyStylePreset(new ExcelChartStylePreset(customStyle, colorXml));
        Assert.False(chart.TryGetSnapshot(out _));
    }

    [Fact]
    public void ExcelProducedPieAndDoughnut_ProjectExplodedPoint() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "Charts", "Excel",
            "exploded-slices.xlsx");
        using ExcelDocument document = ExcelDocument.Load(path);
        ExcelChart[] charts = document.Sheets.Single(sheet => sheet.Name == "Exploded")
            .Charts.ToArray();
        Assert.Equal(2, charts.Length);
        foreach (ExcelChart chart in charts) {
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal(new[] { 25, 0 }, snapshot.Data.Series.Single().PointExplosions);
            Assert.Equal(new[] { "156082", "E97132" }, snapshot.Style!.Palette.Take(2)
                .Select(color => color.ToRgbHex()));
            Assert.Null(snapshot.Data.Series.Single().PointColorArgb);
            Assert.NotEmpty(chart.ExportImage(OfficeImageExportFormat.Svg).Bytes);
            MethodInfo pdfProjection = typeof(ExcelPdfConverterExtensions).GetMethod(
                "CreateOfficeChartSnapshot", BindingFlags.NonPublic | BindingFlags.Static)!;
            OfficeChartSnapshot pdfSnapshot = Assert.IsType<OfficeChartSnapshot>(
                pdfProjection.Invoke(null, new object[] { snapshot, new ExcelToPdfOptions() }));
            Assert.Equal(new[] { 25, 0 }, pdfSnapshot.Data.Series.Single().PointExplosions);
        }
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void ExplodedPoint_RoundTripsIntoStaticExportAndSurvivesValueUpdate(OfficeChartKind kind) {
        var authored = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeChartSeries("Status", new[] { 7d, 3d }).WithPointExplosions(new[] { 25, 0 })
        });
        using var document = ExcelDocument.Create();
        document.AddWorksheet("Results").AddChart(kind, authored, 1, 1);
        using var bytes = new MemoryStream(document.ToBytes());
        using ExcelDocument reopened = ExcelDocument.Load(bytes);
        ExcelChart chart = reopened.Sheets.Single(sheet => sheet.Name == "Results").Charts.Single();
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Equal(new[] { 25, 0 }, snapshot.Data.Series.Single().PointExplosions);
        Assert.NotEmpty(chart.ExportImage(OfficeImageExportFormat.Svg).Bytes);
        chart.UpdateData(new ExcelChartData(new[] { "A", "B" }, new[] {
            new ExcelChartSeries("Status", new[] { 8d, 2d })
        }));
        C.PieChartSeries native = reopened.OpenXmlDocument.WorkbookPart!.WorksheetParts
            .Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single()
            .ChartSpace!.Descendants<C.PieChartSeries>().Single();
        Assert.Equal((uint)25, native.Elements<C.DataPoint>()
            .Single(point => point.Index!.Val!.Value == 0).GetFirstChild<C.Explosion>()!.Val!.Value);
        Assert.Empty(reopened.ValidateDocument());
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void DataPointExplosion_CanBeChangedAndResetWithoutRecreatingTheChart(OfficeChartKind kind) {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddWorksheet("Results").AddChart(kind,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Status", new[] { 7d, 3d })
                    .WithPointExplosions(new[] { 25, 0 })
            }), 1, 1);
        chart.SetDataPointExplosion(0, 1, 30);
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot changed));
        Assert.Equal(new[] { 25, 30 }, changed.Data.Series.Single().PointExplosions);
        chart.SetDataPointExplosion(0, 1, 0);
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot reset));
        Assert.Equal(new[] { 25, 0 }, reset.Data.Series.Single().PointExplosions);
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void DataPointExplosion_ExplicitZeroOverridesInheritedSeriesOffset() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddWorksheet("Results").AddChart(OfficeChartKind.Pie,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Status", new[] { 7d, 3d })
            }), 1, 1);
        C.PieChartSeries native = document.OpenXmlDocument.WorkbookPart!.WorksheetParts
            .Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single()
            .ChartSpace!.Descendants<C.PieChartSeries>().Single();
        native.AddChild(new C.Explosion { Val = 10U }, true);
        chart.SetDataPointExplosion(0, 0, 0);
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot reset));
        Assert.Equal(new[] { 0, 10 }, reset.Data.Series.Single().PointExplosions);
        chart.SetDataPointExplosion(0, 0, null);
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot inherited));
        Assert.Equal(new[] { 10, 10 }, inherited.Data.Series.Single().PointExplosions);
    }

    [Fact]
    public void UnsupportedNativeExplosion_RejectsStaticSnapshotWithoutChangingNativeXml() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddWorksheet("Results").AddChart(OfficeChartKind.Pie,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Status", new[] { 7d, 3d })
            }), 1, 1);
        C.PieChartSeries native = document.OpenXmlDocument.WorkbookPart!.WorksheetParts
            .Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single()
            .ChartSpace!.Descendants<C.PieChartSeries>().Single();
        native.AddChild(new C.Explosion { Val = 401U }, true);
        Assert.False(chart.TryGetSnapshot(out _));
        Assert.Equal((uint)401, native.GetFirstChild<C.Explosion>()!.Val!.Value);
    }

    [Theory]
    [InlineData(OfficeChartKind.Pie)]
    [InlineData(OfficeChartKind.Doughnut)]
    public void ShrinkingValues_IgnoresPreservedExplosionForRemovedPoint(OfficeChartKind kind) {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddWorksheet("Results").AddChart(kind,
            new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeChartSeries("Status", new[] { 7d, 3d }).WithPointExplosions(new[] { 0, 25 })
            }), 1, 1);
        chart.UpdateData(new ExcelChartData(new[] { "A" }, new[] {
            new ExcelChartSeries("Status", new[] { 8d })
        }));
        Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
        Assert.Null(snapshot.Data.Series.Single().PointExplosions);
        Assert.NotEmpty(chart.ExportImage(OfficeImageExportFormat.Svg).Bytes);
        Assert.Empty(document.ValidateDocument());
    }
}
