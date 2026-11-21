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
        ExcelChart chart = document.AddSheet("Results").AddChart(OfficeChartType.Pie, new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeSeries("Status", new[] { 7d, 3d })
            }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
        WorkbookPart workbook = document.OpenXmlDocument.WorkbookPart!;
        if (workbook.ThemePart is ThemePart theme) workbook.DeletePart(theme);

        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));
        Assert.Equal(OfficeColor.Parse("#4F81BD"), snapshot.Style!.Palette[0]);
        Assert.Equal(OfficeColor.Parse("#C0504D"), snapshot.Style.Palette[1]);
    }

    [Fact]
    public void MultiRingDoughnutWithExplicitSliceFillsDoesNotNeedTheTransformedPalette() {
        OfficeColor?[] colors = { OfficeColor.Parse("#234567"), OfficeColor.Parse("#89ABCD") };
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddSheet("Results").AddChart(OfficeChartType.Doughnut, new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeSeries("Inner", new[] { 7d, 3d }, null, null, colors),
                new OfficeSeries("Outer", new[] { 2d, 8d }, null, null, colors)
            }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
        chart.ApplyStylePreset();
        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));
        Assert.All(snapshot.Data.Series, series =>
            Assert.Equal(new[] { "234567", "89ABCD" }, series.PointColorArgb));
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
            Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));
            Assert.Equal(new[] { 25, 0 }, snapshot.Data.Series.Single().PointExplosions);
            Assert.Equal(new[] { "156082", "E97132" }, snapshot.Style!.Palette.Take(2)
                .Select(color => color.ToRgbHex()));
            Assert.Null(snapshot.Data.Series.Single().PointColorArgb);
            Assert.NotEmpty(chart.ToImage(OfficeImageExportFormat.Svg).Bytes);
            MethodInfo pdfProjection = typeof(ExcelPdfConverterExtensions).GetMethod(
                "CreateOfficeChartSnapshot", BindingFlags.NonPublic | BindingFlags.Static)!;
            OfficeChartSnapshot pdfSnapshot = Assert.IsType<OfficeChartSnapshot>(
                pdfProjection.Invoke(null, new object[] { snapshot, new ExcelPdfExportOptions() }));
            Assert.Equal(new[] { 25, 0 }, pdfSnapshot.Data.Series.Single().PointExplosions);
        }
    }

    [Theory]
    [InlineData(OfficeChartType.Pie)]
    [InlineData(OfficeChartType.Doughnut)]
    public void ExplodedPoint_RoundTripsIntoStaticExportAndSurvivesValueUpdate(OfficeChartType kind) {
        var authored = new OfficeChartData(new[] { "A", "B" }, new[] {
            new OfficeSeries("Status", new[] { 7d, 3d }).WithPointExplosions(new[] { 25, 0 })
        });
        using var document = ExcelDocument.Create();
        document.AddSheet("Results").AddChart(kind, authored, new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
        using var bytes = new MemoryStream(document.ToBytes());
        using ExcelDocument reopened = ExcelDocument.Load(bytes);
        ExcelChart chart = reopened.Sheets.Single(sheet => sheet.Name == "Results").Charts.Single();
        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));
        Assert.Equal(new[] { 25, 0 }, snapshot.Data.Series.Single().PointExplosions);
        Assert.NotEmpty(chart.ToImage(OfficeImageExportFormat.Svg).Bytes);
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
    [InlineData(OfficeChartType.Pie)]
    [InlineData(OfficeChartType.Doughnut)]
    public void DataPointExplosion_CanBeChangedAndResetWithoutRecreatingTheChart(OfficeChartType kind) {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddSheet("Results").AddChart(kind, new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeSeries("Status", new[] { 7d, 3d })
                    .WithPointExplosions(new[] { 25, 0 })
            }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
        chart.SetDataPointExplosion(0, 1, 30);
        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo changed));
        Assert.Equal(new[] { 25, 30 }, changed.Data.Series.Single().PointExplosions);
        chart.SetDataPointExplosion(0, 1, 0);
        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo reset));
        Assert.Equal(new[] { 25, 0 }, reset.Data.Series.Single().PointExplosions);
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void DataPointExplosion_ExplicitZeroOverridesInheritedSeriesOffset() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddSheet("Results").AddChart(OfficeChartType.Pie, new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeSeries("Status", new[] { 7d, 3d })
            }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
        C.PieChartSeries native = document.OpenXmlDocument.WorkbookPart!.WorksheetParts
            .Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single()
            .ChartSpace!.Descendants<C.PieChartSeries>().Single();
        native.AddChild(new C.Explosion { Val = 10U }, true);
        chart.SetDataPointExplosion(0, 0, 0);
        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo reset));
        Assert.Equal(new[] { 0, 10 }, reset.Data.Series.Single().PointExplosions);
        chart.SetDataPointExplosion(0, 0, null);
        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo inherited));
        Assert.Equal(new[] { 10, 10 }, inherited.Data.Series.Single().PointExplosions);
    }

    [Fact]
    public void UnsupportedNativeExplosion_RejectsStaticSnapshotWithoutChangingNativeXml() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddSheet("Results").AddChart(OfficeChartType.Pie, new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeSeries("Status", new[] { 7d, 3d })
            }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
        C.PieChartSeries native = document.OpenXmlDocument.WorkbookPart!.WorksheetParts
            .Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single()
            .ChartSpace!.Descendants<C.PieChartSeries>().Single();
        native.AddChild(new C.Explosion { Val = 401U }, true);
        Assert.False(chart.TryGetSnapshot(out _));
        Assert.Equal((uint)401, native.GetFirstChild<C.Explosion>()!.Val!.Value);
    }

    [Fact]
    public void UnsupportedNativeExplosion_BestEffortExportDrawsTheChartAndReportsTheApproximation() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddSheet("Results").AddChart(OfficeChartType.Pie, new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeSeries("Status", new[] { 7d, 3d })
            }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
        document.OpenXmlDocument.WorkbookPart!.WorksheetParts
            .Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single()
            .ChartSpace!.Descendants<C.PieChartSeries>().Single().AddChild(new C.Explosion { Val = 401U }, true);

        NotSupportedException strict = Assert.Throws<NotSupportedException>(() => chart.ToImage(OfficeImageExportFormat.Png));
        Assert.Contains("point explosion", strict.Message, StringComparison.Ordinal);
        OfficeImageExportResult image = chart.ToImage(OfficeImageExportFormat.Png,
            new ExcelImageExportOptions { ChartRendering = { Fidelity = OfficeChartFidelity.BestEffort } });
        OfficeIMO.Pdf.PdfDocumentConversionResult pdf = document.ConvertToPdfWithEvidence(
            new ExcelPdfExportOptions { ChartRendering = { Fidelity = OfficeChartFidelity.BestEffort } });

        Assert.NotEmpty(image.Bytes);
        Assert.Contains(image.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ChartApproximation &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.DoesNotContain(pdf.Warnings, warning => warning.Code == "worksheet-chart");
        Assert.Contains(pdf.Warnings, warning => warning.Code == OfficeDiagnosticCodes.ChartApproximation);
    }

    [Fact]
    public void TrendlineDashTheRendererCannotDraw_StrictSkipsAndBestEffortReportsIt() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddSheet("Results").AddChart(OfficeChartType.Line, new OfficeChartData(new[] { "A", "B", "C" }, new[] {
                new OfficeSeries("Status", new[] { 7d, 3d, 5d })
            }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
        C.LineChartSeries series = document.OpenXmlDocument.WorkbookPart!.WorksheetParts
            .Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single()
            .ChartSpace!.Descendants<C.LineChartSeries>().Single();
        series.InsertBefore(new C.Trendline(
                new C.ChartShapeProperties(new DocumentFormat.OpenXml.Drawing.Outline(
                    new DocumentFormat.OpenXml.Drawing.PresetDash { Val = DocumentFormat.OpenXml.Drawing.PresetLineDashValues.SystemDot })),
                new C.TrendlineType { Val = C.TrendlineValues.Linear }),
            series.GetFirstChild<C.CategoryAxisData>());

        Assert.True(chart.TryGetData(out _));
        NotSupportedException strict = Assert.Throws<NotSupportedException>(() => chart.ToImage(OfficeImageExportFormat.Png));
        Assert.Contains("The trendline dash pattern cannot be projected.", strict.Message, StringComparison.Ordinal);
        OfficeImageExportResult image = chart.ToImage(OfficeImageExportFormat.Png,
            new ExcelImageExportOptions { ChartRendering = { Fidelity = OfficeChartFidelity.BestEffort } });

        Assert.Contains(image.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ChartApproximation &&
            diagnostic.Message.IndexOf("'sysDot' dash pattern", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void HiddenTrendlineWithADashIsDrawnInStrict() {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddSheet("Results").AddChart(OfficeChartType.Line, new OfficeChartData(new[] { "A", "B", "C" }, new[] {
                new OfficeSeries("Status", new[] { 7d, 3d, 5d })
            }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
        C.LineChartSeries series = document.OpenXmlDocument.WorkbookPart!.WorksheetParts
            .Single(part => part.DrawingsPart != null).DrawingsPart!.ChartParts.Single()
            .ChartSpace!.Descendants<C.LineChartSeries>().Single();
        series.InsertBefore(new C.Trendline(
                new C.ChartShapeProperties(new DocumentFormat.OpenXml.Drawing.Outline(new DocumentFormat.OpenXml.Drawing.NoFill(),
                    new DocumentFormat.OpenXml.Drawing.PresetDash { Val = DocumentFormat.OpenXml.Drawing.PresetLineDashValues.SystemDot })),
                new C.TrendlineType { Val = C.TrendlineValues.Linear }, new C.DisplayEquation { Val = true }),
            series.GetFirstChild<C.CategoryAxisData>());

        OfficeImageExportResult image = chart.ToImage(OfficeImageExportFormat.Png);

        Assert.DoesNotContain(image.Diagnostics, diagnostic => diagnostic.Code == OfficeDiagnosticCodes.ChartApproximation);
    }

    [Theory]
    [InlineData(OfficeChartType.Pie)]
    [InlineData(OfficeChartType.Doughnut)]
    public void ShrinkingValues_IgnoresPreservedExplosionForRemovedPoint(OfficeChartType kind) {
        using var document = ExcelDocument.Create();
        ExcelChart chart = document.AddSheet("Results").AddChart(kind, new OfficeChartData(new[] { "A", "B" }, new[] {
                new OfficeSeries("Status", new[] { 7d, 3d }).WithPointExplosions(new[] { 0, 25 })
            }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
        chart.UpdateData(new ExcelChartData(new[] { "A" }, new[] {
            new ExcelChartSeries("Status", new[] { 8d })
        }));
        Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));
        Assert.Null(snapshot.Data.Series.Single().PointExplosions);
        Assert.NotEmpty(chart.ToImage(OfficeImageExportFormat.Svg).Bytes);
        Assert.Empty(document.ValidateDocument());
    }
}
