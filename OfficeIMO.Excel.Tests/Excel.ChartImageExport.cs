using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests;

public sealed class ExcelChartImageExportTests {
    [Fact]
    public void LoadedChartProjectsAndRendersAtAnchorAndRequestedSize() {
        using var bytes = new MemoryStream();
        using (var book = ExcelDocument.Create()) {
            book.AddSheet("Report").AddChart(OfficeChartType.ColumnClustered, new ExcelChartData(new[] { "Q1", "Q2" },
                new[] { new ExcelChartSeries("Revenue", new[] { 12D, 18D }) }),
                new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1, Width = OfficeLength.Pixels(480), Height = OfficeLength.Pixels(300) });
            book.Save(bytes);
        }

        bytes.Position = 0;
        using var loaded = ExcelDocument.Load(bytes);
        ExcelChart chart = Assert.Single(loaded.GetSheet("Report").Charts);
        Assert.True(chart.TryGetOfficeSnapshot(out var snapshot));
        Assert.Equal(new[] { "Q1", "Q2" }, snapshot.Data.Categories);
        Assert.Equal(new[] { 12D, 18D }, Assert.Single(snapshot.Data.Series).Values);

        OfficeImageExportResult png = chart.ToImage(new OfficeChartImageExportOptions());
        Assert.Equal((480, 300), (png.Width, png.Height));

        OfficeImageExportResult svg = chart.ToImage(new OfficeChartImageExportOptions {
            Format = OfficeImageExportFormat.Svg, Width = OfficeLength.Inches(4), Height = OfficeLength.Inches(2.5)
        });
        Assert.Equal(OfficeImageExportFormat.Svg, svg.Format);
        Assert.Equal((384, 240), (svg.Width, svg.Height));
    }

    [Fact]
    public void UnsupportedProjectionReportsItsReason() {
        using var book = ExcelDocument.Create();
        var sheet = book.AddSheet("Results");
        var chart = sheet.AddChart(OfficeChartType.ColumnClustered, new OfficeChartData(new[] { "A" }, new[] {
            new OfficeSeries("Count", new[] { 100D }),
            new OfficeSeries("Ratio", new[] { 1D }, null, null, null, true,
                renderKind: OfficeChartType.Line, axisGroup: OfficeChartAxisGroup.Secondary)
        }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
        sheet.WorksheetPart.DrawingsPart!.ChartParts.Single().ChartSpace!.Descendants<C.ValueAxis>()
            .Single(axis => axis.AxisPosition!.Val!.Value == C.AxisPositionValues.Right)
            .AddChild(new C.DisplayUnits(new C.BuiltInUnit { Val = C.BuiltInUnitValues.Thousands }), true);

        var error = Assert.Throws<NotSupportedException>(() => chart.ToImage(new OfficeChartImageExportOptions()));
        Assert.Contains("secondary numeric axis", error.Message);
        Assert.False(chart.TryGetOfficeSnapshot(out _));
    }

    [Fact]
    public void LoadedChartResolvesSchemeAndAutomaticColorsAgainstTheWorkbookTheme() {
        using var bytes = new MemoryStream();
        using (var book = ExcelDocument.Create()) {
            book.AddSheet("Report").AddChart(OfficeChartType.ColumnClustered, new ExcelChartData(new[] { "Q1", "Q2" }, new[] {
                new ExcelChartSeries("North", new[] { 4D, 6D }),
                new ExcelChartSeries("South", new[] { 5D, 3D }),
                new ExcelChartSeries("East", new[] { 2D, 7D })
            }), new OfficeIMO.Excel.ExcelChartOptions { Row = 1, Column = 1 });
            book.Save(bytes);
        }
        using (var package = SpreadsheetDocument.Open(bytes, true)) {
            WorkbookPart workbook = package.WorkbookPart!;
            ChartImageThemeFixture.Apply(workbook.ThemePart!,
                workbook.WorksheetParts.Where(part => part.DrawingsPart != null).SelectMany(part => part.DrawingsPart!.ChartParts));
        }

        bytes.Position = 0;
        using var loaded = ExcelDocument.Load(bytes);
        string svg = Encoding.UTF8.GetString(Assert.Single(loaded.GetSheet("Report").Charts)
            .ToImage(new OfficeChartImageExportOptions { Format = OfficeImageExportFormat.Svg }).Bytes);
        ChartImageThemeFixture.AssertSourceThemeColors(svg);
    }
}
