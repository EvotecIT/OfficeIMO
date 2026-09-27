using System;
using System.IO;
using OfficeIMO.Excel;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class ExcelSourceLinkedChartFormatTests {
    [Fact]
    public void ChartAxis_ReportsStylesheetsBeyondTheResolutionWorkBound() {
        using var document = ExcelDocument.Create(new MemoryStream());
        var sheet = document.AddWorksheet("Styles");
        sheet.CellValue(1, 1, "Region"); sheet.CellValue(1, 2, "Score");
        sheet.CellValue(2, 1, "North"); sheet.CellValue(2, 2, .94); sheet.CellAt(2, 2).Percent(0);
        sheet.AddChartFromRange("A1:B2", row: 1, column: 4);
        var formats = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!.CellFormats!;
        for (int index = formats.ChildElements.Count; index <= 100_000; index++)
            formats.Append(new DocumentFormat.OpenXml.Spreadsheet.CellFormat { NumberFormatId = 9 });
        var result = sheet.Range("A1:J12").ExportImage(OfficeImageExportFormat.Svg);
        Assert.Contains(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartAxisNumberFormatApproximation);
    }
    [Fact]
    public void ChartAxis_ReportsMixedSourceNumberFormats() {
        using var document = ExcelDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx"));
        var sheet = document.AddWorksheet("Mixed");
        sheet.CellValue(1, 1, "Region"); sheet.CellValue(1, 2, "Score");
        sheet.CellValue(2, 1, "North"); sheet.CellValue(2, 2, .94); sheet.CellAt(2, 2).Percent(0);
        sheet.CellValue(3, 1, "South"); sheet.CellValue(3, 2, .82);
        sheet.AddChartFromRange("A1:B3", row: 1, column: 4);
        var result = sheet.Range("A1:J12").ExportImage(OfficeImageExportFormat.Svg);
        Assert.Contains(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartAxisNumberFormatApproximation);
    }
    [Theory]
    [InlineData(ExcelChartType.ColumnClustered, true)]
    [InlineData(ExcelChartType.BarClustered, true)]
    [InlineData(ExcelChartType.ColumnClustered, false)]
    public void ChartAxis_UsesSourcePercentageOnlyWhenLinked(ExcelChartType kind, bool sourceLinked) {
        using var document = ExcelDocument.Create(Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx"));
        var sheet = document.AddWorksheet("Source");
        sheet.CellValue(1, 1, "Region"); sheet.CellValue(1, 2, "Score");
        sheet.CellValue(2, 1, "North"); sheet.CellValue(2, 2, .94); sheet.CellAt(2, 2).Percent(0);
        sheet.CellValue(3, 1, "South"); sheet.CellValue(3, 2, .82); sheet.CellAt(3, 2).Percent(0);
        var chart = sheet.AddChartFromRange("A1:B3", row: 1, column: 4, type: kind);
        chart.SetValueAxisNumberFormat("General", sourceLinked);
        var visual = Assert.Single(sheet.Range("A1:J12").CreateVisualSnapshot().Charts).Snapshot;
        string? format = kind == ExcelChartType.BarClustered ? visual.Layout!.HorizontalAxisNumberFormat : visual.Layout!.VerticalAxisNumberFormat;
        Assert.Equal(sourceLinked ? "0%" : "General", format);
    }
}
