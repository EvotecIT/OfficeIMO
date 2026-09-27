using System;
using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class ExcelSourceLinkedChartFormatTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ChartAxis_ReportsOverflowingSourceNumberFormatIds(bool custom) {
        using var document = ExcelDocument.Create(new MemoryStream());
        var sheet = document.AddWorksheet("Overflow");
        sheet.CellValue(1, 1, "Region"); sheet.CellValue(1, 2, "Score");
        sheet.CellValue(2, 1, "North"); sheet.CellValue(2, 2, .94); sheet.CellAt(2, 2).Percent(0);
        sheet.AddChartFromRange("A1:B2", row: 1, column: 4);
        var stylesheet = document.WorkbookPartRoot!.WorkbookStylesPart!.Stylesheet!;
        var id = new DocumentFormat.OpenXml.UInt32Value { InnerText = "4294967296" };
        if (custom) {
            var formats = stylesheet.NumberingFormats ?? new DocumentFormat.OpenXml.Spreadsheet.NumberingFormats();
            if (formats.Parent == null) stylesheet.AddChild(formats, true);
            formats.Append(new DocumentFormat.OpenXml.Spreadsheet.NumberingFormat { NumberFormatId = id, FormatCode = "0%" });
        } else stylesheet.CellFormats!.Elements<DocumentFormat.OpenXml.Spreadsheet.CellFormat>().Last().NumberFormatId = id;
        var result = sheet.Range("A1:J12").ExportImage(OfficeImageExportFormat.Svg);
        Assert.Contains(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartAxisNumberFormatApproximation);
    }
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    public void ChartAxis_UsesCellThenRowThenColumnNumberFormats(int precedence) {
        using var document = ExcelDocument.Create(new MemoryStream());
        var sheet = document.AddWorksheet("Inherited");
        sheet.CellValue(1, 1, "Region"); sheet.CellValue(1, 2, "Score");
        sheet.CellValue(2, 1, "North"); sheet.CellValue(2, 2, .94); sheet.CellAt(2, 2).Percent(0);
        var worksheet = sheet.WorksheetPart.Worksheet;
        var row = worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.SheetData>()!.Elements<DocumentFormat.OpenXml.Spreadsheet.Row>().Single(item => item.RowIndex!.Value == 2);
        var cell = row.Elements<DocumentFormat.OpenXml.Spreadsheet.Cell>().Single(item => item.CellReference!.Value == "B2");
        uint percentage = cell.StyleIndex!.Value;
        cell.StyleIndex = precedence == 2 ? 0U : null;
        worksheet.AddChild(new DocumentFormat.OpenXml.Spreadsheet.Columns(new DocumentFormat.OpenXml.Spreadsheet.Column { Min = 2, Max = 2, Style = precedence is 0 or 3 ? percentage : 0U }), true);
        if (precedence > 0) { row.StyleIndex = precedence == 3 ? 0U : percentage; row.CustomFormat = precedence != 3; }
        sheet.AddChartFromRange("A1:B2", row: 1, column: 4);
        var visual = Assert.Single(sheet.Range("A1:J12").CreateVisualSnapshot().Charts).Snapshot;
        Assert.Equal(precedence == 2 ? "General" : "0%", visual.Layout!.VerticalAxisNumberFormat);
    }

    [Fact]
    public void ChartAxis_ReportsUnresolvedNumericCategorySourceFormats() {
        using var document = ExcelDocument.Create(new MemoryStream());
        var sheet = document.AddWorksheet("Categories");
        sheet.CellValue(1, 1, "Fraction"); sheet.CellValue(1, 2, "Count");
        sheet.CellValue(2, 1, .94); sheet.CellAt(2, 1).Percent(0); sheet.CellValue(2, 2, 2);
        sheet.CellValue(3, 1, .82); sheet.CellValue(3, 2, 3);
        sheet.AddChartFromRange("A1:B3", row: 1, column: 4);
        var category = sheet.WorksheetPart.DrawingsPart!.ChartParts.Single().ChartSpace!.Descendants<DocumentFormat.OpenXml.Drawing.Charts.CategoryAxisData>().Single();
        category.RemoveAllChildren();
        category.Append(new DocumentFormat.OpenXml.Drawing.Charts.NumberReference(new DocumentFormat.OpenXml.Drawing.Charts.Formula("Categories!$A$2:$A$3"),
            new DocumentFormat.OpenXml.Drawing.Charts.NumberingCache(new DocumentFormat.OpenXml.Drawing.Charts.FormatCode("General"),
                new DocumentFormat.OpenXml.Drawing.Charts.PointCount { Val = 2 },
                new DocumentFormat.OpenXml.Drawing.Charts.NumericPoint(new DocumentFormat.OpenXml.Drawing.Charts.NumericValue("0.94")) { Index = 0 },
                new DocumentFormat.OpenXml.Drawing.Charts.NumericPoint(new DocumentFormat.OpenXml.Drawing.Charts.NumericValue("0.82")) { Index = 1 })));
        var result = sheet.Range("A1:J12").ExportImage(OfficeImageExportFormat.Svg);
        Assert.Contains(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartAxisNumberFormatApproximation);
    }
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
