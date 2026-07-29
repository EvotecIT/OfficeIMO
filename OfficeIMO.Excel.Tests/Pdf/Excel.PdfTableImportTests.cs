using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using System.Globalization;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void PdfTables_SaveTablesAsExcel_ImportsDetectedTablesAsWorkbookTables() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();

        using var workbook = new MemoryStream();
        PdfExcelTableImportReport report = PdfExcelTableConverterExtensions.SaveTablesAsExcel(
            LoadTables(pdf),
            workbook,
            new PdfExcelTableImportOptions {
                AutoFitColumns = false
            });

        PdfExcelTableImportEntry result = Assert.Single(report.Entries);
        Assert.Equal(1, result.PageNumber);
        Assert.Equal(0, result.TableIndex);
        Assert.Equal(3, result.ColumnCount);
        Assert.Equal(2, result.RowCount);
        Assert.False(result.Truncated);
        Assert.Equal("A1:C3", result.Range);

        byte[] workbookBytes = workbook.ToArray();
        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(workbookBytes), false)) {
            WorksheetPart worksheet = Assert.Single(spreadsheet.WorkbookPart!.WorksheetParts);
            TableDefinitionPart tablePart = Assert.Single(worksheet.TableDefinitionParts);
            Table tableDefinition = tablePart.Table!;
            Assert.Equal(result.TableName, tableDefinition.Name?.Value);
            Assert.Equal("A1:C3", tableDefinition.Reference?.Value);
            Assert.NotNull(tableDefinition.GetFirstChild<AutoFilter>());
        }

        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbookBytes);
        ExcelTableInfo table = Assert.Single(reader.GetTables());
        Assert.Equal(result.TableName, table.Name);
        Assert.Equal(result.SheetName, table.SheetName);
        Assert.Equal(new[] { "Code", "Name", "Qty" }, table.Columns.Select(column => column.Name).ToArray());

        object?[,] values = reader.GetSheet(result.SheetName).ReadRange(result.Range);
        Assert.Equal("Code", values[0, 0]);
        Assert.Equal("A-100", values[1, 0]);
        Assert.Equal(14d, Convert.ToDouble(values[2, 2], CultureInfo.InvariantCulture));
    }

    [Fact]
    public void PdfTables_SaveTablesAsExcel_SupportsNonSeekableDestinationStreams() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Code", "Qty" },
                new[] { "A-100", "2" },
                new[] { "B-200", "3" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 180, 80 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();
        using var workbook = new NonSeekableReadWriteBuffer(Array.Empty<byte>());

        PdfExcelTableImportReport report = PdfExcelTableConverterExtensions.SaveTablesAsExcel(
            LoadTables(pdf),
            workbook,
            new PdfExcelTableImportOptions {
                AutoFitColumns = false
            });

        PdfExcelTableImportEntry result = Assert.Single(report.Entries);
        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbook.ToArray());
        object?[,] values = reader.GetSheet(result.SheetName).ReadRange(result.Range);
        Assert.Equal("A-100", values[1, 0]);
        Assert.Equal(2d, Convert.ToDouble(values[1, 1], CultureInfo.InvariantCulture));
    }

    [Fact]
    public void PdfTables_SaveTablesAsExcel_WritesDetectedNumericColumnsAsNumberCells() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "B-200", "Beta", "14" }
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 70, 170, 60 },
                HeaderRowCount = 1,
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .ToBytes();

        using var workbook = new MemoryStream();
        PdfExcelTableImportReport report = PdfExcelTableConverterExtensions.SaveTablesAsExcel(
            LoadTables(pdf),
            workbook,
            new PdfExcelTableImportOptions {
                AutoFitColumns = false
            });

        PdfExcelTableImportEntry result = Assert.Single(report.Entries);
        byte[] workbookBytes = workbook.ToArray();
        using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(new MemoryStream(workbookBytes), false)) {
            SheetData sheetData = GetOnlySheetData(spreadsheet);
            Cell codeCell = GetCell(sheetData, "A2");
            Cell nameCell = GetCell(sheetData, "B2");
            Cell quantityCell = GetCell(sheetData, "C2");
            Cell secondQuantityCell = GetCell(sheetData, "C3");

            Assert.True(IsTextCell(codeCell));
            Assert.True(IsTextCell(nameCell));
            Assert.True(quantityCell.DataType == null || quantityCell.DataType.Value == CellValues.Number);
            Assert.True(secondQuantityCell.DataType == null || secondQuantityCell.DataType.Value == CellValues.Number);
            Assert.Equal("2", quantityCell.CellValue?.Text);
            Assert.Equal("14", secondQuantityCell.CellValue?.Text);
        }

        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbookBytes);
        object?[,] values = reader.GetSheet(result.SheetName).ReadRange(result.Range);
        Assert.Equal("A-100", values[1, 0]);
        Assert.Equal("Alpha", values[1, 1]);
        Assert.Equal(2d, Convert.ToDouble(values[1, 2], CultureInfo.InvariantCulture));

        using var textWorkbook = new MemoryStream();
        PdfExcelTableConverterExtensions.SaveTablesAsExcel(
            LoadTables(pdf),
            textWorkbook,
            new PdfExcelTableImportOptions {
                AutoFitColumns = false,
                ConvertNumericColumns = false
            });

        using SpreadsheetDocument textSpreadsheet = SpreadsheetDocument.Open(new MemoryStream(textWorkbook.ToArray()), false);
        SheetData textSheetData = GetOnlySheetData(textSpreadsheet);
        Assert.True(IsTextCell(GetCell(textSheetData, "C2")));
    }

    [Fact]
    public void PdfTables_NumericParserHandlesInvoiceNumberText() {
        Assert.True(PdfCore.PdfLogicalTableAnalysis.TryParseNumericValue("$1,234.50", CultureInfo.InvariantCulture, out decimal currency));
        Assert.Equal(1234.50m, currency);

        Assert.True(PdfCore.PdfLogicalTableAnalysis.TryParseNumericValue("(99.95)", CultureInfo.InvariantCulture, out decimal parenthesizedNegative));
        Assert.Equal(-99.95m, parenthesizedNegative);

        Assert.True(PdfCore.PdfLogicalTableAnalysis.TryParseNumericValue("1 234,50", CultureInfo.GetCultureInfo("pl-PL"), out decimal localized));
        Assert.Equal(1234.50m, localized);

        Assert.False(PdfCore.PdfLogicalTableAnalysis.TryParseNumericValue("12%", CultureInfo.InvariantCulture, out _));
    }

    [Fact]
    public void PdfTables_SaveTablesAsExcel_AppliesRowCapsAndKeepsWorkbookValidWhenEmpty() {
        byte[] pdf = PdfCore.PdfDocument.Create(new PdfCore.PdfOptions {
                PageWidth = 420,
                PageHeight = 360,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .KeyValueTable(new[] {
                PdfCore.PdfKeyValueRow.Text("InvoiceId", "INV-001"),
                PdfCore.PdfKeyValueRow.Text("Customer", "Evotec"),
                PdfCore.PdfKeyValueRow.Text("Due", "2026-06-30")
            }, style: new PdfCore.PdfTableStyle {
                ColumnWidthPoints = new List<double?> { 120, 170 },
                CellPaddingX = 6,
                CellPaddingY = 4
            })
            .PageBreak()
            .Paragraph(p => p.Text("No table on this page."))
            .ToBytes();

        using var workbook = new MemoryStream();
        PdfExcelTableImportReport report = PdfExcelTableConverterExtensions.SaveTablesAsExcel(
            LoadTables(pdf, PdfCore.PdfPageRange.From(1, 1)),
            workbook,
            new PdfExcelTableImportOptions {
                MaxRows = 2,
                AutoFitColumns = false
            });

        PdfExcelTableImportEntry result = Assert.Single(report.Entries);
        Assert.Equal(1, result.PageNumber);
        Assert.Equal(2, result.RowCount);
        Assert.Equal(3, result.TotalRowCount);
        Assert.True(result.Truncated);
        Assert.True(report.HasLoss);
        Assert.Throws<InvalidOperationException>(() => report.RequireNoLoss());

        using ExcelDocumentReader reader = ExcelDocumentReader.Open(workbook.ToArray());
        object?[,] values = reader.GetSheet(result.SheetName).ReadRange(result.Range);
        Assert.Equal("Key", values[0, 0]);
        Assert.Equal("InvoiceId", values[1, 0]);
        Assert.Equal("Customer", values[2, 0]);

        using var emptyWorkbook = new MemoryStream();
        PdfExcelTableImportReport emptyReport = PdfExcelTableConverterExtensions.SaveTablesAsExcel(
            LoadTables(pdf, PdfCore.PdfPageRange.From(2, 2)),
            emptyWorkbook,
            new PdfExcelTableImportOptions {
                AutoFitColumns = false
            });

        Assert.Empty(emptyReport.Entries);
        using ExcelDocumentReader emptyReader = ExcelDocumentReader.Open(emptyWorkbook.ToArray());
        object?[,] emptyValues = emptyReader.GetSheet("PDF Tables").ReadRange("A1:A1");
        Assert.Equal("No PDF tables detected.", emptyValues[0, 0]);
    }

    private static PdfCore.PdfLogicalDocument LoadTables(byte[] pdf, params PdfCore.PdfPageRange[] ranges) {
        var layout = new PdfCore.PdfTextLayoutOptions { ForceSingleColumn = true };
        return ranges.Length == 0
            ? PdfCore.PdfLogicalDocument.Load(pdf, layout)
            : PdfCore.PdfLogicalDocument.LoadPageRanges(pdf, layout, ranges);
    }

    private static Cell GetCell(SheetData sheetData, string reference) {
        return sheetData.Descendants<Cell>()
            .Single(cell => string.Equals(cell.CellReference?.Value, reference, StringComparison.OrdinalIgnoreCase));
    }

    private static SheetData GetOnlySheetData(SpreadsheetDocument spreadsheet) {
        WorkbookPart workbookPart = spreadsheet.WorkbookPart ?? throw new InvalidOperationException("Workbook part is missing.");
        WorksheetPart worksheetPart = Assert.Single(workbookPart.WorksheetParts);
        Worksheet worksheet = worksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
        return worksheet.GetFirstChild<SheetData>() ?? throw new InvalidOperationException("SheetData is missing.");
    }

    private static bool IsTextCell(Cell cell) {
        CellValues? dataType = cell.DataType?.Value;
        return dataType == CellValues.SharedString || dataType == CellValues.String || dataType == CellValues.InlineString;
    }
}
