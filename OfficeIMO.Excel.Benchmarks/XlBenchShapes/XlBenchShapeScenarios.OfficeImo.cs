using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using static OfficeIMO.Excel.Benchmarks.XlBenchShapeData;

namespace OfficeIMO.Excel.Benchmarks;

internal static partial class XlBenchShapeScenarios {
    internal static byte[] OfficeImoWrite() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Data");
        sheet.CellValue(1, 1, "Name");
        sheet.CellValue(1, 2, "Amount");
        sheet.CellValue(1, 3, "Date");
        for (int i = 0; i < WriteRows; i++) {
            int row = i + 2;
            sheet.CellValue(row, 1, WriteStrings[i]);
            sheet.CellValue(row, 2, WriteNumbers[i]);
            sheet.CellValue(row, 3, WriteDates[i]);
        }
        int sumRow = WriteRows + 2;
        sheet.CellValue(sumRow, 1, "Total");
        sheet.CellFormula(sumRow, 2, $"SUM(B2:B{WriteRows + 1})");
        return Save(document);
    }

    internal static long OfficeImoReadAll() {
        using var document = ExcelDocument.Load(new MemoryStream(ReadWorkbook));
        long checksum = 0;
        foreach (var cell in document.Sheets[0].EnumerateCells()) {
            checksum += (cell.Value?.ToString() ?? string.Empty).Length;
        }
        return checksum;
    }

    internal static long OfficeImoSheets() {
        using var document = ExcelDocument.Load(new MemoryStream(SheetsWorkbook));
        long cells = 0;
        long text = 0;
        foreach (var sheet in document.Sheets) {
            foreach (var cell in sheet.EnumerateCells()) {
                cells++;
                if (cell.Value is string value) {
                    text += value.Length;
                }
            }
        }
        return checksumCheck(cells, text);

        static long checksumCheck(long cells, long text) => cells == SheetsCellCount ? text : throw new InvalidDataException($"Read {cells} cells.");
    }

    internal static long OfficeImoReadWhileWriting() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Data");
        int half = WriteRows / 2;
        OfficeImoWriteRows(sheet, 0, half);
        long cells = OfficeImoCountCells(sheet);
        OfficeImoWriteRows(sheet, half, WriteRows);
        return cells + OfficeImoCountCells(sheet);
    }

    private static void OfficeImoWriteRows(ExcelSheet sheet, int first, int end) {
        for (int i = first; i < end; i++) {
            sheet.CellValue(i + 1, 1, WriteStrings[i]);
            sheet.CellValue(i + 1, 2, WriteNumbers[i]);
            sheet.CellValue(i + 1, 3, WriteDates[i]);
        }
    }

    private static long OfficeImoCountCells(ExcelSheet sheet) {
        long cells = 0;
        foreach (var cell in sheet.EnumerateCells()) {
            if (cell.Value != null) {
                cells++;
            }
        }
        return cells;
    }

    internal static long OfficeImoWideRead() {
        using var document = ExcelDocument.Load(new MemoryStream(WideWorkbook));
        using var reader = document.CreateDataReader();
        long cells = 0;
        while (reader.Read()) {
            for (int column = 0; column < reader.FieldCount; column++) {
                if (!reader.IsDBNull(column)) {
                    cells++;
                }
            }
        }
        return cells;
    }

    internal static long OfficeImoEditedRead() {
        using var document = ExcelDocument.Load(new MemoryStream(ReadWorkbook));
        var sheet = document.Sheets[0];
        sheet.CellValue(1, 4, EditedValue);
        return OfficeImoCountCells(sheet);
    }

    internal static long OfficeImoEditedSheetsRead() {
        using var document = ExcelDocument.Load(new MemoryStream(SheetsWorkbook));
        document.Sheets[0].CellValue(2, 1, EditedText);
        long cells = 0;
        foreach (var sheet in document.Sheets) {
            cells += OfficeImoCountCells(sheet);
        }
        return cells == SheetsCellCount ? cells : throw new InvalidDataException($"Read {cells} cells.");
    }

    internal static byte[] OfficeImoReport() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet(ReportSheet);
        sheet.CellValue(1, 1, "Week");
        sheet.CellValue(1, 2, "Week Ending");
        for (int s = 0; s < StockSymbols; s++) {
            sheet.CellValue(1, 3 + s, Symbols[s]);
        }
        for (int column = 1; column <= 2 + StockSymbols; column++) {
            sheet.CellBold(1, column);
        }
        for (int w = 0; w < StockWeeks; w++) {
            int row = 2 + w;
            sheet.CellValue(row, 1, w + 1);
            sheet.CellValue(row, 2, WeekEndings[w]);
            sheet.FormatCell(row, 2, DateFormat);
            for (int s = 0; s < StockSymbols; s++) {
                sheet.CellValue(row, 3 + s, Prices[w][s]);
                sheet.FormatCell(row, 3 + s, PriceFormat);
            }
        }
        sheet.AutoFitColumn(2);
        sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
            Range = ChartPriceRange, Type = "expression", Formulas = ["C3>C2"],
            DifferentialFillColorArgb = "C6EFCE", DifferentialFontColorArgb = "006100"
        });
        sheet.AddConditionalFormattingRule(new ExcelConditionalFormattingInfo {
            Range = ChartPriceRange, Type = "expression", Formulas = ["C3<C2"],
            DifferentialFillColorArgb = "FFC7CE", DifferentialFontColorArgb = "9C0006"
        });
        var chart = sheet.AddChart(OfficeChartType.Line,
            new ExcelChartDataRange(ReportSheet, startRow: 1, startColumn: 2, categoryCount: StockWeeks, seriesCount: StockSymbols, hasHeaderRow: true),
            new ExcelChartOptions {
                Row = 1, Column = 3 + StockSymbols, Title = ReportTitle,
                Width = OfficeLength.Pixels(14 * 64), Height = OfficeLength.Pixels(30 * 20)
            });
        chart.SetLegend(OfficeChartLegendPosition.Right);
        return Save(document);
    }

    internal static double OfficeImoEdit(Stream? output) {
        using var document = ExcelDocument.Load(new MemoryStream(EditWorkbook));
        var sheet = document.Sheets[0];
        for (int i = RowsToDelete.Length - 1; i >= 0; i--) {
            sheet.DeleteRows(RowsToDelete[i]);
        }
        int lastRow = EditLastRowAfterDelete;
        for (int row = 2; row <= lastRow; row++) {
            sheet.CellValue(row, 1, EditColumnAValue);
        }
        document.Calculate();
        sheet.TryGetCachedFormulaValue(lastRow, EditSumColumn, out string? cached);
        double total = double.Parse(cached!, NumberStyles.Float, CultureInfo.InvariantCulture);
        if (output != null) {
            document.Save(output);
        }
        return total;
    }

    internal static double OfficeImoInsert(Stream? output) {
        using var document = ExcelDocument.Load(new MemoryStream(EditWorkbook));
        var sheet = document.Sheets[0];
        sheet.InsertColumns(InsertAt, InsertCount);
        int lastRow = EditRows + 1;
        for (int row = 2; row <= lastRow; row++) {
            for (int column = InsertAt; column < InsertAt + InsertCount; column++) {
                sheet.CellValue(row, column, InsertedValue);
            }
        }
        document.Calculate();
        sheet.TryGetCachedFormulaValue(lastRow, EditSumColumn + InsertCount, out string? cached);
        double total = double.Parse(cached!, NumberStyles.Float, CultureInfo.InvariantCulture);
        if (output != null) {
            document.Save(output);
        }
        return total;
    }

    internal static double OfficeImoCalc(Stream? output) {
        using var document = ExcelDocument.Load(new MemoryStream(CalcWorkbook));
        var products = document[CalcProductsSheet];
        for (int i = 0; i < CalcProducts; i++) {
            products.CellValue(2 + i, CalcUnitPriceColumn, CalcNewPrice(i));
        }
        document.Calculate();
        document[CalcSummarySheet].TryGetCachedFormulaValue(CalcGrandTotalRow, CalcStatColumn, out string? cached);
        double total = double.Parse(cached!, NumberStyles.Float, CultureInfo.InvariantCulture);
        if (output != null) {
            document.Save(output);
        }
        return total;
    }

    internal static byte[] OfficeImoTemplate() {
        using var document = ExcelDocument.Load(new MemoryStream(TemplateWorkbook));
        var sheet = document.Sheets.First(candidate => candidate.Name == "Report");
        var (titleRow, titleColumn) = ParseCell(document.GetNamedRange("ReportTitle")!);
        sheet.CellValue(titleRow, titleColumn, TemplateTitle);
        var (dateRow, dateColumn) = ParseCell(document.GetNamedRange("ReportDate")!);
        sheet.CellValue(dateRow, dateColumn, TemplateDate);
        var (regionRow, regionColumn) = ParseCell(document.GetNamedRange("Region")!);
        sheet.CellValue(regionRow, regionColumn, TemplateRegion);
        for (int i = 0; i < Lines.Length; i++) {
            int row = TemplateFirstRow + i;
            TemplateLine line = Lines[i];
            sheet.CellValue(row, 1, line.Date);
            sheet.CellValue(row, 2, line.Product);
            sheet.CellValue(row, 3, line.Category);
            sheet.CellValue(row, 4, line.Quantity);
            sheet.CellValue(row, 5, line.UnitPrice);
            sheet.CellFormula(row, 6, $"D{row}*E{row}");
        }
        sheet.ResizeTable("Sales", $"A5:F{TemplateFirstRow + TemplateRows - 1}");
        return Save(document);
    }

    internal static byte[] OfficeImoProperties() {
        using var document = ExcelDocument.Load(new MemoryStream(PropertiesWorkbook));
        document.BuiltinDocumentProperties.Title = PropertyTitle;
        document.BuiltinDocumentProperties.Category = PropertyCategory;
        return Save(document);
    }

    private static byte[] Save(ExcelDocument document) {
        using var output = new MemoryStream();
        document.Save(output);
        return output.ToArray();
    }
}
