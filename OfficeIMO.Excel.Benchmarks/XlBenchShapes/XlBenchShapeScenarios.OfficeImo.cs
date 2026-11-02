using System.Globalization;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using static OfficeIMO.Excel.Benchmarks.XlBenchShapeData;

namespace OfficeIMO.Excel.Benchmarks;

internal static partial class XlBenchShapeScenarios {
    internal static byte[] OfficeImoWrite() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet("Data");
        sheet.Cell(1, 1).SetValue("Name");
        sheet.Cell(1, 2).SetValue("Amount");
        sheet.Cell(1, 3).SetValue("Date");
        for (int i = 0; i < WriteRows; i++) {
            int row = i + 2;
            sheet.Cell(row, 1).SetValue(WriteStrings[i]);
            sheet.Cell(row, 2).SetValue(WriteNumbers[i]);
            sheet.Cell(row, 3).SetValue(WriteDates[i]);
        }
        int sumRow = WriteRows + 2;
        sheet.Cell(sumRow, 1).SetValue("Total");
        sheet.Cell(sumRow, 2).SetFormula($"SUM(B2:B{WriteRows + 1})");
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
            sheet.Cell(i + 1, 1).SetValue(WriteStrings[i]);
            sheet.Cell(i + 1, 2).SetValue(WriteNumbers[i]);
            sheet.Cell(i + 1, 3).SetValue(WriteDates[i]);
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
        sheet.Cell(1, 4).SetValue(EditedValue);
        return OfficeImoCountCells(sheet);
    }

    internal static long OfficeImoEditedSheetsRead() {
        using var document = ExcelDocument.Load(new MemoryStream(SheetsWorkbook));
        document.Sheets[0].Cell(2, 1).SetValue(EditedText);
        long cells = 0;
        foreach (var sheet in document.Sheets) {
            cells += OfficeImoCountCells(sheet);
        }
        return cells == SheetsCellCount ? cells : throw new InvalidDataException($"Read {cells} cells.");
    }

    internal static byte[] OfficeImoReport() {
        using var document = ExcelDocument.Create();
        var sheet = document.AddSheet(ReportSheet);
        sheet.Cell(1, 1).SetValue("Week");
        sheet.Cell(1, 2).SetValue("Week Ending");
        for (int s = 0; s < StockSymbols; s++) {
            sheet.Cell(1, 3 + s).SetValue(Symbols[s]);
        }
        for (int column = 1; column <= 2 + StockSymbols; column++) {
            sheet.Cell(1, column).SetBold();
        }
        for (int w = 0; w < StockWeeks; w++) {
            int row = 2 + w;
            sheet.Cell(row, 1).SetValue(w + 1);
            sheet.Cell(row, 2).SetValue(WeekEndings[w]);
            sheet.Cell(row, 2).SetNumberFormat(DateFormat);
            for (int s = 0; s < StockSymbols; s++) {
                sheet.Cell(row, 3 + s).SetValue(Prices[w][s]);
                sheet.Cell(row, 3 + s).SetNumberFormat(PriceFormat);
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
            sheet.Cell(row, 1).SetValue(EditColumnAValue);
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
                sheet.Cell(row, column).SetValue(InsertedValue);
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
            products.Cell(2 + i, CalcUnitPriceColumn).SetValue(CalcNewPrice(i));
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
        document.GetNamedRange("ReportTitle")!.FirstCell.SetValue(TemplateTitle);
        document.GetNamedRange("ReportDate")!.FirstCell.SetValue(TemplateDate);
        document.GetNamedRange("Region")!.FirstCell.SetValue(TemplateRegion);
        for (int i = 0; i < Lines.Length; i++) {
            int row = TemplateFirstRow + i;
            TemplateLine line = Lines[i];
            sheet.Cell(row, 1).SetValue(line.Date);
            sheet.Cell(row, 2).SetValue(line.Product);
            sheet.Cell(row, 3).SetValue(line.Category);
            sheet.Cell(row, 4).SetValue(line.Quantity);
            sheet.Cell(row, 5).SetValue(line.UnitPrice);
            sheet.Cell(row, 6).SetFormula($"D{row}*E{row}");
        }
        sheet.ResizeTable("Sales", $"A5:F{TemplateFirstRow + TemplateRows - 1}");
        return Save(document);
    }

    internal static byte[] OfficeImoProperties() {
        using var document = ExcelDocument.Load(new MemoryStream(PropertiesWorkbook));
        document.Properties.Title = PropertyTitle;
        document.Properties.Category = PropertyCategory;
        return Save(document);
    }

    private static byte[] Save(ExcelDocument document) {
        using var output = new MemoryStream();
        document.Save(output);
        return output.ToArray();
    }
}
