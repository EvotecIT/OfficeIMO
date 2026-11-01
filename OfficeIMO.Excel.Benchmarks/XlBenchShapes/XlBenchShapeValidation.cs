using System.Globalization;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using static OfficeIMO.Excel.Benchmarks.XlBenchShapeData;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Checks one saved workbook per implementation outside measurement, reopening it with the SDK and ClosedXML.</summary>
internal static class XlBenchShapeValidation {
    internal static void ValidateSchema(string label, byte[] bytes, byte[]? source) {
        var baseline = source == null ? [] : SchemaErrors(source).ToHashSet(StringComparer.Ordinal);
        string[] introduced = SchemaErrors(bytes).Where(error => !baseline.Contains(error)).Take(5).ToArray();
        if (introduced.Length > 0) {
            throw new InvalidDataException($"{label} failed Open XML validation: {string.Join(" | ", introduced)}");
        }
    }

    private static IEnumerable<string> SchemaErrors(byte[] bytes) {
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        return new OpenXmlValidator().Validate(document).Select(static error => $"{error.Part?.Uri} {error.Path?.XPath}: {error.Description}").ToArray();
    }

    internal static void Write(string label, byte[] bytes) {
        ValidateSchema(label, bytes, null);
        using var workbook = new XLWorkbook(new MemoryStream(bytes));
        var sheet = workbook.Worksheet("Data");
        Equal(label, "last row", WriteRows + 2, sheet.LastRowUsed()!.RowNumber());
        Equal(label, "first text", WriteStrings[0], sheet.Cell(2, 1).GetString());
        Equal(label, "last number", WriteNumbers[^1], sheet.Cell(WriteRows + 1, 2).GetDouble());
        Equal(label, "last date", WriteDates[^1], sheet.Cell(WriteRows + 1, 3).GetDateTime());
        Equal(label, "total formula", $"SUM(B2:B{WriteRows + 1})", sheet.Cell(WriteRows + 2, 2).FormulaA1);
    }

    internal static void Report(string label, byte[] bytes) {
        ValidateSchema(label, bytes, null);
        using (var workbook = new XLWorkbook(new MemoryStream(bytes))) {
            var sheet = workbook.Worksheet(ReportSheet);
            Equal(label, "last row", 1 + StockWeeks, sheet.LastRowUsed()!.RowNumber());
            Equal(label, "last price", Prices[^1][^1], sheet.Cell(1 + StockWeeks, 2 + StockSymbols).GetDouble());
            Equal(label, "date", WeekEndings[^1], sheet.Cell(1 + StockWeeks, 2).GetDateTime());
            Equal(label, "price format", PriceFormat, sheet.Cell(2, 3).Style.NumberFormat.Format);
            Equal(label, "date format", DateFormat, sheet.Cell(2, 2).Style.NumberFormat.Format);
            Equal(label, "header bold", true, sheet.Cell(1, 3).Style.Font.Bold);
        }

        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var worksheetPart = document.WorkbookPart!.WorksheetParts.Single();
        Equal(label, "conditional rules", 2, worksheetPart.Worksheet!.Descendants<ConditionalFormattingRule>().Count());
        var chart = worksheetPart.DrawingsPart!.ChartParts.Single().ChartSpace!;
        Equal(label, "chart series", StockSymbols, chart.Descendants<C.LineChartSeries>().Count());
        Equal(label, "differential formats", true, document.WorkbookPart.WorkbookStylesPart!.Stylesheet!.DifferentialFormats!.Elements<DifferentialFormat>().Count() >= 2);
    }

    internal static void Edit(string label, byte[] bytes) {
        ValidateSchema(label, bytes, null);
        using var workbook = new XLWorkbook(new MemoryStream(bytes));
        var sheet = workbook.Worksheet("Numbers");
        Equal(label, "last row", EditLastRowAfterDelete, sheet.LastRowUsed()!.RowNumber());
        for (int i = 0; i < SurvivingRows.Length; i++) {
            int row = 2 + i;
            int source = SurvivingRows[i];
            Equal(label, $"A{row}", EditColumnAValue, sheet.Cell(row, 1).GetDouble());
            Equal(label, $"T{row}", EditValues[source][EditColumns - 1], sheet.Cell(row, EditColumns).GetDouble());
            Equal(label, $"formula {row}", EditFormula(row), sheet.Cell(row, EditSumColumn).FormulaA1);
            Near(label, $"total {row}", EditExpectedTotal(i), sheet.Cell(row, EditSumColumn).CachedValue.GetNumber());
            Equal(label, $"bold {row}", (source + 2) % 2 == 0, sheet.Cell(row, 2).Style.Font.Bold);
        }
    }

    internal static void Insert(string label, byte[] bytes) {
        ValidateSchema(label, bytes, null);
        using var workbook = new XLWorkbook(new MemoryStream(bytes));
        var sheet = workbook.Worksheet("Numbers");
        int sumColumn = EditSumColumn + InsertCount;
        Equal(label, "shifted header", "Column_02", sheet.Cell(1, InsertAt + InsertCount).GetString());
        for (int i = 0; i < EditRows; i++) {
            int row = 2 + i;
            Equal(label, $"inserted {row}", InsertedValue, sheet.Cell(row, InsertAt).GetDouble());
            Equal(label, $"shifted {row}", EditValues[i][1], sheet.Cell(row, InsertAt + InsertCount).GetDouble());
            Equal(label, $"formula {row}", InsertFormula(row), sheet.Cell(row, sumColumn).FormulaA1);
            Near(label, $"total {row}", InsertExpectedTotal(i), sheet.Cell(row, sumColumn).CachedValue.GetNumber());
            Equal(label, $"bold {row}", row % 2 == 0, sheet.Cell(row, sumColumn - 1).Style.Font.Bold);
        }
    }

    internal static void Template(string label, byte[] bytes) {
        ValidateSchema(label, bytes, TemplateWorkbook);
        int lastRow = TemplateFirstRow + TemplateRows - 1;
        using (var workbook = new XLWorkbook(new MemoryStream(bytes))) {
            var sheet = workbook.Worksheet("Report");
            TemplateLine line = Lines[^1];
            Equal(label, "title", TemplateTitle, sheet.Cell("A1").GetString());
            Equal(label, "date", TemplateDate, sheet.Cell("B2").GetDateTime());
            Equal(label, "region", TemplateRegion, sheet.Cell("B3").GetString());
            Equal(label, "last product", line.Product, sheet.Cell(lastRow, 2).GetString());
            Equal(label, "last quantity", (double)line.Quantity, sheet.Cell(lastRow, 4).GetDouble());
            Equal(label, "last date", line.Date, sheet.Cell(lastRow, 1).GetDateTime());
            Equal(label, "amount formula", $"D{lastRow}*E{lastRow}", sheet.Cell(lastRow, 6).FormulaA1);
            Equal(label, "summary preserved", "Total", sheet.Cell("H11").GetString());
            Equal(label, "merge preserved", 1, sheet.MergedRanges.Count);
        }

        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        var workbookPart = document.WorkbookPart!;
        var worksheetPart = workbookPart.WorksheetParts.Single();
        var worksheet = worksheetPart.Worksheet!;
        Equal(label, "table range", $"A5:F{lastRow}", worksheetPart.TableDefinitionParts.Single().Table!.Reference!.Value);
        Equal(label, "validation", "C6:C5000", worksheet.Descendants<DataValidation>().Single().SequenceOfReferences!.InnerText);
        Equal(label, "conditional format", "F6:F5000", worksheet.Elements<ConditionalFormatting>().Single().SequenceOfReferences!.InnerText);
        Equal(label, "chart categories", "Report!$H$6:$H$10", worksheetPart.DrawingsPart!.ChartParts.Single().ChartSpace!.Descendants<C.CategoryAxisData>().Single().Descendants<C.Formula>().Single().Text);
        Equal(label, "defined names", 3, workbookPart.Workbook!.DefinedNames!.Elements<DefinedName>().Count());
    }

    internal static void Properties(string label, byte[] bytes) {
        ValidateSchema(label, bytes, null);
        using var stream = new MemoryStream(bytes, writable: false);
        using var document = SpreadsheetDocument.Open(stream, isEditable: false);
        Equal(label, "title", PropertyTitle, document.PackageProperties.Title);
        Equal(label, "category", PropertyCategory, document.PackageProperties.Category);
        using var original = new XLWorkbook(new MemoryStream(PropertiesWorkbook));
        using var saved = new XLWorkbook(new MemoryStream(bytes));
        var expected = original.Worksheet("Numbers");
        var actual = saved.Worksheet("Numbers");
        for (int row = 1; row <= PropertyRows + 1; row++) {
            for (int column = 1; column <= PropertyColumns; column++) {
                Equal(label, $"R{row}C{column}", expected.Cell(row, column).Value.ToString(CultureInfo.InvariantCulture), actual.Cell(row, column).Value.ToString(CultureInfo.InvariantCulture));
            }
        }
    }

    private static void Near(string label, string field, double expected, double actual) {
        if (Math.Abs(expected - actual) > 1e-6) {
            throw new InvalidDataException($"{label}: {field} expected {expected:R}, found {actual:R}.");
        }
    }

    private static void Equal<T>(string label, string field, T expected, T actual) {
        if (!EqualityComparer<T>.Default.Equals(expected, actual)) {
            throw new InvalidDataException($"{label}: {field} expected '{expected}', found '{actual}'.");
        }
    }
}
