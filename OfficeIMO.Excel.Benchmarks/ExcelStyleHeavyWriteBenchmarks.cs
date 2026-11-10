using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Writes cells that each carry one of 240 distinct font, fill and number-format combinations, the shape of a
/// formatted export. The formatted report lanes reuse a handful of styles, so they never grow the style table.
/// Both outputs are checked through the Open XML SDK for every cell and at least 240 distinct cell formats.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Performance", "StyleHeavyWrite")]
public class ExcelStyleHeavyWriteBenchmarks {
    private const int Columns = 6;
    private static readonly string[] Fills = ["FFF2CC", "DDEBF7", "E2EFDA", "FCE4D6", "EDEDED", "D9D2E9"];
    private static readonly string[] Formats = ["0", "0.00", "#,##0", "#,##0.00", "0%", "0.0%", "0.00E+00", "yyyy-mm-dd", "[$-409]mmm d", "#,##0.00 \"EUR\""];

    [Params(20_000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        Validate(OfficeIMO_Write_Styled());
        Validate(ClosedXML_Write_Styled());
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO_Write_Styled() {
        using var output = new MemoryStream();
        using (var document = ExcelDocument.Create(output)) {
            ExcelSheet sheet = document.AddSheet("Styled");
            for (int row = 1; row <= RowCount; row++) {
                for (int column = 1; column <= Columns; column++) {
                    (bool bold, bool italic, string fill, string format) = StyleFor(row, column);
                    sheet.Cell(row, column).SetValue(row * column + 0.25).SetBold(bold).SetItalic(italic)
                        .SetFillColor(OfficeIMO.OfficeColor.Parse(fill)).SetNumberFormat(format);
                }
            }
            document.Save(output);
        }
        return output.ToArray();
    }

    [Benchmark]
    public byte[] ClosedXML_Write_Styled() {
        using var output = new MemoryStream();
        using (var workbook = new XLWorkbook()) {
            IXLWorksheet sheet = workbook.Worksheets.Add("Styled");
            for (int row = 1; row <= RowCount; row++) {
                for (int column = 1; column <= Columns; column++) {
                    (bool bold, bool italic, string fill, string format) = StyleFor(row, column);
                    IXLCell cell = sheet.Cell(row, column);
                    cell.Value = row * column + 0.25;
                    cell.Style.Font.Bold = bold;
                    cell.Style.Font.Italic = italic;
                    cell.Style.Fill.BackgroundColor = XLColor.FromHtml("#" + fill);
                    cell.Style.NumberFormat.Format = format;
                }
            }
            workbook.SaveAs(output);
        }
        return output.ToArray();
    }

    /// <summary>Cycles through every combination so the style table holds all 240 of them.</summary>
    private static (bool Bold, bool Italic, string Fill, string Format) StyleFor(int row, int column) {
        int index = ((row - 1) * Columns + column - 1) % 240;
        return ((index & 1) != 0, (index & 2) != 0, Fills[index / 4 % Fills.Length], Formats[index / 24]);
    }

    private void Validate(byte[] package) {
        using var document = SpreadsheetDocument.Open(new MemoryStream(package, writable: false), false);
        WorkbookPart workbook = document.WorkbookPart!;
        int cells = workbook.WorksheetParts.Single().Worksheet!.Descendants<Cell>().Count(cell => cell.StyleIndex != null);
        if (cells != RowCount * Columns) throw new InvalidDataException($"Expected {RowCount * Columns} styled cells, found {cells}.");
        int formats = workbook.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().Count();
        if (formats < 240) throw new InvalidDataException($"Expected at least 240 cell formats, found {formats}.");
    }
}
