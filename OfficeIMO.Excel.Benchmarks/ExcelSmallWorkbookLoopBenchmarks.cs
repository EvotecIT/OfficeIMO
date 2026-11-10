using System.Globalization;
using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Opens 200 small, distinct workbooks one after another and reads every value, the shape of a script that
/// processes a folder of exports. Per-open costs (package parsing, style and shared-string setup) dominate
/// here, which the large single-workbook lanes hide. All three libraries must return the same checksum.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Performance", "SmallWorkbookLoop")]
public class ExcelSmallWorkbookLoopBenchmarks {
    private const int Files = 200;
    private const int Rows = 25;
    private const int Columns = 6;
    private byte[][] _workbooks = [];

    [GlobalSetup]
    public void Setup() {
        _workbooks = new byte[Files][];
        for (int file = 0; file < Files; file++) {
            using var output = new MemoryStream();
            using (var document = ExcelDocument.Create(output)) {
                ExcelSheet sheet = document.AddSheet("Export");
                for (int row = 1; row <= Rows; row++) {
                    for (int column = 1; column <= Columns; column++) {
                        if (column % 2 == 1) sheet.Cell(row, column).SetValue($"F{file}-R{row}-C{column}");
                        else sheet.Cell(row, column).SetValue(file * 1000 + row * 10 + column + 0.5);
                    }
                }
                document.Save(output);
            }
            _workbooks[file] = output.ToArray();
        }

        double expected = OfficeIMO_Open_Read();
        if (ClosedXML_Open_Read() != expected || OpenXml_Open_Read() != expected) {
            throw new InvalidDataException("The libraries read different values from the same workbooks.");
        }
    }

    [Benchmark(Baseline = true)]
    public double OfficeIMO_Open_Read() {
        double checksum = 0;
        foreach (byte[] workbook in _workbooks) {
            using var document = ExcelDocument.Load(new MemoryStream(workbook, writable: false));
            foreach (ExcelCellValueInfo cell in document.Sheets[0].EnumerateCells()) {
                checksum += cell.Value switch { string text => text.Length, double number => number, _ => throw new InvalidDataException() };
            }
        }
        return checksum;
    }

    [Benchmark]
    public double ClosedXML_Open_Read() {
        double checksum = 0;
        foreach (byte[] workbook in _workbooks) {
            using var document = new XLWorkbook(new MemoryStream(workbook, writable: false));
            foreach (IXLCell cell in document.Worksheet(1).CellsUsed()) {
                checksum += cell.Value.IsText ? cell.Value.GetText().Length : cell.Value.GetNumber();
            }
        }
        return checksum;
    }

    [Benchmark]
    public double OpenXml_Open_Read() {
        double checksum = 0;
        foreach (byte[] workbook in _workbooks) {
            using var document = SpreadsheetDocument.Open(new MemoryStream(workbook, writable: false), false);
            WorkbookPart part = document.WorkbookPart!;
            string[] shared = part.SharedStringTablePart?.SharedStringTable!.Elements<SharedStringItem>().Select(item => item.InnerText).ToArray() ?? [];
            foreach (Cell cell in part.WorksheetParts.Single().Worksheet!.Descendants<Cell>()) {
                string text = cell.CellValue?.Text ?? cell.InlineString?.InnerText ?? string.Empty;
                checksum += cell.DataType?.Value == CellValues.SharedString ? shared[int.Parse(text, CultureInfo.InvariantCulture)].Length
                    : cell.DataType?.Value == CellValues.InlineString || cell.DataType?.Value == CellValues.String ? text.Length
                    : double.Parse(text, CultureInfo.InvariantCulture);
            }
        }
        return checksum;
    }
}
