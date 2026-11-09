using System.Globalization;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static OfficeIMO.Excel.Benchmarks.XlBenchShapeData;

namespace OfficeIMO.Excel.Benchmarks;

internal static partial class XlBenchShapeScenarios {
    private const int ReadColumns = 15;
    private static long? _expectedReadAllChecksum;

    internal static long OfficeImoReadAll() => OfficeImoReadAllCore(null);

    private static long OfficeImoReadAllCore(Action<int, int, object?>? validate) {
        using var document = ExcelDocument.Load(new MemoryStream(ReadWorkbook));
        long checksum = 0;
        int cells = 0;
        foreach (var cell in document.Sheets[0].EnumerateCells()) {
            validate?.Invoke(cell.Row, cell.Column, cell.Value);
            checksum += ReadAllTextLength(cell.Value);
            cells++;
        }
        return CompleteReadAll(cells, checksum);
    }

    internal static long OpenXmlReadAll() => OpenXmlReadAllCore(null);

    private static long OpenXmlReadAllCore(Action<int, int, object?>? validate) {
        using var document = SpreadsheetDocument.Open(new MemoryStream(ReadWorkbook), false);
        var workbookPart = document.WorkbookPart!;
        string[] shared = SharedStrings(workbookPart);
        // BuildRead uses the standard short-date format for its three date columns.
        bool[] dates = workbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!
            .Elements<CellFormat>().Select(static format => format.NumberFormatId?.Value == 14U).ToArray();
        long checksum = 0;
        int cells = 0;
        using var reader = OpenXmlReader.Create(workbookPart.WorksheetParts.First());
        while (reader.Read()) {
            if (reader.ElementType != typeof(Cell) || !reader.IsStartElement) continue;
            var cell = (Cell)reader.LoadCurrentElement()!;
            string raw = cell.CellValue?.InnerText ?? cell.InnerText;
            object? value;
            if (cell.DataType?.Value == CellValues.SharedString) {
                value = shared[int.Parse(raw, CultureInfo.InvariantCulture)];
            } else if (cell.DataType?.Value == CellValues.InlineString || cell.CellValue == null) {
                value = raw;
            } else if (raw.Length == 0) {
                value = null;
            } else {
                double number = double.Parse(raw, CultureInfo.InvariantCulture);
                value = dates[checked((int)(cell.StyleIndex?.Value ?? 0))]
                    ? DateTime.FromOADate(number) : number;
            }
            if (validate != null) {
                string reference = cell.CellReference?.Value
                    ?? throw new InvalidDataException("The fixed read-all fixture is missing a cell reference.");
                int column = reference[0] - 'A' + 1;
                int row = int.Parse(reference.AsSpan(1), CultureInfo.InvariantCulture);
                validate(row, column, value);
            }
            checksum += ReadAllTextLength(value);
            cells++;
        }
        return CompleteReadAll(cells, checksum);
    }

    private static int ReadAllTextLength(object? value) => value is IFormattable formattable
        ? formattable.ToString(null, CultureInfo.InvariantCulture)?.Length ?? 0
        : value?.ToString()?.Length ?? 0;

    private static long CompleteReadAll(int cells, long checksum) {
        if (cells != ReadRows * ReadColumns ||
            (_expectedReadAllChecksum is long expected && checksum != expected)) {
            throw new InvalidDataException($"Read-all returned {cells} cells and checksum {checksum}.");
        }
        return checksum;
    }

    /// <summary>Checks every logical value and its position through a separate reader before measurement.</summary>
    internal static void ValidateReadAll() {
        var expected = new object?[ReadRows * ReadColumns];
        long expectedChecksum = 0;
        using (var workbook = new XLWorkbook(new MemoryStream(ReadWorkbook))) {
            var sheet = workbook.Worksheet("Data");
            if (sheet.LastRowUsed()!.RowNumber() != ReadRows ||
                sheet.LastColumnUsed()!.ColumnNumber() != ReadColumns) {
                throw new InvalidDataException("The read-all fixture dimensions do not match the declared scenario.");
            }
            for (int row = 1; row <= ReadRows; row++) {
                for (int column = 1; column <= ReadColumns; column++) {
                    var cell = sheet.Cell(row, column);
                    var cached = cell.CachedValue;
                    object? value = cached.Type switch {
                        XLDataType.Blank when cell.HasFormula => cell.FormulaA1,
                        XLDataType.Blank => null,
                        XLDataType.Text => cached.GetText(),
                        XLDataType.Number => cached.GetNumber(),
                        XLDataType.DateTime => cached.GetDateTime(),
                        _ => throw new InvalidDataException($"Unsupported read-all fixture type at R{row}C{column}: {cached.Type}.")
                    };
                    expected[(row - 1) * ReadColumns + column - 1] = value;
                    expectedChecksum += ReadAllTextLength(value);
                }
            }
        }

        foreach (bool officeImo in new[] { true, false }) {
            int index = 0;
            void Check(int row, int column, object? value) {
                if (index >= expected.Length || row != index / ReadColumns + 1 ||
                    column != index % ReadColumns + 1) {
                    throw new InvalidDataException($"Read-all returned an out-of-order cell at R{row}C{column}.");
                }
                object? target = expected[index++];
                bool equal = target is double number &&
                    value is sbyte or byte or short or ushort or int or uint or long or ulong or float or double or decimal
                    ? Convert.ToDouble(value, CultureInfo.InvariantCulture) == number
                    : Equals(target, value);
                if (!equal) {
                    throw new InvalidDataException($"Read-all R{row}C{column}: expected {target} ({target?.GetType().Name}), found {value} ({value?.GetType().Name}).");
                }
            }
            long actual = officeImo ? OfficeImoReadAllCore(Check) : OpenXmlReadAllCore(Check);
            if (index != expected.Length || actual != expectedChecksum) {
                throw new InvalidDataException("Read-all did not materialize every expected value.");
            }
        }
        _expectedReadAllChecksum = expectedChecksum;
        Console.WriteLine($"VALIDATED_READ_ALL cells={expected.Length} checksum={expectedChecksum} oracle=ClosedXML invariant-values");
    }
}
