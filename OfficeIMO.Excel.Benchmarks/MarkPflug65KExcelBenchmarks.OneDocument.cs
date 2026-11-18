#if NET10_0_OR_GREATER
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;
using RecursiveAI.OneDocument.Spreadsheets;
using RecursiveAI.OneDocument.Spreadsheets.Streaming;

namespace OfficeIMO.Excel.Benchmarks;

public partial class MarkPflug65KXlsxBenchmarks {
    /// <summary>Observes the same typed sales values through OneDocument's forward-only XLSX cursor.</summary>
    [Benchmark]
    [BenchmarkCategory(BenchmarkLibrarySelection.CategoryPrefix + "onedocument")]
    public ExcelReadObservation OneDocument() {
        using XlsxStreamReader workbook = XlsxStreamReader.Open(MarkPflug65KFixture.XlsxPath);
        using SheetCursor sheet = workbook.OpenSheet(0);
        return Validate(nameof(OneDocument), OneDocumentSalesObservation.Observe(sheet));
    }
}

public partial class MarkPflug65KXlsbBenchmarks {
    /// <summary>Observes the same typed sales values through OneDocument's forward-only XLSB cursor.</summary>
    [Benchmark]
    [BenchmarkCategory(BenchmarkLibrarySelection.CategoryPrefix + "onedocument")]
    public ExcelReadObservation OneDocument() {
        using XlsbStreamReader workbook = XlsbStreamReader.Open(MarkPflug65KFixture.XlsbPath);
        using XlsbSheetCursor sheet = workbook.OpenSheet(0);
        return Validate(nameof(OneDocument), OneDocumentSalesObservation.Observe(sheet));
    }
}

public partial class MarkPflug65KXlsBenchmarks {
    /// <summary>Observes the same typed sales values through OneDocument's forward-only BIFF cursor.</summary>
    [Benchmark]
    [BenchmarkCategory(BenchmarkLibrarySelection.CategoryPrefix + "onedocument")]
    public ExcelReadObservation OneDocument() {
        using XlsStreamReader workbook = XlsStreamReader.Open(MarkPflug65KFixture.XlsPath);
        using XlsSheetCursor sheet = workbook.OpenSheet(0);
        return Validate(nameof(OneDocument), OneDocumentSalesObservation.Observe(sheet));
    }
}

/// <summary>
/// Consumes the fourteen typed sales columns from each OneDocument sheet cursor. The
/// cursor types share no base class, so each has its own overload rather than an
/// interface-dispatched generic.
/// </summary>
internal static class OneDocumentSalesObservation {
    internal static ExcelReadObservation Observe(SheetCursor sheet) {
        var observation = new ExcelObservationAccumulator();
        bool header = true;
        while (sheet.ReadRow()) {
            if (header) {
                header = false;
                continue;
            }

            observation.BeginRow();
            for (int column = 1; column <= 5; column++) {
                MoveTo(sheet, column);
                observation.Add(sheet.GetText());
            }

            MoveTo(sheet, 6);
            observation.Add(ReadDate(sheet));
            MoveTo(sheet, 7);
            observation.Add(checked((int)ReadNumber(sheet)));
            MoveTo(sheet, 8);
            observation.Add(ReadDate(sheet));
            MoveTo(sheet, 9);
            observation.Add(checked((int)ReadNumber(sheet)));
            // All numeric-cell lanes observe the same IEEE-to-decimal values.
            for (int column = 10; column <= 14; column++) {
                MoveTo(sheet, column);
                observation.Add((decimal)ReadNumber(sheet));
            }
        }

        return observation.Build();
    }

    private static void MoveTo(SheetCursor sheet, int column) {
        if (!sheet.ReadCell() || sheet.Column != column) {
            throw new InvalidDataException($"OneDocument did not return column {column} in row {sheet.RowNumber}.");
        }
    }

    private static DateTime ReadDate(SheetCursor sheet) =>
        sheet.TryGetDateTime(out DateTime value)
            ? value
            : throw new InvalidDataException($"OneDocument did not produce a date at {sheet.Address}.");

    private static double ReadNumber(SheetCursor sheet) =>
        sheet.Kind == CellValueKind.Number
            ? sheet.Number
            : throw new InvalidDataException($"OneDocument did not produce a number at {sheet.Address}.");

    internal static ExcelReadObservation Observe(XlsbSheetCursor sheet) {
        var observation = new ExcelObservationAccumulator();
        bool header = true;
        while (sheet.ReadRow()) {
            if (header) {
                header = false;
                continue;
            }

            observation.BeginRow();
            for (int column = 1; column <= 5; column++) {
                MoveTo(sheet, column);
                observation.Add(sheet.GetText());
            }

            MoveTo(sheet, 6);
            observation.Add(ReadDate(sheet));
            MoveTo(sheet, 7);
            observation.Add(checked((int)ReadNumber(sheet)));
            MoveTo(sheet, 8);
            observation.Add(ReadDate(sheet));
            MoveTo(sheet, 9);
            observation.Add(checked((int)ReadNumber(sheet)));
            // All numeric-cell lanes observe the same IEEE-to-decimal values.
            for (int column = 10; column <= 14; column++) {
                MoveTo(sheet, column);
                observation.Add((decimal)ReadNumber(sheet));
            }
        }

        return observation.Build();
    }

    private static void MoveTo(XlsbSheetCursor sheet, int column) {
        if (!sheet.ReadCell() || sheet.Column != column) {
            throw new InvalidDataException($"OneDocument did not return column {column} in row {sheet.RowNumber}.");
        }
    }

    private static DateTime ReadDate(XlsbSheetCursor sheet) =>
        sheet.TryGetDateTime(out DateTime value)
            ? value
            : throw new InvalidDataException($"OneDocument did not produce a date at {sheet.Address}.");

    private static double ReadNumber(XlsbSheetCursor sheet) =>
        sheet.Kind == CellValueKind.Number
            ? sheet.Number
            : throw new InvalidDataException($"OneDocument did not produce a number at {sheet.Address}.");

    internal static ExcelReadObservation Observe(XlsSheetCursor sheet) {
        var observation = new ExcelObservationAccumulator();
        bool header = true;
        while (sheet.ReadRow()) {
            if (header) {
                header = false;
                continue;
            }

            observation.BeginRow();
            for (int column = 1; column <= 5; column++) {
                MoveTo(sheet, column);
                observation.Add(sheet.GetText());
            }

            MoveTo(sheet, 6);
            observation.Add(ReadDate(sheet));
            MoveTo(sheet, 7);
            observation.Add(checked((int)ReadNumber(sheet)));
            MoveTo(sheet, 8);
            observation.Add(ReadDate(sheet));
            MoveTo(sheet, 9);
            observation.Add(checked((int)ReadNumber(sheet)));
            // All numeric-cell lanes observe the same IEEE-to-decimal values.
            for (int column = 10; column <= 14; column++) {
                MoveTo(sheet, column);
                observation.Add((decimal)ReadNumber(sheet));
            }
        }

        return observation.Build();
    }

    private static void MoveTo(XlsSheetCursor sheet, int column) {
        if (!sheet.ReadCell() || sheet.Column != column) {
            throw new InvalidDataException($"OneDocument did not return column {column} in row {sheet.RowNumber}.");
        }
    }

    private static DateTime ReadDate(XlsSheetCursor sheet) =>
        sheet.TryGetDateTime(out DateTime value)
            ? value
            : throw new InvalidDataException($"OneDocument did not produce a date at {sheet.Address}.");

    private static double ReadNumber(XlsSheetCursor sheet) =>
        sheet.Kind == CellValueKind.Number
            ? sheet.Number
            : throw new InvalidDataException($"OneDocument did not produce a number at {sheet.Address}.");
}
#endif
