using System.Data;
using System.Globalization;
using ExcelDataReader;
using SalesRecord = OfficeIMO.Excel.Benchmarks.ExcelBenchmarkScenarioFactory.SalesRecord;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Validates complete sales-fixture values outside measured operations.</summary>
internal static class ExcelSalesOutputValidator {
    static ExcelSalesOutputValidator() => System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

    internal static void ValidateDictionaries(IReadOnlyList<SalesRecord> expected, IReadOnlyList<Dictionary<string, object?>> actual) {
        if (actual.Count != expected.Count) throw new InvalidDataException("Incorrect materialized row count.");
        for (int row = 0; row < expected.Count; row++) {
            if (actual[row].Count != ExcelBenchmarkScenarioFactory.SalesColumns.Length) {
                throw new InvalidDataException("Incorrect materialized column count.");
            }
            for (int column = 0; column < ExcelBenchmarkScenarioFactory.SalesColumns.Length; column++) {
                var definition = ExcelBenchmarkScenarioFactory.SalesColumns[column];
                AssertValue(definition.Selector(expected[row]), actual[row][definition.Header], row, column);
            }
        }
    }

    internal static void ValidateDataTable(IReadOnlyList<SalesRecord> expected, DataTable actual) {
        if (actual.Rows.Count != expected.Count || actual.Columns.Count != ExcelBenchmarkScenarioFactory.SalesColumns.Length) {
            throw new InvalidDataException("Incorrect DataTable dimensions.");
        }
        for (int column = 0; column < actual.Columns.Count; column++) {
            if (actual.Columns[column].ColumnName != ExcelBenchmarkScenarioFactory.SalesColumnNames[column]
                || actual.Columns[column].DataType != typeof(object)) {
                throw new InvalidDataException("Incorrect DataTable schema.");
            }
            for (int row = 0; row < expected.Count; row++) {
                AssertValue(ExcelBenchmarkScenarioFactory.SalesColumns[column].Selector(expected[row]), actual.Rows[row][column], row, column);
            }
        }
    }

    internal static void ValidateWorkbook(byte[] bytes, IReadOnlyList<SalesRecord> expected, string sheetName = "Data", bool reviewed = false) {
        using var input = new MemoryStream(bytes, writable: false);
        using var reader = ExcelReaderFactory.CreateReader(input);
        int columns = ExcelBenchmarkScenarioFactory.SalesColumns.Length + (reviewed ? 1 : 0);
        if (reader.Name != sheetName || reader.FieldCount != columns || !reader.Read()) {
            throw new InvalidDataException("Incorrect workbook sheet or dimensions.");
        }
        for (int column = 0; column < columns; column++) {
            string header = column == 8 ? "ReviewStatus" : ExcelBenchmarkScenarioFactory.SalesColumnNames[column];
            AssertValue(header, reader.GetValue(column), -1, column);
        }
        for (int row = 0; row < expected.Count; row++) {
            if (!reader.Read()) throw new InvalidDataException($"Workbook stopped at row {row}.");
            for (int column = 0; column < 8; column++) {
                AssertValue(ExcelBenchmarkScenarioFactory.SalesColumns[column].Selector(expected[row]), reader.GetValue(column), row, column);
            }
            if (reviewed) AssertValue(row < 100 ? "Reviewed" : null, reader.GetValue(8), row, 8);
        }
        if (reader.Read() || reader.NextResult()) throw new InvalidDataException("Workbook contains unexpected rows or sheets.");
    }

    private static void AssertValue(object? expected, object? actual, int row, int column) {
        bool valid = expected switch {
            null => actual is null or DBNull,
            int or double => actual is int or double && Convert.ToDouble(expected, CultureInfo.InvariantCulture) == Convert.ToDouble(actual, CultureInfo.InvariantCulture),
            _ => Equals(expected, actual)
        };
        if (!valid) throw new InvalidDataException($"Cell ({row + 2}, {column + 1}) contains '{actual}' ({actual?.GetType().Name}); expected '{expected}' ({expected?.GetType().Name}).");
    }
}
