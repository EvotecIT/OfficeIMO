#nullable enable

using System.Data;
using System.Globalization;
using BenchmarkDotNet.Attributes;
using ExcelDataReader;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Compares the public automatic, sequential, and parallel DataTable write
/// requests over the same prepared 25,000-row table. Every method creates and
/// saves a complete XLSX package; setup reopens and validates every header and
/// cell before timing.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class ExcelDataTableExecutionBenchmarks {
    private DataTable _table = null!;

    [Params(25000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(RowCount);
        _table = ExcelLibraryComparisonRunner.CreateSalesDataTable(rows, "Data");

        ValidateOutput(nameof(Automatic), WritePackage(null));
        ValidateOutput(nameof(Sequential), WritePackage(ExcelExecutionMode.Sequential));
        ValidateOutput(nameof(Parallel), WritePackage(ExcelExecutionMode.Parallel));
    }

    [Benchmark(Baseline = true)]
    public int Automatic()
        => WritePackage(null).Length;

    [Benchmark]
    public int Sequential()
        => WritePackage(ExcelExecutionMode.Sequential).Length;

    [Benchmark]
    public int Parallel()
        => WritePackage(ExcelExecutionMode.Parallel).Length;

    private byte[] WritePackage(ExcelExecutionMode? mode) {
        using var stream = new MemoryStream();
        using (var document = ExcelDocument.Create(stream)) {
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.InsertDataTable(_table, mode: mode);
            document.Save(stream);
        }

        return stream.ToArray();
    }

    private void ValidateOutput(string method, byte[] packageBytes) {
        using var stream = new MemoryStream(packageBytes, writable: false);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        if (!reader.Read()) {
            throw new InvalidOperationException($"{method} did not write a header row.");
        }

        for (int column = 0; column < _table.Columns.Count; column++) {
            string expectedHeader = _table.Columns[column].ColumnName;
            string? actualHeader = Convert.ToString(reader.GetValue(column), CultureInfo.InvariantCulture);
            if (!string.Equals(actualHeader, expectedHeader, StringComparison.Ordinal)) {
                throw new InvalidOperationException(
                    $"{method} header {column} was '{actualHeader}'; expected '{expectedHeader}'.");
            }
        }

        for (int row = 0; row < _table.Rows.Count; row++) {
            if (!reader.Read()) {
                throw new InvalidOperationException(
                    $"{method} stopped after {row:N0} data rows; expected {_table.Rows.Count:N0}.");
            }

            for (int column = 0; column < _table.Columns.Count; column++) {
                AssertEquivalent(method, row, column, _table.Rows[row][column], reader.GetValue(column));
            }
        }

        if (reader.Read()) {
            throw new InvalidOperationException($"{method} wrote more than {_table.Rows.Count:N0} data rows.");
        }
    }

    private static void AssertEquivalent(string method, int row, int column, object expected, object actual) {
        bool equivalent = expected switch {
            DateTime expectedDate => actual is DateTime actualDate && actualDate == expectedDate,
            double expectedNumber => Math.Abs(Convert.ToDouble(actual, CultureInfo.InvariantCulture) - expectedNumber) < 0.0000001,
            int expectedNumber => Convert.ToInt32(actual, CultureInfo.InvariantCulture) == expectedNumber,
            bool expectedBoolean => Convert.ToBoolean(actual, CultureInfo.InvariantCulture) == expectedBoolean,
            _ => string.Equals(
                Convert.ToString(actual, CultureInfo.InvariantCulture),
                Convert.ToString(expected, CultureInfo.InvariantCulture),
                StringComparison.Ordinal)
        };

        if (!equivalent) {
            throw new InvalidOperationException(
                $"{method} cell ({row + 2}, {column + 1}) was '{actual}'; expected '{expected}'.");
        }
    }
}
