using System.Data;
using System.Globalization;
using BenchmarkDotNet.Attributes;
using ExcelDataReader;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Compares compact forward-only XLSX writers using the established 25,000-row
/// DataReader contract. Setup verifies every header and cell before timing.
/// </summary>
[MemoryDiagnoser]
public partial class ExcelDataReaderWriteBenchmarks {
    private DataTable _table = null!;

    [Params(25000)]
    public int RowCount { get; set; }

    public void Setup() {
        Initialize();

        ValidateOutput(nameof(OfficeIMO), ExcelLibraryComparisonRunner.OfficeImoWriteDataReaderCompactPackageBytes(_table));
        ValidateOutput(nameof(SpreadCheetah), ExcelLibraryComparisonRunner.SpreadCheetahWriteDataReaderPlainBytes(_table));
        ValidateOutput(nameof(Sylvan), ExcelLibraryComparisonRunner.SylvanWriteDataReaderPlainBytes(_table));
        ValidateOutput(nameof(LargeXlsx), ExcelLibraryComparisonRunner.LargeXlsxWriteDataReaderPlainBytes(_table, requireCellReferences: false));
#if NET10_0_OR_GREATER
        ValidateOutput(nameof(OneDocument), ExcelLibraryComparisonRunner.OneDocumentWriteDataReaderPlainBytes(_table));
#endif
    }

    public void SetupOfficeIMOAndSpreadCheetah() {
        Initialize();
        ValidateOutput(nameof(OfficeIMO), ExcelLibraryComparisonRunner.OfficeImoWriteDataReaderCompactPackageBytes(_table));
        ValidateOutput(nameof(SpreadCheetah), ExcelLibraryComparisonRunner.SpreadCheetahWriteDataReaderPlainBytes(_table));
    }

    [GlobalSetup(Target = nameof(OfficeIMO))]
    public void SetupOfficeIMO() {
        Initialize();
        ValidateOutput(nameof(OfficeIMO), ExcelLibraryComparisonRunner.OfficeImoWriteDataReaderCompactPackageBytes(_table));
    }

    [GlobalSetup(Target = nameof(SpreadCheetah))]
    public void SetupSpreadCheetah() {
        Initialize();
        ValidateOutput(nameof(SpreadCheetah), ExcelLibraryComparisonRunner.SpreadCheetahWriteDataReaderPlainBytes(_table));
    }

    [GlobalSetup(Target = nameof(Sylvan))]
    public void SetupSylvan() {
        Initialize();
        ValidateOutput(nameof(Sylvan), ExcelLibraryComparisonRunner.SylvanWriteDataReaderPlainBytes(_table));
    }

    [GlobalSetup(Target = nameof(LargeXlsx))]
    public void SetupLargeXlsx() {
        Initialize();
        ValidateOutput(nameof(LargeXlsx), ExcelLibraryComparisonRunner.LargeXlsxWriteDataReaderPlainBytes(_table, requireCellReferences: false));
    }

    private void Initialize() {
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) {
            BenchmarkProcessorAffinity.ApplyPriority(priority);
        }

        var rows = ExcelBenchmarkScenarioFactory.CreateSalesRecords(RowCount);
        _table = ExcelLibraryComparisonRunner.CreateSalesDataTable(rows, "Data");
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public int OfficeIMO()
        => ExcelLibraryComparisonRunner.OfficeImoWriteDataReaderCompactPackageBytes(_table).Length;

    [Benchmark]
    [BenchmarkCategory("library:SpreadCheetah", "case:Default")]
    public int SpreadCheetah()
        => ExcelLibraryComparisonRunner.SpreadCheetahWriteDataReaderPlainBytes(_table).Length;

    [Benchmark]
    [BenchmarkCategory("library:Sylvan.Data.Excel", "case:Default")]
    public int Sylvan()
        => ExcelLibraryComparisonRunner.SylvanWriteDataReaderPlainBytes(_table).Length;

    [Benchmark]
    [BenchmarkCategory("library:LargeXlsx", "case:Default")]
    public int LargeXlsx()
        => ExcelLibraryComparisonRunner.LargeXlsxWriteDataReaderPlainBytes(_table, requireCellReferences: false).Length;

    private void ValidateOutput(string method, byte[] packageBytes) {
        string library = method switch {
            nameof(OfficeIMO) => "OfficeIMO",
            nameof(SpreadCheetah) => "SpreadCheetah",
            nameof(Sylvan) => "Sylvan.Data.Excel",
            nameof(LargeXlsx) => "LargeXlsx",
            "OneDocument" => "OneDocument.Spreadsheets",
            _ => throw new ArgumentOutOfRangeException(nameof(method), method, "Unknown write lane.")
        };
        try {
            ValidateOutputCells(method, packageBytes);
        } catch (Exception exception) {
            BenchmarkOutputValidationLog.Failed(library, $"RowCount={RowCount}; {method}: {exception.Message}");
            throw;
        }

        BenchmarkOutputValidationLog.Passed(library,
            $"RowCount={RowCount}; {method}: ExcelDataReader read the header and all {_table.Rows.Count} x {_table.Columns.Count} cells back unchanged.");
    }

    private void ValidateOutputCells(string method, byte[] packageBytes) {
        using var stream = new MemoryStream(packageBytes, writable: false);
        using var reader = ExcelReaderFactory.CreateReader(stream);

        if (!reader.Read()) {
            throw new InvalidOperationException($"{method} did not write a header row.");
        }

        for (var column = 0; column < _table.Columns.Count; column++) {
            var expectedHeader = _table.Columns[column].ColumnName;
            var actualHeader = Convert.ToString(reader.GetValue(column), CultureInfo.InvariantCulture);
            if (!string.Equals(actualHeader, expectedHeader, StringComparison.Ordinal)) {
                throw new InvalidOperationException(
                    $"{method} header {column} was '{actualHeader}'; expected '{expectedHeader}'.");
            }
        }

        for (var row = 0; row < _table.Rows.Count; row++) {
            if (!reader.Read()) {
                throw new InvalidOperationException($"{method} stopped after {row:N0} data rows; expected {_table.Rows.Count:N0}.");
            }

            for (var column = 0; column < _table.Columns.Count; column++) {
                AssertEquivalent(method, row, column, _table.Rows[row][column], reader.GetValue(column));
            }
        }

        if (reader.Read()) {
            throw new InvalidOperationException($"{method} wrote more than {_table.Rows.Count:N0} data rows.");
        }
    }

    private static void AssertEquivalent(string method, int row, int column, object expected, object actual) {
        var equivalent = expected switch {
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
