using System.Data;
using System.Globalization;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

public enum WorksheetPreparationDimension {
    Correct,
    Absent,
    StaleNarrow
}

/// <summary>Measures public XLSX preparation and traversal with a later wider row.</summary>
[MemoryDiagnoser]
public class ExcelWorksheetPreparationBenchmarks {
    private byte[] _workbook = [];
    private long _expectedHeaderChecksum;
    private long _expectedFullChecksum;

    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 25_000;

    [Params(WorksheetPreparationDimension.Correct, WorksheetPreparationDimension.Absent,
        WorksheetPreparationDimension.StaleNarrow)]
    public WorksheetPreparationDimension Dimension { get; set; }

    [Params(false, true)]
    public bool StoredWorksheet { get; set; }

    [Params(false, true)]
    public bool EnableWorksheetPrefetch { get; set; }

    public IEnumerable<int> RowCounts() {
        string value = Environment.GetEnvironmentVariable("OFFICEIMO_TYPED_BENCHMARK_ROWS") ?? "25000";
        if (!int.TryParse(value.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out int count)
            || count is < 1 or > 1_000_000) {
            throw new ArgumentException("Worksheet preparation requires one OFFICEIMO_TYPED_BENCHMARK_ROWS count between 1 and 1000000.");
        }
        yield return count;
    }

    [GlobalSetup]
    public void Setup() {
        if (RowCount is < 1 or > 1_000_000) throw new ArgumentOutOfRangeException(nameof(RowCount));
        string? priority = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_PROCESS_PRIORITY");
        if (!string.IsNullOrEmpty(priority)) BenchmarkProcessorAffinity.ApplyPriority(priority);
        try {
            _workbook = ExcelWorksheetPreparationFixture.Create(RowCount, Dimension, StoredWorksheet);
            ExcelWorksheetPreparationFixture.ValidatePackage(_workbook, RowCount, Dimension, StoredWorksheet);
            _expectedHeaderChecksum = ExpectedChecksum(0);
            _expectedFullChecksum = ExpectedChecksum(RowCount);
            ValidateCompleteFields();
            OpenThroughFirstRow();
            FullScan();
            Console.WriteLine($"Qualified worksheet preparation: dataRows={RowCount}; dimension={Dimension}; "
                + $"storedWorksheet={StoredWorksheet}; prefetch={EnableWorksheetPrefetch}; fields=3; "
                + $"headerChecksum={_expectedHeaderChecksum}; fullChecksum={_expectedFullChecksum}.");
        } catch {
            Cleanup();
            throw;
        }
    }

    /// <summary>Checks every public field, scalar type, null, row and schema outside measurement.</summary>
    public void ValidateCompleteFields() {
        using ExcelWorkbookDataReader reader = OpenReader();
        ValidateSchema(reader);
        int row = 0;
        long checksum = 0;
        while (reader.Read()) {
            if (row > RowCount) throw new InvalidDataException("Preparation fixture returned an extra row.");
            for (int column = 0; column < 3; column++) {
                object expected = ExpectedValue(row, column);
                object actual = reader.GetValue(column);
                if (actual.GetType() != expected.GetType() || !Equals(actual, expected)
                    || reader.IsDBNull(column) != ReferenceEquals(expected, DBNull.Value)) {
                    throw new InvalidDataException($"Preparation field differs at row {row + 1}, column {column + 1}.");
                }
                if (expected is double number && reader.GetDouble(column) != number
                    || expected is string text && reader.GetString(column) != text) {
                    throw new InvalidDataException($"Preparation typed getter differs at row {row + 1}, column {column + 1}.");
                }
                checksum = AddValue(checksum, actual);
            }
            row++;
        }
        ValidateSchema(reader);
        if (row != RowCount + 1 || checksum != _expectedFullChecksum || reader.NextResult())
            throw new InvalidDataException("Preparation fixture row order/count, checksum or result count differs.");
    }

    [Benchmark]
    [BenchmarkCategory("WorksheetPreparationFirstRow")]
    public long OpenThroughFirstRow() {
        using ExcelWorkbookDataReader reader = OpenReader();
        if (reader.FieldCount != 3 || !reader.Read())
            throw new InvalidDataException("Preparation first row is missing or has the wrong width.");
        long checksum = AccumulateRow(reader, 0);
        return checksum == _expectedHeaderChecksum ? checksum
            : throw new InvalidDataException("Preparation first-row checksum differs.");
    }

    [Benchmark]
    [BenchmarkCategory("WorksheetPreparationFullScan")]
    public long FullScan() {
        using ExcelWorkbookDataReader reader = OpenReader();
        if (reader.FieldCount != 3) throw new InvalidDataException("Preparation width differs.");
        int count = 0;
        long checksum = 0;
        while (reader.Read()) {
            checksum = AccumulateRow(reader, checksum);
            count++;
        }
        return count == RowCount + 1 && reader.FieldCount == 3
            && checksum == _expectedFullChecksum && !reader.NextResult() ? checksum
            : throw new InvalidDataException("Preparation full-scan count or checksum differs.");
    }

    [GlobalCleanup]
    public void Cleanup() => _workbook = [];

    private ExcelWorkbookDataReader OpenReader() => ExcelDocument.OpenDataReader(_workbook,
        new ExcelReadOptions {
            HasHeaderRow = false,
            NumericAsDecimal = false,
            EnableWorksheetPrefetch = EnableWorksheetPrefetch
        });

    private object ExpectedValue(int row, int column) => row == 0
        ? column switch { 0 => "Id", 1 => "Value", _ => DBNull.Value }
        : column switch {
            0 => (double)row,
            1 => row == RowCount ? DBNull.Value : ExcelWorksheetPreparationFixture.Value(row),
            _ => row == RowCount ? ExcelWorksheetPreparationFixture.Value(row) : DBNull.Value
        };

    private long ExpectedChecksum(int dataRows) {
        long checksum = 0;
        for (int row = 0; row <= dataRows; row++)
            for (int column = 0; column < 3; column++)
                checksum = AddValue(checksum, ExpectedValue(row, column));
        return checksum;
    }

    private static long AccumulateRow(ExcelWorkbookDataReader reader, long checksum) {
        for (int column = 0; column < 3; column++) checksum = AddValue(checksum, reader.GetValue(column));
        return checksum;
    }

    private static long AddValue(long checksum, object value) {
        if (value is double number) return unchecked(checksum * 31 + BitConverter.DoubleToInt64Bits(number));
        if (ReferenceEquals(value, DBNull.Value)) return unchecked(checksum * 31 - 1);
        if (value is string text) {
            checksum = unchecked(checksum * 31 + text.Length);
            foreach (char character in text) checksum = unchecked(checksum * 31 + character);
            return checksum;
        }
        throw new InvalidDataException("Preparation raw field has an unexpected scalar type.");
    }

    private static void ValidateSchema(ExcelWorkbookDataReader reader) {
        if (reader.FieldCount != 3 || reader.SheetNames.Count != 1 || reader.CurrentSheetName != "Data"
            || reader.CurrentSheetIndex != 0 || reader.CurrentResultIndex != 0) {
            throw new InvalidDataException("Preparation workbook selection or width differs.");
        }
        using DataTable schema = reader.GetSchemaTable() ?? throw new InvalidDataException("Preparation raw schema is missing.");
        if (schema.Rows.Count != 3) throw new InvalidDataException("Preparation schema width differs.");
        for (int column = 0; column < 3; column++) {
            string name = "Column" + (column + 1).ToString(CultureInfo.InvariantCulture);
            DataRow field = schema.Rows[column];
            if (reader.GetName(column) != name || reader.GetOrdinal(name) != column || reader.GetFieldType(column) != typeof(object)
                || (string)field[SchemaTableColumn.ColumnName] != name || (int)field[SchemaTableColumn.ColumnOrdinal] != column
                || (Type)field[SchemaTableColumn.DataType] != typeof(object)) {
                throw new InvalidDataException($"Preparation schema differs at column {column + 1}.");
            }
        }
    }
}
