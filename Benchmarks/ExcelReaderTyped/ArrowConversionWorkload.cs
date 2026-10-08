#if OFFICEIMO_BENCHMARK_ARROW && OFFICEIMO_BENCHMARK_NEW_APIS
using System.Data.Common;
using System.Globalization;
using System.Text;
using Apache.Arrow;
using Apache.Arrow.Types;
using ExcelReader.Arrow;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Schema;
using OfficeIMO.CSV;
using OfficeIMO.Data.Arrow;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

public enum ArrowConversionScenario { CsvAllString, CsvTyped, XlsbTyped }

/// <summary>
/// Shared inputs and ordinary conversion routes for the pinned upstream
/// Interop/ArrowConversionBenchmark.cs workload. Inputs are never rewritten per engine.
/// </summary>
// The workload follows ExcelReader commit ca5b50f99e8ef57ab476f0a2bc8043558d58b28d:
// tests/ExcelReader.Benchmarks/Interop/ArrowConversionBenchmark.cs and
// tests/ExcelReader.Benchmarks/Shared/{CsvGenerator,WorkbookGenerator}.cs.
// The upstream MIT attribution is retained in THIRD-PARTY-NOTICES.md.
internal sealed class ArrowConversionWorkload {
    internal static readonly string[] Pool = ["alpha", "beta", "gamma", "delta", "epsilon", "zeta", "eta", "theta"];
    internal static readonly string[] WideHeaders = Enumerable.Range(0, 8).Select(i => $"c{i}").ToArray();
    private static readonly Type[] TypedTypes = [typeof(string), typeof(long), typeof(DateTime), typeof(double)];
    private static readonly Type[] StringTypes = Enumerable.Repeat(typeof(string), 8).ToArray();
    private static readonly ExcelColumnSchema[] TypedSchema = [
        new() { Index = 0, Name = "Name", Type = ExcelColumnType.StringColumn },
        new() { Index = 1, Name = "Id", Type = ExcelColumnType.Int64Column },
        new() { Index = 2, Name = "Date", Type = ExcelColumnType.TimestampColumn },
        new() { Index = 3, Name = "Value", Type = ExcelColumnType.Float64Column }
    ];
    private static readonly ExcelColumnSchema[] WideSchema = Enumerable.Range(0, 8)
        .Select(i => new ExcelColumnSchema { Index = i, Name = WideHeaders[i], Type = ExcelColumnType.StringColumn }).ToArray();
    private readonly byte[] _csvWide;
    private readonly byte[] _csvTyped;
    private readonly byte[] _xlsbTyped;

    private ArrowConversionWorkload(int rows, byte[] csvWide, byte[] csvTyped, byte[] xlsbTyped) {
        Rows = rows;
        _csvWide = csvWide;
        _csvTyped = csvTyped;
        _xlsbTyped = xlsbTyped;
    }

    internal int Rows { get; }

    internal static async Task<ArrowConversionWorkload> CreateAsync(int rows) {
        var wide = new StringBuilder(rows * 8 * 6);
        for (int row = 0; row < rows; row++) {
            for (int column = 0; column < 8; column++) {
                if (column > 0) wide.Append(',');
                wide.Append(Pool[(row + column) % Pool.Length]);
            }
            wide.Append('\n');
        }
        var typed = new StringBuilder(rows * 40);
        typed.Append("Name,Id,Date,Value\n");
        for (int row = 1; row <= rows; row++) {
            TypedRecord record = TypedWorkbookFixture.ExpectedRecord(row);
            typed.Append(record.Name).Append(',')
                .Append(record.Id.ToString(CultureInfo.InvariantCulture)).Append(',')
                .Append(record.Date.ToString("O", CultureInfo.InvariantCulture)).Append(',')
                .Append(record.Value.ToString(CultureInfo.InvariantCulture)).Append('\n');
        }
        byte[] csvWide = Encoding.UTF8.GetBytes(wide.ToString());
        byte[] csvTyped = Encoding.UTF8.GetBytes(typed.ToString());
        BenchmarkInput.WriteFixtureIdentity($"read/arrow/CsvAllString/rows={rows}/columns=8", csvWide);
        BenchmarkInput.WriteFixtureIdentity($"read/arrow/CsvTyped/dataRows={rows}/columns=4", csvTyped);
        return new ArrowConversionWorkload(rows, csvWide, csvTyped, await TypedWorkbookFixture.CreateXlsbAsync(rows));
    }

    internal long ConvertPeer(ArrowConversionScenario scenario, Action<RecordBatch, int>? validate = null) {
        if (scenario == ArrowConversionScenario.XlsbTyped) {
            using var source = new MemoryStream(_xlsbTyped, writable: false);
            using IExcelWorkbook reader = ExcelReaderApi.FromXlsb(source);
            using RecordBatch batch = reader.FirstSheet.ToArrowRecordBatch(TypedSchema);
            validate?.Invoke(batch, 0);
            return batch.Length;
        }
        using IExcelWorkbook csv = ExcelReaderApi.FromCsv(
            scenario == ArrowConversionScenario.CsvAllString ? _csvWide : _csvTyped);
        using RecordBatch converted = csv.FirstSheet.ToArrowRecordBatch(
            scenario == ArrowConversionScenario.CsvAllString ? WideSchema : TypedSchema,
            headerRow: scenario == ArrowConversionScenario.CsvAllString ? 0 : 1);
        validate?.Invoke(converted, 0);
        return converted.Length;
    }

    internal long ConvertOfficeIMO(
        ArrowConversionScenario scenario, int batchSize, Action<RecordBatch, int>? validate = null) {
        using var source = new MemoryStream(scenario switch {
            ArrowConversionScenario.CsvAllString => _csvWide,
            ArrowConversionScenario.CsvTyped => _csvTyped,
            _ => _xlsbTyped
        }, writable: false);
        using DbDataReader reader = scenario == ArrowConversionScenario.XlsbTyped
            ? ExcelDocument.OpenDataReader(source, new ExcelReadOptions { NumericAsDecimal = false })
            : CsvDocument.OpenDataReader(source,
                scenario == ArrowConversionScenario.CsvAllString
                    ? new CsvLoadOptions { HasHeaderRow = false, Header = WideHeaders }
                    : new CsvLoadOptions());
        var options = new ArrowReadOptions {
            BatchSize = batchSize,
            ColumnTypes = scenario == ArrowConversionScenario.CsvAllString ? StringTypes : TypedTypes,
            ColumnNullability = new bool[scenario == ArrowConversionScenario.CsvAllString ? 8 : 4],
            TemporalUnit = TimeUnit.Microsecond
        };
        return ConsumeOfficeIMO(reader, options, validate);
    }

    internal long ConvertPeerInferred(Action<RecordBatch, int>? validate = null) {
        using IExcelWorkbook reader = ExcelReaderApi.FromCsv(_csvTyped);
        using RecordBatch batch = reader.FirstSheet.ToArrowRecordBatch();
        validate?.Invoke(batch, 0);
        return batch.Length;
    }

    internal long ConvertOfficeIMOInferred(Action<RecordBatch, int>? validate = null) {
        using var source = new MemoryStream(_csvTyped, writable: false);
        using DbDataReader reader = CsvDocument.OpenDataReader(source, readerOptions:
            new CsvDataReaderOptions { InferSchema = true, SchemaSampleSize = 100 });
        return ConsumeOfficeIMO(reader, new ArrowReadOptions {
            BatchSize = Rows,
            ColumnNullability = new bool[4],
            TemporalUnit = TimeUnit.Microsecond
        }, validate);
    }

    private long ConsumeOfficeIMO(DbDataReader reader, ArrowReadOptions options, Action<RecordBatch, int>? validate) {
        int rows = 0, batches = 0;
        foreach (RecordBatch batch in reader.ReadArrowBatches(options)) {
            using (batch) {
                validate?.Invoke(batch, rows);
                rows += batch.Length;
                batches++;
            }
        }
        if (validate != null && (rows != Rows || batches != (Rows + options.BatchSize - 1) / options.BatchSize))
            throw new InvalidDataException($"Arrow returned {rows} rows in {batches} batches.");
        return rows;
    }
}
#endif
