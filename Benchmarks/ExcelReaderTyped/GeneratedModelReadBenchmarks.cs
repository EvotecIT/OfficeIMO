#if OFFICEIMO_BENCHMARK_GENERATED_MAPPING
using BenchmarkDotNet.Attributes;
using OfficeIMO.Data;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>Measures existing automatic, explicit and generated ordinary-model mapping routes.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("TypedMappingStrategies")]
public class GeneratedModelReadBenchmarks {
    private byte[] _bytes = [];
    private long _expected;

    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;
    [ParamsAllValues]
    public TypedModelKind Model { get; set; }
    public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();

    [GlobalSetup]
    public async Task SetupAsync() {
        BenchmarkInput.WriteDescription();
        _bytes = await TypedWorkbookFixture.CreateAsync(RowCount, "Original");
        _expected = TypedWorkbookFixture.ExpectedChecksum(RowCount);
        TypedWorkbookFixture.ValidateShape(_bytes, RowCount, "Original");
        if (Model == TypedModelKind.Class) {
            ValidateClass(null);
            ValidateClass(ConfigureClass);
            ValidateClass(TypedRecordRowMapping.Configure);
        } else {
            ValidateStruct(null);
            ValidateStruct(ConfigureStruct);
            ValidateStruct(TypedValueRecordRowMapping.Configure);
        }
        OfficeIMOAutomatic();
        OfficeIMOExplicit();
        OfficeIMOGenerated();
        Console.WriteLine($"Validated automatic, explicit and generated {Model} mapping: "
            + $"rows={RowCount}, all four fields, headers, row order, checksum={_expected}.");
    }

    [Benchmark(Baseline = true)]
    public long OfficeIMOAutomatic() {
        using var reader = OpenReader();
        return Model == TypedModelKind.Class
            ? ScanClass(reader.RowsAs<TypedRecord>())
            : ScanStruct(reader.RowsAs<TypedValueRecord>());
    }

    [Benchmark]
    public long OfficeIMOExplicit() {
        using var reader = OpenReader();
        return Model == TypedModelKind.Class
            ? ScanClass(reader.RowsAs<TypedRecord>(ConfigureClass))
            : ScanStruct(reader.RowsAs<TypedValueRecord>(ConfigureStruct));
    }

    [Benchmark]
    public long OfficeIMOGenerated() {
        using var reader = OpenReader();
        return Model == TypedModelKind.Class
            ? ScanClass(reader.RowsAs<TypedRecord>(TypedRecordRowMapping.Configure))
            : ScanStruct(reader.RowsAs<TypedValueRecord>(TypedValueRecordRowMapping.Configure));
    }

    private ExcelWorkbookDataReader OpenReader() =>
        ExcelDocument.OpenDataReader(_bytes, new ExcelReadOptions { HasHeaderRow = true });

    private void ValidateClass(Action<RowMapper<TypedRecord>>? configure) {
        using var reader = OpenReader();
        ValidateHeaders(reader);
        int count = 0;
        long sum = 0;
        IEnumerable<TypedRecord> rows = configure == null ? reader.RowsAs<TypedRecord>() : reader.RowsAs(configure);
        foreach (TypedRecord row in rows) {
            TypedWorkbookFixture.ValidateRecord(row, ++count);
            sum = unchecked(sum + Accumulate(row.Name, row.Id, row.Date, row.Value));
        }
        Check(sum, count);
        if (reader.NextResult()) throw new InvalidDataException("Generated mapping input has more than one sheet.");
    }

    private void ValidateStruct(Action<RowMapper<TypedValueRecord>>? configure) {
        using var reader = OpenReader();
        ValidateHeaders(reader);
        int count = 0;
        long sum = 0;
        IEnumerable<TypedValueRecord> rows = configure == null ? reader.RowsAs<TypedValueRecord>() : reader.RowsAs(configure);
        foreach (TypedValueRecord row in rows) {
            TypedWorkbookFixture.ValidateRecord(new TypedRecord {
                Name = row.Name, Id = row.Id, Date = row.Date, Value = row.Value
            }, ++count);
            sum = unchecked(sum + Accumulate(row.Name, row.Id, row.Date, row.Value));
        }
        Check(sum, count);
        if (reader.NextResult()) throw new InvalidDataException("Generated mapping input has more than one sheet.");
    }

    private static void ValidateHeaders(ExcelWorkbookDataReader reader) {
        if (reader.FieldCount != TypedWorkbookFixture.Headers.Length
            || !Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).SequenceEqual(TypedWorkbookFixture.Headers))
            throw new InvalidDataException("Generated mapping input headers differ from the four-property model.");
    }

    private long ScanClass(IEnumerable<TypedRecord> rows) {
        long sum = 0;
        int count = 0;
        foreach (TypedRecord row in rows) {
            sum = unchecked(sum + Accumulate(row.Name, row.Id, row.Date, row.Value));
            count++;
        }
        return Check(sum, count);
    }

    private long ScanStruct(IEnumerable<TypedValueRecord> rows) {
        long sum = 0;
        int count = 0;
        foreach (TypedValueRecord row in rows) {
            sum = unchecked(sum + Accumulate(row.Name, row.Id, row.Date, row.Value));
            count++;
        }
        return Check(sum, count);
    }

    private static void ConfigureClass(RowMapper<TypedRecord> mapper) {
        mapper.FromColumn<string?>("Name", static (row, value) => { row.Name = value; return row; });
        mapper.FromColumn<int>("Id", static (row, value) => { row.Id = value; return row; });
        mapper.FromColumn<DateTime>("Date", static (row, value) => { row.Date = value; return row; });
        mapper.FromColumn<double>("Value", static (row, value) => { row.Value = value; return row; });
    }

    private static void ConfigureStruct(RowMapper<TypedValueRecord> mapper) {
        mapper.FromColumn<string?>("Name", static (row, value) => { row.Name = value; return row; });
        mapper.FromColumn<int>("Id", static (row, value) => { row.Id = value; return row; });
        mapper.FromColumn<DateTime>("Date", static (row, value) => { row.Date = value; return row; });
        mapper.FromColumn<double>("Value", static (row, value) => { row.Value = value; return row; });
    }

    private static long Accumulate(string? name, int id, DateTime date, double value) =>
        id + (long)value + (name?.Length ?? 0) + date.Ticks;
    private long Check(long sum, int count) => sum == _expected && count == RowCount
        ? sum : throw new InvalidDataException("Ordinary model mapping returned an incorrect count or checksum.");
}
#endif
