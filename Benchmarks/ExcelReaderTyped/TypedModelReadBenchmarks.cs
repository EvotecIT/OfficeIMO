using System.Text;
using System.Data;
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using OfficeIMO.Data;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>The ordinary materialized model strategies in the upstream parser suite.</summary>
public enum TypedModelKind { Class, Struct }

/// <summary>The value-type counterpart of the upstream four-property record.</summary>
public struct TypedValueRecord {
    public string? Name { get; set; }
    public int Id { get; set; }
    public DateTime Date { get; set; }
    public double Value { get; set; }
}

/// <summary>Compares public automatic mapping to class and struct models.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("TypedAutomaticModel")]
public class TypedModelReadBenchmarks {
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
        using (var stream = new MemoryStream(_bytes, writable: false)) {
            using var workbook = ExcelReaderApi.FromXlsx(stream);
            int count = 0;
            if (Model == TypedModelKind.Class) {
                foreach (TypedRecord row in ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet))
                    Validate(row.Name, row.Id, row.Date, row.Value, ++count);
            } else {
                foreach (TypedValueRecord row in ExcelParser.FromAttributes<TypedValueRecord>().Parse(workbook.FirstSheet))
                    Validate(row.Name, row.Id, row.Date, row.Value, ++count);
            }
            Check(_expected, count);
        }
        using (var reader = ExcelDocument.OpenDataReader(_bytes, new ExcelReadOptions { HasHeaderRow = true })) {
            int count = 0;
            if (Model == TypedModelKind.Class) {
                foreach (TypedRecord row in reader.RowsAs<TypedRecord>())
                    Validate(row.Name, row.Id, row.Date, row.Value, ++count);
            } else {
                foreach (TypedValueRecord row in reader.RowsAs<TypedValueRecord>())
                    Validate(row.Name, row.Id, row.Date, row.Value, ++count);
            }
            Check(_expected, count);
        }
        ExcelReaderAutomatic();
        OfficeIMOAutomatic();
        Console.WriteLine($"Validated automatic {Model} mapping: rows={RowCount}, all four fields and row order, checksum={_expected}.");
    }

    [Benchmark(Baseline = true)]
    public long ExcelReaderAutomatic() {
        using var stream = new MemoryStream(_bytes, writable: false);
        using var workbook = ExcelReaderApi.FromXlsx(stream);
        long sum = 0;
        int count = 0;
        if (Model == TypedModelKind.Class) {
            foreach (TypedRecord row in ExcelParser.FromAttributes<TypedRecord>().Parse(workbook.FirstSheet)) {
                sum = unchecked(sum + Accumulate(row.Name, row.Id, row.Date, row.Value));
                count++;
            }
        } else {
            foreach (TypedValueRecord row in ExcelParser.FromAttributes<TypedValueRecord>().Parse(workbook.FirstSheet)) {
                sum = unchecked(sum + Accumulate(row.Name, row.Id, row.Date, row.Value));
                count++;
            }
        }
        return Check(sum, count);
    }

    [Benchmark]
    public long OfficeIMOAutomatic() {
        using var reader = ExcelDocument.OpenDataReader(_bytes, new ExcelReadOptions { HasHeaderRow = true });
        long sum = 0;
        int count = 0;
        if (Model == TypedModelKind.Class) {
            foreach (TypedRecord row in reader.RowsAs<TypedRecord>()) {
                sum = unchecked(sum + Accumulate(row.Name, row.Id, row.Date, row.Value));
                count++;
            }
        } else {
            foreach (TypedValueRecord row in reader.RowsAs<TypedValueRecord>()) {
                sum = unchecked(sum + Accumulate(row.Name, row.Id, row.Date, row.Value));
                count++;
            }
        }
        return Check(sum, count);
    }

    private static long Accumulate(string? name, int id, DateTime date, double value) =>
        id + (long)value + (name?.Length ?? 0) + date.Ticks;
    private static void Validate(string? name, int id, DateTime date, double value, int index) =>
        TypedWorkbookFixture.ValidateRecord(new TypedRecord { Name = name, Id = id, Date = date, Value = value }, index);
    private long Check(long sum, int count) => sum == _expected && count == RowCount
        ? sum : throw new InvalidDataException("Automatic model mapping returned an incorrect count or checksum.");
}

#if OFFICEIMO_BENCHMARK_NEW_APIS
/// <summary>A borrowed UTF-8 model that never escapes the current reader row.</summary>
public readonly ref struct BorrowedTypedRecord {
    public ReadOnlySpan<byte> Name { get; init; }
    public int Id { get; init; }
    public DateTime Date { get; init; }
    public double Value { get; init; }
}

/// <summary>Consumes equivalent ref-struct values with explicitly different mapping strategies.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("TypedBorrowedRefStruct")]
public class BorrowedRefStructReadBenchmarks {
    private byte[] _bytes = [];
    private long _expected;

    [ParamsSource(nameof(RowCounts))]
    public int RowCount { get; set; } = 50_000;
    public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();

    [GlobalSetup]
    public async Task SetupAsync() {
        BenchmarkInput.WriteDescription();
        _bytes = await TypedWorkbookFixture.CreateAsync(RowCount, "Original");
        _expected = TypedWorkbookFixture.ExpectedChecksum(RowCount);
        TypedWorkbookFixture.ValidateShape(_bytes, RowCount, "Original");
        using (var stream = new MemoryStream(_bytes, writable: false)) {
            using var workbook = ExcelReaderApi.FromXlsx(stream);
            int count = 0;
            foreach (BorrowedTypedRecord row in ExcelParser.FromAttributes<BorrowedTypedRecord>().Parse(workbook.FirstSheet))
                Validate(row, ++count);
            Check(_expected, count);
        }
        using (var reader = ExcelDocument.OpenDataReader(_bytes, new ExcelReadOptions { HasHeaderRow = true })) {
            int count = 0;
            while (reader.Read()) Validate(ReadOfficeIMO(reader), ++count);
            Check(_expected, count);
        }
        using (var reader = ExcelDocument.OpenDataReader(_bytes, new ExcelReadOptions { HasHeaderRow = true })) {
            int count = 0;
            foreach (BorrowedTypedRecord row in reader.RowsAsBorrowed<BorrowedTypedRecord>(ReadOfficeIMO))
                Validate(row, ++count);
            Check(_expected, count);
        }
        ExcelReaderMappedRefStruct();
        OfficeIMOManualRefStruct();
        OfficeIMOFactoryRefStruct();
        Console.WriteLine($"Validated borrowed ref-struct values: rows={RowCount}; peer automatic mapping, OfficeIMO factory mapping and manual public getters.");
    }

    [Benchmark(Baseline = true)]
    public long ExcelReaderMappedRefStruct() {
        using var stream = new MemoryStream(_bytes, writable: false);
        using var workbook = ExcelReaderApi.FromXlsx(stream);
        long sum = 0;
        int count = 0;
        foreach (BorrowedTypedRecord row in ExcelParser.FromAttributes<BorrowedTypedRecord>().Parse(workbook.FirstSheet)) {
            sum = unchecked(sum + Accumulate(row));
            count++;
        }
        return Check(sum, count);
    }

    [Benchmark]
    public long OfficeIMOManualRefStruct() {
        using var reader = ExcelDocument.OpenDataReader(_bytes, new ExcelReadOptions { HasHeaderRow = true });
        long sum = 0;
        int count = 0;
        while (reader.Read()) {
            sum = unchecked(sum + Accumulate(ReadOfficeIMO(reader)));
            count++;
        }
        return Check(sum, count);
    }

    [Benchmark]
    public long OfficeIMOFactoryRefStruct() {
        using var reader = ExcelDocument.OpenDataReader(_bytes, new ExcelReadOptions { HasHeaderRow = true });
        long sum = 0;
        int count = 0;
        foreach (BorrowedTypedRecord row in reader.RowsAsBorrowed<BorrowedTypedRecord>(ReadOfficeIMO)) {
            sum = unchecked(sum + Accumulate(row));
            count++;
        }
        return Check(sum, count);
    }

    private static BorrowedTypedRecord ReadOfficeIMO(IDataRecord reader) {
        if (!reader.TryGetUtf8Text(0, out ReadOnlySpan<byte> name))
            throw new InvalidDataException("The name field did not expose borrowed UTF-8 text.");
        return new BorrowedTypedRecord {
            Name = name, Id = reader.GetInt32(1), Date = reader.GetDateTime(2), Value = reader.GetDouble(3),
        };
    }

    private static void Validate(BorrowedTypedRecord row, int index) {
        TypedRecord expected = TypedWorkbookFixture.ExpectedRecord(index);
        if (!row.Name.SequenceEqual(Encoding.UTF8.GetBytes(expected.Name!))
            || row.Id != expected.Id || row.Date != expected.Date || row.Value != expected.Value)
            throw new InvalidDataException($"Borrowed fields differ at data row {index}.");
    }
    private static long Accumulate(BorrowedTypedRecord row) => row.Id + (long)row.Value + row.Name.Length + row.Date.Ticks;
    private long Check(long sum, int count) => sum == _expected && count == RowCount
        ? sum : throw new InvalidDataException("Ref-struct scan returned an incorrect count or checksum.");
}
#endif
