#if NET10_0_OR_GREATER
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using ExcelReader.Core.Parser;
using OfficeIMO.Data;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>Matched, materialized four-column scans and object mapping of the same generated workbook.</summary>
[MemoryDiagnoser]
[CategoriesColumn]
[GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
public partial class ExcelReaderGeneratedBenchmarks {
    internal static readonly string[] Headers = ["Name", "Id", "CreatedOn", "Amount"];
    private string _path = string.Empty;
    private long _expected;
    private bool _validate;
    private bool _ownsInput;
    internal string InputPath => _path;
    private static readonly ExcelParser<GeneratedRecord> Parser = ExcelParser.FromAttributes<GeneratedRecord>();

    [Params(ExcelFileFormat.Xlsx, ExcelFileFormat.Xlsb, ExcelFileFormat.Xls)]
    public ExcelFileFormat Format { get; set; }

    [Params(50_000)]
    public int RowCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_DATA") ?? Path.GetTempPath();
        Directory.CreateDirectory(root);
        string? pinnedRoot = Environment.GetEnvironmentVariable("OFFICEIMO_EXCEL_GENERATED_INPUT");
        _ownsInput = string.IsNullOrEmpty(pinnedRoot);
        _path = _ownsInput
            ? Path.Combine(root, $"generated-four-columns-{Guid.NewGuid():N}.{Format.ToString().ToLowerInvariant()}")
            : Path.Combine(pinnedRoot!, $"generated-four-columns-{RowCount}.{Format.ToString().ToLowerInvariant()}");
        if (_ownsInput) File.WriteAllBytes(_path, WriteOfficeIMO(Format, RowCount));
        else if (!File.Exists(_path)) throw new FileNotFoundException("Pinned benchmark input is missing.", _path);
        _expected = 0;
        for (int index = 0; index < RowCount; index++) _expected = Add(_expected, CreateRecord(index));
        _validate = true;
        try {
            OfficeIMO_Read();
            ExcelReaderNet_Read();
            Sylvan_Read();
            ExcelDataReader_Read();
            OfficeIMO_Parse();
            ExcelReaderNet_Parse();
        } catch { Cleanup(); throw; }
        finally { _validate = false; }
    }

    [GlobalCleanup]
    public void Cleanup() { if (_ownsInput && _path.Length > 0) File.Delete(_path); }

    [Benchmark(Baseline = true), BenchmarkCategory("Read")]
    public long OfficeIMO_Read() {
        using var reader = ExcelDocument.OpenReader(_path);
        return Read(reader);
    }

    [Benchmark, BenchmarkCategory("Read")]
    public long Sylvan_Read() {
        using var reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(_path);
        return Read(reader);
    }

    [Benchmark, BenchmarkCategory("Read")]
    public long ExcelDataReader_Read() {
        using var stream = File.OpenRead(_path);
        using var reader = global::ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
        if (!reader.Read()) throw new InvalidDataException("Missing header.");
        ValidateHeaders(reader);
        return Read(reader);
    }

    internal long Read(IDataReader reader) {
        ValidateHeaders(reader, reader is DbDataReader);
        long checksum = 0;
        int count = 0;
        while (reader.Read()) {
            string name = reader.GetString(0);
            int id = reader is DbDataReader ? reader.GetInt32(1) : checked((int)reader.GetDouble(1));
            DateTime date = reader.GetDateTime(2);
            double amount = reader.GetDouble(3);
            if (_validate) Validate(count, name, id, date, amount);
            checksum = Add(checksum, name, id, date, amount);
            count++;
        }
        return Check(checksum, count);
    }

    [Benchmark(Baseline = true), BenchmarkCategory("Parse")]
    public long OfficeIMO_Parse() {
        using var reader = ExcelDocument.OpenReader(_path);
        return Observe(reader.RowsAs<GeneratedRecord>());
    }

    [Benchmark, BenchmarkCategory("Parse")]
    public long ExcelReaderNet_Parse() {
        switch (Format) {
            case ExcelFileFormat.Xlsx:
                using (var reader = ExcelReaderApi.FromXlsxFile(_path)) return Observe(Parser.Parse(reader.FirstSheet));
            case ExcelFileFormat.Xlsb:
                using (var reader = ExcelReaderApi.FromXlsbFile(_path)) return Observe(Parser.Parse(reader.FirstSheet));
            case ExcelFileFormat.Xls:
                using (var reader = ExcelReaderApi.FromXlsFile(_path)) return Observe(Parser.Parse(reader.FirstSheet));
            default: throw new ArgumentOutOfRangeException();
        }
    }

    private long Observe(IEnumerable<GeneratedRecord> rows) {
        long checksum = 0;
        int count = 0;
        foreach (var row in rows) {
            if (_validate) Validate(count, row.Name, row.Id, row.CreatedOn, row.Amount);
            checksum = Add(checksum, row);
            count++;
        }
        return Check(checksum, count);
    }

    private long Check(long checksum, int count) {
        if (count != RowCount || checksum != _expected) throw new InvalidDataException($"{Format}: unexpected row count or checksum: {count}, {checksum}.");
        return checksum;
    }

    internal static void ValidateHeaders(IDataRecord reader, bool useNames = false) {
        if (reader.FieldCount != Headers.Length) throw new InvalidDataException("Incorrect field count.");
        for (int i = 0; i < Headers.Length; i++) {
            if ((useNames ? reader.GetName(i) : reader.GetString(i)) != Headers[i]) throw new InvalidDataException("Incorrect header.");
        }
    }

    internal static GeneratedRecord CreateRecord(int index) => new() {
        Name = "Name " + index.ToString(CultureInfo.InvariantCulture), Id = index + 1,
        CreatedOn = new DateTime(2024, 1, 1).AddDays(index % 365), Amount = index * 1.25
    };

    internal static void Validate(int index, string name, int id, DateTime date, double amount) {
        var expected = CreateRecord(index);
        if (name != expected.Name || id != expected.Id || date != expected.CreatedOn || amount != expected.Amount)
            throw new InvalidDataException($"Incorrect row {index}.");
    }

    private static long Add(long checksum, GeneratedRecord row) => Add(checksum, row.Name, row.Id, row.CreatedOn, row.Amount);
    private static long Add(long checksum, string name, int id, DateTime date, double amount) {
        unchecked {
            // Observe every decoded character and typed value, identically for all engines.
            foreach (char c in name) checksum = checksum * 31 + c;
            return ((checksum * 31 + id) * 31 + date.Ticks) * 31 + BitConverter.DoubleToInt64Bits(amount);
        }
    }

    public sealed class GeneratedRecord {
        public string Name { get; set; } = string.Empty;
        public int Id { get; set; }
        public DateTime CreatedOn { get; set; }
        public double Amount { get; set; }
    }
}
#endif
