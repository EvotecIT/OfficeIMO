using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using CsvHelperConfiguration = CsvHelper.Configuration.CsvConfiguration;
using CsvHelperWriter = CsvHelper.CsvWriter;
using DataplatCsvDataReader = Dataplat.Dbatools.Csv.Reader.CsvDataReader;
using DataplatCsvReaderOptions = Dataplat.Dbatools.Csv.Reader.CsvReaderOptions;
using DataplatCsvWriter = Dataplat.Dbatools.Csv.Writer.CsvWriter;
using DataplatCsvWriterOptions = Dataplat.Dbatools.Csv.Writer.CsvWriterOptions;
using SylvanCsvDataWriter = Sylvan.Data.Csv.CsvDataWriter;
using SylvanCsvDataWriterOptions = Sylvan.Data.Csv.CsvDataWriterOptions;
#if NET9_0_OR_GREATER
using nietras.SeparatedValues;
using SepLib = nietras.SeparatedValues.Sep;
#endif

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// SQL Server table export and typed re-import at one million rows, shaped like the DbaClientX CSV export
/// benchmark: <c>int</c>, <c>nvarchar</c>, <c>decimal(18,2)</c> and <c>datetime2</c> columns, no header and the
/// <c>yyyy-MM-dd HH:mm:ss.fffffff</c> date format. The source is a forward-only <see cref="DbDataReader"/> with
/// typed getters, as a <c>SqlDataReader</c> provides. dbatools' <c>Export-DbaCsv</c> writes through the
/// Dataplat.Dbatools.Csv writer measured here. The pass-through lane wraps the same reader in a plain
/// <see cref="IDataReader"/>, as PSWriteOffice's <c>Export-OfficeCsv</c> does before handing it to the engine.
/// The project also targets net472 on Windows so the Windows PowerShell 5.1 runtime is measured too.
/// </summary>
[MemoryDiagnoser]
public class CsvDbatoolsExportBenchmarks {
    private const string DateFormat = "yyyy-MM-dd HH:mm:ss.fffffff";
    private const string ExpectedDate = "2026-07-08 00:00:00.0000000";
    private const int WriteBufferSize = 64 * 1024;
    private static readonly string[] Columns = ["Id", "DisplayName", "Score", "CreatedUtc"];

    private string _writePath = string.Empty;
    private string _readPath = string.Empty;
    private CsvSchema _officeSchema = null!;
    private Dictionary<string, Type> _dataplatColumnTypes = null!;

    [Params(1_000_000)]
    public int RowCount { get; set; }

    [GlobalSetup(Targets = [nameof(OfficeIMO_DataReader)])]
    public void SetupOfficeIMO() => SetupWrite(nameof(OfficeIMO_DataReader), OfficeIMO_DataReader);

    [GlobalSetup(Targets = [nameof(OfficeIMO_PassThroughDataReader)])]
    public void SetupOfficeIMOPassThrough() => SetupWrite(nameof(OfficeIMO_PassThroughDataReader), OfficeIMO_PassThroughDataReader);

    [GlobalSetup(Targets = [nameof(Dataplat_WriteFromReader)])]
    public void SetupDataplat() => SetupWrite(nameof(Dataplat_WriteFromReader), Dataplat_WriteFromReader);

    [GlobalSetup(Targets = [nameof(Sylvan_DataReader)])]
    public void SetupSylvan() => SetupWrite(nameof(Sylvan_DataReader), Sylvan_DataReader);

    [GlobalSetup(Targets = [nameof(CsvHelper_DataReader)])]
    public void SetupCsvHelper() => SetupWrite(nameof(CsvHelper_DataReader), CsvHelper_DataReader);

#if NET9_0_OR_GREATER
    [GlobalSetup(Targets = [nameof(Sep_DataReader)])]
    public void SetupSep() => SetupWrite(nameof(Sep_DataReader), Sep_DataReader);
#endif

    [GlobalSetup(Targets = [nameof(OfficeIMO_TypedRead), nameof(Dataplat_TypedRead)])]
    public void SetupRead() {
        Initialize();
        using (var writer = new StreamWriter(_readPath, append: false, new UTF8Encoding(false), WriteBufferSize)) {
            writer.WriteLine(string.Join(",", Columns));
            for (int id = 1; id <= RowCount; id++) {
                writer.Write(id.ToString(CultureInfo.InvariantCulture));
                writer.Write(",Row ");
                writer.Write(id.ToString(CultureInfo.InvariantCulture));
                writer.Write(',');
                writer.Write(ExportShapeDataReader.Score(id).ToString(CultureInfo.InvariantCulture));
                writer.Write(',');
                writer.WriteLine(ExpectedDate);
            }
        }

        ValidateRead(nameof(OfficeIMO_TypedRead), OfficeIMO_TypedRead());
        ValidateRead(nameof(Dataplat_TypedRead), Dataplat_TypedRead());
    }

    [GlobalCleanup]
    public void Cleanup() {
        if (File.Exists(_writePath)) File.Delete(_writePath);
        if (File.Exists(_readPath)) File.Delete(_readPath);
    }

    /// <summary>OfficeIMO.CSV row writer streaming the typed reader, as Export-OfficeCsv configures it.</summary>
    [Benchmark(Baseline = true)]
    public long OfficeIMO_DataReader() {
        using (var writer = CsvRowWriter.CreateFile(_writePath, CreateOfficeOptions(), append: false, bufferSize: WriteBufferSize)) {
            writer.WriteDataReader(new ExportShapeDataReader(RowCount));
        }

        return new FileInfo(_writePath).Length;
    }

    /// <summary>The same export through a plain <see cref="IDataReader"/> wrapper, as PSWriteOffice passes it.</summary>
    [Benchmark]
    public long OfficeIMO_PassThroughDataReader() {
        using (var writer = CsvRowWriter.CreateFile(_writePath, CreateOfficeOptions(), append: false, bufferSize: WriteBufferSize)) {
            writer.WriteDataReader(new PassThroughDataReader(new ExportShapeDataReader(RowCount)));
        }

        return new FileInfo(_writePath).Length;
    }

    /// <summary>Dataplat.Dbatools.Csv writer, the one dbatools' Export-DbaCsv uses.</summary>
    [Benchmark]
    public long Dataplat_WriteFromReader() {
        var options = new DataplatCsvWriterOptions {
            WriteHeader = false,
            Delimiter = ",",
            DateTimeFormat = DateFormat,
            NewLine = Environment.NewLine,
            Encoding = new UTF8Encoding(false),
            BufferSize = WriteBufferSize
        };
        using (var writer = new DataplatCsvWriter(_writePath, options)) {
            writer.WriteFromReader(new ExportShapeDataReader(RowCount));
        }

        return new FileInfo(_writePath).Length;
    }

    [Benchmark]
    public long Sylvan_DataReader() {
        var options = new SylvanCsvDataWriterOptions {
            WriteHeaders = false,
            DateTimeFormat = DateFormat,
            NewLine = Environment.NewLine,
            Culture = CultureInfo.InvariantCulture
        };
        using (var text = new StreamWriter(_writePath, append: false, new UTF8Encoding(false), WriteBufferSize))
        using (var writer = SylvanCsvDataWriter.Create(text, options)) {
            writer.Write(new ExportShapeDataReader(RowCount));
        }

        return new FileInfo(_writePath).Length;
    }

    [Benchmark]
    public long CsvHelper_DataReader() {
        var configuration = new CsvHelperConfiguration(CultureInfo.InvariantCulture) {
            HasHeaderRecord = false,
            NewLine = Environment.NewLine
        };
        using (var text = new StreamWriter(_writePath, append: false, new UTF8Encoding(false), WriteBufferSize))
        using (var csv = new CsvHelperWriter(text, configuration)) {
            csv.Context.TypeConverterOptionsCache.GetOptions<DateTime>().Formats = [DateFormat];
            using var reader = new ExportShapeDataReader(RowCount);
            while (reader.Read()) {
                csv.WriteField(reader.GetInt32(0));
                csv.WriteField(reader.GetString(1));
                csv.WriteField(reader.GetDecimal(2));
                csv.WriteField(reader.GetDateTime(3));
                csv.NextRecord();
            }
        }

        return new FileInfo(_writePath).Length;
    }

#if NET9_0_OR_GREATER
    [Benchmark]
    public long Sep_DataReader() {
        using (var csv = SepLib.New(',').Writer(options => options with { WriteHeader = false, Escape = true }).ToFile(_writePath)) {
            csv.Header.Add(Columns);
            using var reader = new ExportShapeDataReader(RowCount);
            Span<char> date = stackalloc char[32];
            while (reader.Read()) {
                using SepWriter.Row row = csv.NewRow();
                row[0].Format(reader.GetInt32(0));
                row[1].Set(reader.GetString(1));
                row[2].Format(reader.GetDecimal(2));
                reader.GetDateTime(3).TryFormat(date, out int written, DateFormat, CultureInfo.InvariantCulture);
                row[3].Set(date[..written]);
            }
        }

        return new FileInfo(_writePath).Length;
    }
#endif

    /// <summary>Typed import of the exported shape (with a header), consumed through GetValues as bulk copy does.</summary>
    [Benchmark]
    public long OfficeIMO_TypedRead() {
        using var reader = CsvDocument.OpenReader(_readPath, new CsvReaderOptions {
            Schema = _officeSchema,
            DateTimeFormats = [DateFormat],
            DetectDelimiter = false
        });
        return Consume(reader);
    }

    [Benchmark]
    public long Dataplat_TypedRead() {
        using var reader = new DataplatCsvDataReader(_readPath, new DataplatCsvReaderOptions {
            HasHeaderRow = true,
            Culture = CultureInfo.InvariantCulture,
            ColumnTypes = _dataplatColumnTypes,
            DateTimeFormats = [DateFormat]
        });
        return Consume(reader);
    }

    private static CsvSaveOptions CreateOfficeOptions() => new() {
        IncludeHeader = false,
        Delimiter = ',',
        Culture = CultureInfo.InvariantCulture,
        DateTimeFormat = DateFormat,
        Encoding = new UTF8Encoding(false)
    };

    private void Initialize() {
        // The Dataplat writer formats with the current culture; the DbaClientX harness runs invariant.
        CultureInfo.DefaultThreadCurrentCulture = CultureInfo.InvariantCulture;
        CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.CSV.Benchmarks", "DbatoolsExport");
        Directory.CreateDirectory(directory);
        string suffix = Guid.NewGuid().ToString("N");
        _writePath = Path.Combine(directory, "export-" + suffix + ".csv");
        _readPath = Path.Combine(directory, "import-" + suffix + ".csv");
        _officeSchema = new CsvSchemaBuilder()
            .Column("Id").AsInt32()
            .Column("DisplayName").AsString()
            .Column("Score").AsType(typeof(decimal))
            .Column("CreatedUtc").AsDateTime()
            .Done()
            .Build();
        _dataplatColumnTypes = new Dictionary<string, Type>(StringComparer.OrdinalIgnoreCase) {
            ["Id"] = typeof(int),
            ["DisplayName"] = typeof(string),
            ["Score"] = typeof(decimal),
            ["CreatedUtc"] = typeof(DateTime)
        };
    }

    private void SetupWrite(string lane, Func<long> write) {
        Initialize();
        write();
        ValidateExport(lane, _writePath, RowCount);
    }

    /// <summary>Applies the DbaClientX export benchmark's file check to every row.</summary>
    private static void ValidateExport(string lane, string path, int rowCount) {
        int rows = 0;
        using (var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true)) {
            string? line;
            while ((line = reader.ReadLine()) != null) {
                if (line.Length == 0) continue;
                string[] fields = line.Split(',');
                if (fields.Length != 4) throw new InvalidOperationException($"{lane}: row {rows + 1} has {fields.Length} fields: '{line}'.");
                int id = int.Parse(fields[0], NumberStyles.Integer, CultureInfo.InvariantCulture);
                if (id != rows + 1) throw new InvalidOperationException($"{lane}: row {rows + 1} has Id {id}.");
                if (fields[1] != "Row " + fields[0]) throw new InvalidOperationException($"{lane}: row {id} has DisplayName '{fields[1]}'.");
                string expectedScore = ExportShapeDataReader.Score(id).ToString("0.00", CultureInfo.InvariantCulture);
                if (fields[2] != expectedScore) throw new InvalidOperationException($"{lane}: row {id} has Score '{fields[2]}', expected '{expectedScore}'.");
                if (fields[3] != ExpectedDate) throw new InvalidOperationException($"{lane}: row {id} has CreatedUtc '{fields[3]}'.");
                rows++;
            }
        }

        if (rows != rowCount) throw new InvalidOperationException($"{lane}: wrote {rows} rows, expected {rowCount}.");
    }

    private void ValidateRead(string lane, long checksum) {
        long expected = 0;
        for (int id = 1; id <= RowCount; id++) expected += ExpectedRowChecksum(id, ("Row " + id.ToString(CultureInfo.InvariantCulture)).Length, ExportShapeDataReader.Score(id));
        if (checksum != expected) throw new InvalidOperationException($"{lane}: checksum {checksum}, expected {expected}.");
    }

    private static long Consume(IDataReader reader) {
        var values = new object[reader.FieldCount];
        long checksum = 0;
        var expectedDate = new DateTime(2026, 7, 8);
        while (reader.Read()) {
            if (reader.GetValues(values) != 4) throw new InvalidOperationException("The reader did not return every field.");
            if ((DateTime)values[3] != expectedDate) throw new InvalidOperationException($"Unexpected date {values[3]}.");
            checksum += ExpectedRowChecksum((int)values[0], ((string)values[1]).Length, (decimal)values[2]);
        }

        return checksum;
    }

    private static long ExpectedRowChecksum(int id, int nameLength, decimal score) =>
        id + nameLength + (long)(score * 100m);

    /// <summary>Generates the export rows on demand with typed getters, like a <c>SqlDataReader</c>.</summary>
    private sealed class ExportShapeDataReader(int rowCount) : DbDataReader {
        private static readonly DateTime Created = new(2026, 7, 8);
        private int _id;
        private string? _name;

        /// <summary><c>CONVERT(decimal(18,2), Id * 1.25)</c>: two decimal places, as SQL Server returns it.</summary>
        internal static decimal Score(int id) => new(checked(id * 125), 0, 0, false, 2);

        public override int FieldCount => 4;
        public override int RecordsAffected => -1;
        public override bool HasRows => rowCount > 0;
        public override bool IsClosed => false;
        public override int Depth => 0;
        public override object this[int ordinal] => GetValue(ordinal);
        public override object this[string name] => GetValue(GetOrdinal(name));

        public override bool Read() {
            if (_id >= rowCount) return false;
            _id++;
            _name = null;
            return true;
        }

        public override bool NextResult() => false;
        public override string GetName(int ordinal) => Columns[ordinal];
        public override int GetOrdinal(string name) => Array.IndexOf(Columns, name);
        public override string GetDataTypeName(int ordinal) => GetFieldType(ordinal).Name;

        public override Type GetFieldType(int ordinal) => ordinal switch {
            0 => typeof(int),
            1 => typeof(string),
            2 => typeof(decimal),
            3 => typeof(DateTime),
            _ => throw new IndexOutOfRangeException()
        };

        public override int GetInt32(int ordinal) => ordinal == 0 ? _id : throw new InvalidCastException();
        public override string GetString(int ordinal) => ordinal == 1 ? _name ??= "Row " + _id.ToString(CultureInfo.InvariantCulture) : throw new InvalidCastException();
        public override decimal GetDecimal(int ordinal) => ordinal == 2 ? Score(_id) : throw new InvalidCastException();
        public override DateTime GetDateTime(int ordinal) => ordinal == 3 ? Created : throw new InvalidCastException();

        public override object GetValue(int ordinal) => ordinal switch {
            0 => _id,
            1 => GetString(1),
            2 => Score(_id),
            3 => Created,
            _ => throw new IndexOutOfRangeException()
        };

        public override int GetValues(object[] values) {
            int count = Math.Min(values.Length, 4);
            for (int i = 0; i < count; i++) values[i] = GetValue(i);
            return count;
        }

        public override T GetFieldValue<T>(int ordinal) {
            if (typeof(T) == typeof(int)) return (T)(object)GetInt32(ordinal);
            if (typeof(T) == typeof(decimal)) return (T)(object)GetDecimal(ordinal);
            if (typeof(T) == typeof(DateTime)) return (T)(object)GetDateTime(ordinal);
            return (T)GetValue(ordinal);
        }

        public override bool IsDBNull(int ordinal) => false;
        public override bool GetBoolean(int ordinal) => throw new InvalidCastException();
        public override byte GetByte(int ordinal) => throw new InvalidCastException();
        public override long GetBytes(int ordinal, long dataOffset, byte[]? buffer, int bufferOffset, int length) => throw new InvalidCastException();
        public override char GetChar(int ordinal) => throw new InvalidCastException();
        public override long GetChars(int ordinal, long dataOffset, char[]? buffer, int bufferOffset, int length) => throw new InvalidCastException();
        public override Guid GetGuid(int ordinal) => throw new InvalidCastException();
        public override short GetInt16(int ordinal) => throw new InvalidCastException();
        public override long GetInt64(int ordinal) => throw new InvalidCastException();
        public override float GetFloat(int ordinal) => throw new InvalidCastException();
        public override double GetDouble(int ordinal) => throw new InvalidCastException();
        public override System.Collections.IEnumerator GetEnumerator() => new DbEnumerator(this);
    }

    /// <summary>Forwards every member to the wrapped reader without being a <see cref="DbDataReader"/>.</summary>
    private sealed class PassThroughDataReader(IDataReader inner) : IDataReader {
        public object this[int i] => inner[i];
        public object this[string name] => inner[name];
        public int Depth => inner.Depth;
        public bool IsClosed => inner.IsClosed;
        public int RecordsAffected => inner.RecordsAffected;
        public int FieldCount => inner.FieldCount;
        public void Close() => inner.Close();
        public void Dispose() => inner.Dispose();
        public bool GetBoolean(int i) => inner.GetBoolean(i);
        public byte GetByte(int i) => inner.GetByte(i);
        public long GetBytes(int i, long fieldOffset, byte[]? buffer, int bufferoffset, int length) => inner.GetBytes(i, fieldOffset, buffer, bufferoffset, length);
        public char GetChar(int i) => inner.GetChar(i);
        public long GetChars(int i, long fieldoffset, char[]? buffer, int bufferoffset, int length) => inner.GetChars(i, fieldoffset, buffer, bufferoffset, length);
        public IDataReader GetData(int i) => inner.GetData(i);
        public string GetDataTypeName(int i) => inner.GetDataTypeName(i);
        public DateTime GetDateTime(int i) => inner.GetDateTime(i);
        public decimal GetDecimal(int i) => inner.GetDecimal(i);
        public double GetDouble(int i) => inner.GetDouble(i);
        public Type GetFieldType(int i) => inner.GetFieldType(i);
        public float GetFloat(int i) => inner.GetFloat(i);
        public Guid GetGuid(int i) => inner.GetGuid(i);
        public short GetInt16(int i) => inner.GetInt16(i);
        public int GetInt32(int i) => inner.GetInt32(i);
        public long GetInt64(int i) => inner.GetInt64(i);
        public string GetName(int i) => inner.GetName(i);
        public int GetOrdinal(string name) => inner.GetOrdinal(name);
        public DataTable? GetSchemaTable() => inner.GetSchemaTable();
        public string GetString(int i) => inner.GetString(i);
        public object GetValue(int i) => inner.GetValue(i);
        public int GetValues(object[] values) => inner.GetValues(values);
        public bool IsDBNull(int i) => inner.IsDBNull(i);
        public bool NextResult() => inner.NextResult();
        public bool Read() => inner.Read();
    }
}
