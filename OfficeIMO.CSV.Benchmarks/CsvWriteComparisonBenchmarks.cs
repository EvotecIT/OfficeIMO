using System.Data;
using System.Globalization;
using BenchmarkDotNet.Attributes;
using nietras.SeparatedValues;
using CsvHelperConfiguration = CsvHelper.Configuration.CsvConfiguration;
using CsvHelperWriter = CsvHelper.CsvWriter;
using SepLib = nietras.SeparatedValues.Sep;
using SylvanCsvDataWriter = Sylvan.Data.Csv.CsvDataWriter;
using SylvanCsvDataWriterOptions = Sylvan.Data.Csv.CsvDataWriterOptions;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>
/// CSV writing from typed objects, a <see cref="DataTable"/> and an <see cref="IDataReader"/>, to a UTF-8 stream
/// and a <see cref="TextWriter"/>, against CsvHelper 33, Sep, Sylvan and ExcelReader.NET. The OfficeIMO lanes,
/// shapes and byte-identity validation come from <see cref="CsvWriteBenchmarkBase"/>; CsvHelper 34 runs in the
/// Modern project. Each library is driven as its users drive it: object mapping where it has one, typed field
/// writes otherwise.
/// </summary>
[MemoryDiagnoser]
public class CsvWriteComparisonBenchmarks : CsvWriteBenchmarkBase {
    // With NewLine set, CsvHelper 33 quotes only fields containing that exact newline; keep its default rule of
    // quoting any CR or LF so a bare line feed inside a field cannot split the record.
    private static readonly CsvHelperConfiguration CsvHelperOptions = new(CultureInfo.InvariantCulture) {
        NewLine = NewLine,
        ShouldQuote = args => args.Field is { Length: > 0 } field
            && (field.AsSpan().IndexOfAny("\",\r\n") >= 0 || field[0] == ' ' || field[^1] == ' ')
    };
    private static readonly SylvanCsvDataWriterOptions SylvanOptions = new() { NewLine = NewLine, DateTimeFormat = DateFormat };
    private static readonly SepWriterOptions SepOptions = SepLib.New(',').Writer(options => options with { Escape = true });

    protected override IEnumerable<(string Name, Func<int> Write)> CompetitorLanes() => [
        (nameof(CsvHelper_ObjectsUtf8), CsvHelper_ObjectsUtf8), (nameof(CsvHelper_ObjectsText), CsvHelper_ObjectsText),
        (nameof(CsvHelper_DataTableUtf8), CsvHelper_DataTableUtf8), (nameof(CsvHelper_DataTableText), CsvHelper_DataTableText),
        (nameof(CsvHelper_DataReaderUtf8), CsvHelper_DataReaderUtf8), (nameof(CsvHelper_DataReaderText), CsvHelper_DataReaderText),
#if NET9_0_OR_GREATER
        (nameof(Sep_ObjectsUtf8), Sep_ObjectsUtf8), (nameof(Sep_ObjectsText), Sep_ObjectsText),
#endif
        (nameof(Sep_DataTableUtf8), Sep_DataTableUtf8), (nameof(Sep_DataTableText), Sep_DataTableText),
        (nameof(Sep_DataReaderUtf8), Sep_DataReaderUtf8), (nameof(Sep_DataReaderText), Sep_DataReaderText),
        (nameof(Sylvan_DataTableUtf8), Sylvan_DataTableUtf8), (nameof(Sylvan_DataTableText), Sylvan_DataTableText),
        (nameof(Sylvan_DataReaderUtf8), Sylvan_DataReaderUtf8), (nameof(Sylvan_DataReaderText), Sylvan_DataReaderText),
#if NET10_0_OR_GREATER
        (nameof(ExcelReaderNet_ObjectsUtf8), ExcelReaderNet_ObjectsUtf8),
        (nameof(ExcelReaderNet_DataTableUtf8), ExcelReaderNet_DataTableUtf8),
        (nameof(ExcelReaderNet_DataReaderUtf8), ExcelReaderNet_DataReaderUtf8),
#endif
    ];

    // CsvHelper 33: object mapping for records, typed fields for readers.

    [Benchmark]
    public int CsvHelper_ObjectsUtf8() {
        using (var text = new StreamWriter(ResetStream(), Utf8NoBom, 64 * 1024, leaveOpen: true)) WriteCsvHelperObjects(text);
        return CompleteUtf8();
    }

    [Benchmark]
    public int CsvHelper_ObjectsText() {
        StringWriter writer = NewTextWriter();
        WriteCsvHelperObjects(writer);
        return CompleteText(writer);
    }

    [Benchmark]
    public int CsvHelper_DataTableUtf8() {
        using DataTableReader reader = Source.Table.CreateDataReader();
        using (var text = new StreamWriter(ResetStream(), Utf8NoBom, 64 * 1024, leaveOpen: true)) WriteCsvHelperReader(text, reader);
        return CompleteUtf8();
    }

    [Benchmark]
    public int CsvHelper_DataTableText() {
        StringWriter writer = NewTextWriter();
        using DataTableReader reader = Source.Table.CreateDataReader();
        WriteCsvHelperReader(writer, reader);
        return CompleteText(writer);
    }

    [Benchmark]
    public int CsvHelper_DataReaderUtf8() {
        using BenchmarkArrayDataReader reader = NewDataReader();
        using (var text = new StreamWriter(ResetStream(), Utf8NoBom, 64 * 1024, leaveOpen: true)) WriteCsvHelperReader(text, reader);
        return CompleteUtf8();
    }

    [Benchmark]
    public int CsvHelper_DataReaderText() {
        StringWriter writer = NewTextWriter();
        using BenchmarkArrayDataReader reader = NewDataReader();
        WriteCsvHelperReader(writer, reader);
        return CompleteText(writer);
    }

    private static CsvHelperWriter CreateCsvHelper(TextWriter writer) {
        var csv = new CsvHelperWriter(writer, CsvHelperOptions, leaveOpen: true);
        csv.Context.TypeConverterOptionsCache.GetOptions<DateTime>().Formats = [DateFormat];
        return csv;
    }

    private void WriteCsvHelperObjects(TextWriter writer) {
        using CsvHelperWriter csv = CreateCsvHelper(writer);
        switch (Shape) {
            case CsvWriteShape.QuotedText: csv.WriteRecords(Source.Records.Cast<QuotedTextRecord>()); break;
            case CsvWriteShape.Numbers: csv.WriteRecords(Source.Records.Cast<NumberRecord>()); break;
            case CsvWriteShape.Dates: csv.WriteRecords(Source.Records.Cast<DateRecord>()); break;
            default: csv.WriteRecords(Source.Records.Cast<WideRecord>()); break;
        }
    }

    private void WriteCsvHelperReader(TextWriter writer, IDataReader reader) {
        using CsvHelperWriter csv = CreateCsvHelper(writer);
        CsvWriteFieldKind[] kinds = Source.Kinds;
        for (int column = 0; column < kinds.Length; column++) csv.WriteField(reader.GetName(column));
        csv.NextRecord();
        while (reader.Read()) {
            for (int column = 0; column < kinds.Length; column++) {
                switch (kinds[column]) {
                    case CsvWriteFieldKind.Int: csv.WriteField(reader.GetInt32(column)); break;
                    case CsvWriteFieldKind.Long: csv.WriteField(reader.GetInt64(column)); break;
                    case CsvWriteFieldKind.Double: csv.WriteField(reader.GetDouble(column)); break;
                    case CsvWriteFieldKind.Decimal: csv.WriteField(reader.GetDecimal(column)); break;
                    case CsvWriteFieldKind.Date: csv.WriteField(reader.GetDateTime(column)); break;
                    default: csv.WriteField(reader.GetString(column)); break;
                }
            }
            csv.NextRecord();
        }
    }

    // Sep: typed column writes by index after declaring the header.
#if NET9_0_OR_GREATER
    [Benchmark]
    public int Sep_ObjectsUtf8() {
        using (SepWriter csv = SepOptions.To(ResetStream(), leaveOpen: true)) WriteSepRecords(csv);
        return CompleteUtf8();
    }

    [Benchmark]
    public int Sep_ObjectsText() {
        StringWriter writer = NewTextWriter();
        using (SepWriter csv = SepOptions.To(writer, leaveOpen: true)) WriteSepRecords(csv);
        return CompleteText(writer);
    }

    private void WriteSepRecords(SepWriter csv) {
        csv.Header.Add(Source.Headers);
        foreach (ICsvWriteRecord record in Source.Records) {
            using SepWriter.Row row = csv.NewRow();
            var sink = new SepFieldSink(row);
            record.Write(ref sink);
        }
    }

    private ref struct SepFieldSink(SepWriter.Row row) : ICsvWriteFieldSink {
        private readonly SepWriter.Row _row = row;
        private int _column;
        public void Int(int value) => _row[_column++].Format(value);
        public void Long(long value) => _row[_column++].Format(value);
        public void Double(double value) => _row[_column++].Format(value);
        public void Decimal(decimal value) => _row[_column++].Format(value);
        public void Text(string value) => _row[_column++].Set(value);
        public void Date(DateTime value) {
            Span<char> buffer = stackalloc char[32];
            value.TryFormat(buffer, out int written, DateFormat, CultureInfo.InvariantCulture);
            _row[_column++].Set(buffer.Slice(0, written));
        }
    }
#endif

    [Benchmark]
    public int Sep_DataTableUtf8() {
        using DataTableReader reader = Source.Table.CreateDataReader();
        using (SepWriter csv = SepOptions.To(ResetStream(), leaveOpen: true)) WriteSepReader(csv, reader);
        return CompleteUtf8();
    }

    [Benchmark]
    public int Sep_DataTableText() {
        StringWriter writer = NewTextWriter();
        using DataTableReader reader = Source.Table.CreateDataReader();
        using (SepWriter csv = SepOptions.To(writer, leaveOpen: true)) WriteSepReader(csv, reader);
        return CompleteText(writer);
    }

    [Benchmark]
    public int Sep_DataReaderUtf8() {
        using BenchmarkArrayDataReader reader = NewDataReader();
        using (SepWriter csv = SepOptions.To(ResetStream(), leaveOpen: true)) WriteSepReader(csv, reader);
        return CompleteUtf8();
    }

    [Benchmark]
    public int Sep_DataReaderText() {
        StringWriter writer = NewTextWriter();
        using BenchmarkArrayDataReader reader = NewDataReader();
        using (SepWriter csv = SepOptions.To(writer, leaveOpen: true)) WriteSepReader(csv, reader);
        return CompleteText(writer);
    }

    private void WriteSepReader(SepWriter csv, IDataReader reader) {
        csv.Header.Add(Source.Headers);
        CsvWriteFieldKind[] kinds = Source.Kinds;
        Span<char> buffer = stackalloc char[32];
        while (reader.Read()) {
            using SepWriter.Row row = csv.NewRow();
            for (int column = 0; column < kinds.Length; column++) {
                switch (kinds[column]) {
                    case CsvWriteFieldKind.Int: row[column].Format(reader.GetInt32(column)); break;
                    case CsvWriteFieldKind.Long: row[column].Format(reader.GetInt64(column)); break;
                    case CsvWriteFieldKind.Double: row[column].Format(reader.GetDouble(column)); break;
                    case CsvWriteFieldKind.Decimal: row[column].Format(reader.GetDecimal(column)); break;
                    case CsvWriteFieldKind.Date:
                        reader.GetDateTime(column).TryFormat(buffer, out int written, DateFormat, CultureInfo.InvariantCulture);
                        row[column].Set(buffer.Slice(0, written));
                        break;
                    default: row[column].Set(reader.GetString(column)); break;
                }
            }
        }
    }

    // Sylvan writes data readers.

    [Benchmark]
    public int Sylvan_DataTableUtf8() {
        using DataTableReader reader = Source.Table.CreateDataReader();
        using (var text = new StreamWriter(ResetStream(), Utf8NoBom, 64 * 1024, leaveOpen: true))
        using (SylvanCsvDataWriter csv = SylvanCsvDataWriter.Create(text, SylvanOptions)) csv.Write(reader);
        return CompleteUtf8();
    }

    [Benchmark]
    public int Sylvan_DataTableText() {
        StringWriter writer = NewTextWriter();
        using DataTableReader reader = Source.Table.CreateDataReader();
        using (SylvanCsvDataWriter csv = SylvanCsvDataWriter.Create(writer, SylvanOptions)) csv.Write(reader);
        return CompleteText(writer);
    }

    [Benchmark]
    public int Sylvan_DataReaderUtf8() {
        using BenchmarkArrayDataReader reader = NewDataReader();
        using (var text = new StreamWriter(ResetStream(), Utf8NoBom, 64 * 1024, leaveOpen: true))
        using (SylvanCsvDataWriter csv = SylvanCsvDataWriter.Create(text, SylvanOptions)) csv.Write(reader);
        return CompleteUtf8();
    }

    [Benchmark]
    public int Sylvan_DataReaderText() {
        StringWriter writer = NewTextWriter();
        using BenchmarkArrayDataReader reader = NewDataReader();
        using (SylvanCsvDataWriter csv = SylvanCsvDataWriter.Create(writer, SylvanOptions)) csv.Write(reader);
        return CompleteText(writer);
    }

#if NET10_0_OR_GREATER
    // ExcelReader.NET writes CSV to streams only, with typed row writes.

    [Benchmark]
    public int ExcelReaderNet_ObjectsUtf8() {
        using (ExcelReaderNetCsvWriter workbook = ExcelReaderNetCsvWriter.Create(ResetStream(), leaveOpen: true))
        using (var csv = workbook.AddSheet("Data")) {
            using (ExcelReaderNetCsvRowWriter header = csv.StartRow()) {
                foreach (string name in Source.Headers) header.Write(name);
            }
            foreach (ICsvWriteRecord record in Source.Records) {
                using ExcelReaderNetCsvRowWriter row = csv.StartRow();
                var sink = new ExcelReaderFieldSink(row);
                record.Write(ref sink);
            }
        }
        return CompleteUtf8();
    }

    [Benchmark]
    public int ExcelReaderNet_DataTableUtf8() {
        using DataTableReader reader = Source.Table.CreateDataReader();
        WriteExcelReaderReader(reader);
        return CompleteUtf8();
    }

    [Benchmark]
    public int ExcelReaderNet_DataReaderUtf8() {
        using BenchmarkArrayDataReader reader = NewDataReader();
        WriteExcelReaderReader(reader);
        return CompleteUtf8();
    }

    private void WriteExcelReaderReader(IDataReader reader) {
        using ExcelReaderNetCsvWriter workbook = ExcelReaderNetCsvWriter.Create(ResetStream(), leaveOpen: true);
        using var csv = workbook.AddSheet("Data");
        using (ExcelReaderNetCsvRowWriter header = csv.StartRow()) {
            foreach (string name in Source.Headers) header.Write(name);
        }
        CsvWriteFieldKind[] kinds = Source.Kinds;
        while (reader.Read()) {
            using ExcelReaderNetCsvRowWriter row = csv.StartRow();
            var sink = new ExcelReaderFieldSink(row);
            for (int column = 0; column < kinds.Length; column++) {
                switch (kinds[column]) {
                    case CsvWriteFieldKind.Int: sink.Int(reader.GetInt32(column)); break;
                    case CsvWriteFieldKind.Long: sink.Long(reader.GetInt64(column)); break;
                    case CsvWriteFieldKind.Double: sink.Double(reader.GetDouble(column)); break;
                    case CsvWriteFieldKind.Decimal: sink.Decimal(reader.GetDecimal(column)); break;
                    case CsvWriteFieldKind.Date: sink.Date(reader.GetDateTime(column)); break;
                    default: sink.Text(reader.GetString(column)); break;
                }
            }
        }
    }

    private readonly struct ExcelReaderFieldSink(ExcelReaderNetCsvRowWriter row) : ICsvWriteFieldSink {
        private readonly ExcelReaderNetCsvRowWriter _row = row;
        public void Int(int value) => _row.Write(value);
        public void Long(long value) => _row.Write(value);
        public void Double(double value) => _row.Write(value);
        public void Decimal(decimal value) => _row.Write(value);
        public void Text(string value) => _row.Write(value);
        public void Date(DateTime value) {
            Span<char> buffer = stackalloc char[32];
            value.TryFormat(buffer, out int written, DateFormat, CultureInfo.InvariantCulture);
            _row.Write(new string(buffer.Slice(0, written)));
        }
    }
#endif
}
