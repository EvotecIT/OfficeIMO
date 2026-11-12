using System.Data;
using System.Globalization;
using BenchmarkDotNet.Attributes;
using CsvHelper;
using CsvHelperWriter = CsvHelper.CsvWriter;

namespace OfficeIMO.CSV.Benchmarks;

/// <summary>CsvHelper 34 source-generated record writers for the shared write records.</summary>
[CsvRecordMap(typeof(QuotedTextRecord))] internal partial class QuotedTextRecordMap { }
[CsvRecordMap(typeof(NumberRecord))] internal partial class NumberRecordMap { }
[CsvRecordMap(typeof(WideRecord))] internal partial class WideRecordMap { }

/// <summary>
/// The CSV write comparison against CsvHelper 34: the same shapes, sources, sinks and byte-identity validation
/// as <c>CsvWriteComparisonBenchmarks</c>, which compares CsvHelper 33, Sep, Sylvan and ExcelReader.NET.
/// CsvHelper 34 writes records through its object mapping and readers through typed fields.
/// </summary>
[MemoryDiagnoser]
public class CsvWriteModernBenchmarks : CsvWriteBenchmarkBase {
    private static readonly CsvWriterOptions CsvHelperOptions = new() {
        NewLine = NewLine,
        FormatProvider = CultureInfo.InvariantCulture,
        WriteByteOrderMark = false
    };
    private DateRecordWithFormat[] _dateRecords = [];

    protected override void PrepareCompetitors() =>
        _dateRecords = Shape == CsvWriteShape.Dates
            ? Source.Records.Cast<DateRecord>().Select(DateRecordWithFormat.From).ToArray()
            : [];

    protected override IEnumerable<(string Name, Func<int> Write)> CompetitorLanes() => [
        (nameof(CsvHelper_ObjectsUtf8), CsvHelper_ObjectsUtf8), (nameof(CsvHelper_ObjectsText), CsvHelper_ObjectsText),
        (nameof(CsvHelper_DataTableUtf8), CsvHelper_DataTableUtf8), (nameof(CsvHelper_DataTableText), CsvHelper_DataTableText),
        (nameof(CsvHelper_DataReaderUtf8), CsvHelper_DataReaderUtf8), (nameof(CsvHelper_DataReaderText), CsvHelper_DataReaderText)];

    [Benchmark]
    public int CsvHelper_ObjectsUtf8() {
        using (CsvHelperWriter csv = CsvHelperWriter.Create(ResetStream(), CsvHelperOptions, leaveOpen: true)) WriteRecords(csv);
        return CompleteUtf8();
    }

    [Benchmark]
    public int CsvHelper_ObjectsText() {
        StringWriter writer = NewTextWriter();
        using (CsvHelperWriter csv = CsvHelperWriter.Create(writer, CsvHelperOptions, leaveOpen: true)) WriteRecords(csv);
        return CompleteText(writer);
    }

    [Benchmark]
    public int CsvHelper_DataTableUtf8() {
        using DataTableReader reader = Source.Table.CreateDataReader();
        using (CsvHelperWriter csv = CsvHelperWriter.Create(ResetStream(), CsvHelperOptions, leaveOpen: true)) WriteReader(csv, reader);
        return CompleteUtf8();
    }

    [Benchmark]
    public int CsvHelper_DataTableText() {
        StringWriter writer = NewTextWriter();
        using DataTableReader reader = Source.Table.CreateDataReader();
        using (CsvHelperWriter csv = CsvHelperWriter.Create(writer, CsvHelperOptions, leaveOpen: true)) WriteReader(csv, reader);
        return CompleteText(writer);
    }

    [Benchmark]
    public int CsvHelper_DataReaderUtf8() {
        using BenchmarkArrayDataReader reader = NewDataReader();
        using (CsvHelperWriter csv = CsvHelperWriter.Create(ResetStream(), CsvHelperOptions, leaveOpen: true)) WriteReader(csv, reader);
        return CompleteUtf8();
    }

    [Benchmark]
    public int CsvHelper_DataReaderText() {
        StringWriter writer = NewTextWriter();
        using BenchmarkArrayDataReader reader = NewDataReader();
        using (CsvHelperWriter csv = CsvHelperWriter.Create(writer, CsvHelperOptions, leaveOpen: true)) WriteReader(csv, reader);
        return CompleteText(writer);
    }

    private void WriteRecords(CsvHelperWriter csv) {
        switch (Shape) {
            case CsvWriteShape.QuotedText: csv.WriteRecords(Source.Records.Cast<QuotedTextRecord>()); break;
            case CsvWriteShape.Numbers: csv.WriteRecords(Source.Records.Cast<NumberRecord>()); break;
            case CsvWriteShape.Dates: csv.WriteRecords(_dateRecords); break;
            default: csv.WriteRecords(Source.Records.Cast<WideRecord>()); break;
        }
    }

    private void WriteReader(CsvHelperWriter csv, IDataReader reader) {
        csv.WriteHeader(Source.Headers);
        CsvWriteFieldKind[] kinds = Source.Kinds;
        while (reader.Read()) {
            for (int column = 0; column < kinds.Length; column++) {
                switch (kinds[column]) {
                    case CsvWriteFieldKind.Int: csv.WriteField(reader.GetInt32(column), default); break;
                    case CsvWriteFieldKind.Long: csv.WriteField(reader.GetInt64(column), default); break;
                    case CsvWriteFieldKind.Double: csv.WriteField(reader.GetDouble(column), default); break;
                    case CsvWriteFieldKind.Decimal: csv.WriteField(reader.GetDecimal(column), default); break;
                    case CsvWriteFieldKind.Date: csv.WriteField(reader.GetDateTime(column), DateFormat); break;
                    default: csv.WriteField(reader.GetString(column)); break;
                }
            }
            csv.EndRecord();
        }
    }

}

/// <summary>The date record with CsvHelper 34's format attribute, which is how its mapping takes a date format.</summary>
[CsvRecord]
public sealed partial class DateRecordWithFormat {
    public int Id { get; set; }
    [CsvFormat(CsvWriteSource.DateFormat)] public DateTime Created { get; set; }
    [CsvFormat(CsvWriteSource.DateFormat)] public DateTime Modified { get; set; }
    [CsvFormat(CsvWriteSource.DateFormat)] public DateTime Due { get; set; }
    public string Label { get; set; } = string.Empty;

    internal static DateRecordWithFormat From(DateRecord record) => new() {
        Id = record.Id, Created = record.Created, Modified = record.Modified, Due = record.Due, Label = record.Label
    };
}
