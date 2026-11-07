using System.Data.Common;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using BenchmarkDotNet.Attributes;
using OfficeIMO.CSV;
using OfficeIMO.Data;

using OfficeIMO.Benchmarks;

namespace OfficeIMO.CSV.Benchmarks;

[global::CsvHelper.CsvRecordMap(typeof(CsvBenchmarkRow))]
internal partial class GeneratedCsvRecordMap { }

/// <summary>Compares complete typed CSV reads and writes using current public APIs.</summary>
[MemoryDiagnoser]
public class CsvGeneratedModernTypedBenchmarks {
    private CsvBenchmarkRow[] _rows = [];
    private byte[] _input = [];
    private bool _validate;
    private int _expectedChecksum;

    [GlobalSetup]
    public void Setup() {
        BenchmarkProcessPriority.ApplyConfiguredPriority();
        _rows = CsvBenchmarkData.Create(50_000, CsvBenchmarkShape.Mixed);
        _input = Encoding.UTF8.GetBytes(CsvBenchmarkData.CreateText(_rows));
        _expectedChecksum = _rows.Aggregate(17, AddChecksum);
        _validate = true;
        try {
            OfficeIMO_Parse();
            CsvHelper_Parse();
            foreach (var (name, output) in new[] { ("OfficeIMO", OfficeIMO_Write()), ("OfficeIMO_Stream", OfficeIMO_WriteStream()), ("CsvHelper", CsvHelper_Write()) }) {
                ValidateHeaders(output);
                Consume(ReadOfficeIMO(output));
                Consume(ReadCsvHelper(output));
                Console.WriteLine($"VALIDATED_WRITE {name} bytes={output.Length} SHA256={Convert.ToHexString(SHA256.HashData(output))}");
            }
        } finally { _validate = false; }
        Console.WriteLine("Generated typed CSV input: 50000 rows, 10 columns, UTF-8 SHA256 " + Convert.ToHexString(SHA256.HashData(_input)));
    }

    [Benchmark]
    public int OfficeIMO_Parse() {
        using var input = new MemoryStream(_input, writable: false);
        using DbDataReader reader = CsvDocument.OpenReader(input, new CsvReaderOptions { Culture = CultureInfo.InvariantCulture });
        return Consume(reader.RowsAs<CsvBenchmarkRow>());
    }

    [Benchmark]
    public int CsvHelper_Parse() {
        using var input = new MemoryStream(_input, writable: false);
        using var reader = global::CsvHelper.CsvReader.Create(input,
            new global::CsvHelper.CsvReaderOptions { FormatProvider = CultureInfo.InvariantCulture });
        return Consume(reader.GetRecords<CsvBenchmarkRow>());
    }

    [Benchmark]
    public byte[] OfficeIMO_Write() {
        using var output = new MemoryStream();
        using (var text = new StreamWriter(output, new UTF8Encoding(false), leaveOpen: true))
            CsvDocument.WriteObjects(text, _rows, new CsvSaveOptions { NewLine = "\n", Culture = CultureInfo.InvariantCulture });
        return output.ToArray();
    }

    [Benchmark]
    public byte[] OfficeIMO_WriteStream() {
        using var output = new MemoryStream();
        using (var writer = CsvRowWriter.CreateStream(output,
            new CsvSaveOptions { NewLine = "\n", Culture = CultureInfo.InvariantCulture }, leaveOpen: true))
            foreach (CsvBenchmarkRow row in _rows) writer.WriteObject(row);
        return output.ToArray();
    }

    [Benchmark]
    public byte[] CsvHelper_Write() {
        using var output = new MemoryStream();
        using (var writer = global::CsvHelper.CsvWriter.Create(output,
            new global::CsvHelper.CsvWriterOptions { NewLine = "\n", FormatProvider = CultureInfo.InvariantCulture }, leaveOpen: true))
            writer.WriteRecords(_rows);
        return output.ToArray();
    }

    private static IEnumerable<CsvBenchmarkRow> ReadOfficeIMO(byte[] data) {
        using var input = new MemoryStream(data, writable: false);
        using DbDataReader reader = CsvDocument.OpenReader(input, new CsvReaderOptions { Culture = CultureInfo.InvariantCulture });
        foreach (CsvBenchmarkRow row in reader.RowsAs<CsvBenchmarkRow>()) yield return row;
    }

    private static IEnumerable<CsvBenchmarkRow> ReadCsvHelper(byte[] data) {
        using var input = new MemoryStream(data, writable: false);
        using var reader = global::CsvHelper.CsvReader.Create(input,
            new global::CsvHelper.CsvReaderOptions { FormatProvider = CultureInfo.InvariantCulture });
        foreach (CsvBenchmarkRow row in reader.GetRecords<CsvBenchmarkRow>()) yield return row;
    }

    private int Consume(IEnumerable<CsvBenchmarkRow> rows) {
        int checksum = 17, index = 0;
        foreach (CsvBenchmarkRow row in rows) {
            if (_validate) ValidateRow(index, row);
            checksum = AddChecksum(checksum, row);
            index++;
        }
        if (index != _rows.Length || checksum != _expectedChecksum)
            throw new InvalidDataException("Typed CSV read changed the row count or values.");
        return checksum;
    }

    private static int AddChecksum(int checksum, CsvBenchmarkRow row) {
        unchecked {
            checksum = checksum * 31 + row.Id;
            checksum = checksum * 31 + StringComparer.Ordinal.GetHashCode(row.Name);
            checksum = checksum * 31 + StringComparer.Ordinal.GetHashCode(row.Department);
            checksum = checksum * 31 + StringComparer.Ordinal.GetHashCode(row.Region);
            checksum = checksum * 31 + (row.IsEnabled ? 1 : 0);
            checksum = checksum * 31 + row.Created.GetHashCode();
            checksum = checksum * 31 + row.Score.GetHashCode();
            checksum = checksum * 31 + StringComparer.Ordinal.GetHashCode(row.Owner);
            checksum = checksum * 31 + row.TicketCount;
            return checksum * 31 + StringComparer.Ordinal.GetHashCode(row.Notes);
        }
    }

    private static void ValidateHeaders(byte[] output) {
        using var input = new MemoryStream(output, writable: false);
        using var reader = CsvDocument.OpenReader(input);
        if (reader.FieldCount != CsvBenchmarkData.Headers.Length ||
            !Enumerable.Range(0, reader.FieldCount).Select(reader.GetName).SequenceEqual(CsvBenchmarkData.Headers))
            throw new InvalidDataException("The typed writer changed header order.");
    }

    private void ValidateRow(int index, CsvBenchmarkRow row) {
        if ((uint)index >= (uint)_rows.Length) throw new InvalidDataException("Extra typed row.");
        CsvBenchmarkRow expected = _rows[index];
        if (row.Id != expected.Id || row.Name != expected.Name || row.Department != expected.Department ||
            row.Region != expected.Region || row.IsEnabled != expected.IsEnabled || row.Created != expected.Created ||
            row.Score != expected.Score || row.Owner != expected.Owner || row.TicketCount != expected.TicketCount || row.Notes != expected.Notes)
            throw new InvalidDataException($"Typed CSV changed row {index + 1}.");
    }
}
