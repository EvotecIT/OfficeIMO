using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;
using OfficeIMO.CSV;
using OfficeIMO.Data;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;
using SylvanCsv = Sylvan.Data.Csv.CsvDataReader;
using SylvanOptions = Sylvan.Data.Csv.CsvDataReaderOptions;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks;

/// <summary>The pinned real CSV fields, including the header, through the actual span APIs.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("CsvRealDataOriginal")]
public class CsvRealDataReadBenchmarks {
    private readonly CsvRealDataWorkload _workload = new();
    [GlobalSetup]
    public void Setup() => _workload.Setup();
    [Benchmark(Baseline = true)]
    public long ExcelReaderStream() => _workload.Peer(memory: false, materialize: false);
    [Benchmark]
    public long ExcelReaderMemory() => _workload.Peer(memory: true, materialize: false);
#if OFFICEIMO_BENCHMARK_NEW_APIS
    [Benchmark]
    public long OfficeIMOStreamBorrowedUtf8() => _workload.Office(materialize: false);
#endif
    [Benchmark]
    public long SylvanFieldSpan() => _workload.Sylvan(materialize: false);
}

/// <summary>Separates matched string materialization from borrowed real CSV reads.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("CsvRealDataMaterialized")]
public class CsvRealDataMaterializedReadBenchmarks {
    private readonly CsvRealDataWorkload _workload = new();
    [GlobalSetup]
    public void Setup() => _workload.Setup();
    [Benchmark(Baseline = true)]
    public long ExcelReaderMaterialized() => _workload.Peer(memory: false, materialize: true);
    [Benchmark]
    public long OfficeIMOMaterialized() => _workload.Office(materialize: true);
    [Benchmark]
    public long SylvanMaterialized() => _workload.Sylvan(materialize: true);
}

internal sealed class CsvRealDataWorkload {
    private byte[] _bytes = [];
    private long _byteCount, _characterCount;
    private const int Rows = MarkPflug65KFixture.ExpectedRows + 1;
    private const int Columns = MarkPflug65KFixture.ExpectedColumns;
    internal void Setup() {
        BenchmarkInput.WriteDescription();
        _byteCount = _characterCount = 0;
        MarkPflug65KFixture.EnsureAuthentic(MarkPflug65KFixture.CsvFileName,
            MarkPflug65KFixture.GetHashes()[MarkPflug65KFixture.CsvFileName], new Uri(
                "https://raw.githubusercontent.com/GabrielMarquezMatte/ExcelReader/ca5b50f99e8ef57ab476f0a2bc8043558d58b28d/tests/ExcelReader.Benchmarks/Data/65K_Records_Data.csv"));
        _bytes = File.ReadAllBytes(MarkPflug65KFixture.CsvPath);
        using var oracleStream = new MemoryStream(_bytes, false);
        using var oracleText = new StreamReader(oracleStream);
        using var oracle = SylvanCsv.Create(oracleText, Options());
        using var peer = ExcelReaderApi.FromCsv(_bytes.AsMemory());
        using var peerRows = peer.FirstSheet.GetEnumerator();
        using var peerStream = new MemoryStream(_bytes, false);
        using var streamedPeer = ExcelReaderApi.FromCsv(peerStream);
        using var streamedPeerRows = streamedPeer.FirstSheet.GetEnumerator();
        using var officeStream = new MemoryStream(_bytes, false);
        using var office = CsvDocument.OpenDataReader(officeStream, new CsvLoadOptions { HasHeaderRow = false });
        int rows = 0;
        while (oracle.Read()) {
            if (!peerRows.MoveNext() || !streamedPeerRows.MoveNext() || !office.Read() || oracle.FieldCount != Columns
                || peerRows.Current.ColumnCount != Columns || streamedPeerRows.Current.ColumnCount != Columns
                || office.FieldCount != Columns)
                throw new InvalidDataException("Real CSV row count/width differs.");
            for (int column = 0; column < Columns; column++) {
                string text = oracle.GetString(column);
                if (peerRows.Current[column].GetString() != text || streamedPeerRows.Current[column].GetString() != text
                    || office.GetString(column) != text || office.GetFieldType(column) != typeof(string)
                    || oracle.GetFieldType(column) != typeof(string)
                    || !oracle.GetFieldSpan(column).SequenceEqual(text))
                    throw new InvalidDataException($"Real CSV text/order differs at {rows + 1}/{column + 1}.");
                byte[] utf8 = Encoding.UTF8.GetBytes(text);
                if (!peerRows.Current[column].Value.SequenceEqual(utf8)
                    || !streamedPeerRows.Current[column].Value.SequenceEqual(utf8))
                    throw new InvalidDataException("Peer real CSV borrowed bytes differ.");
#if OFFICEIMO_BENCHMARK_NEW_APIS
                if (!office.TryGetUtf8Text(column, out var borrowed) || !borrowed.SequenceEqual(utf8))
                    throw new InvalidDataException("OfficeIMO real CSV borrowed bytes differ.");
#endif
                _byteCount += utf8.Length;
                _characterCount += text.Length;
            }
            rows++;
        }
        if (rows != Rows || peerRows.MoveNext() || streamedPeerRows.MoveNext() || office.Read()
            || oracle.NextResult() || office.NextResult())
            throw new InvalidDataException("Real CSV row/sheet count differs.");
        Peer(memory: false, materialize: false);
        Peer(memory: true, materialize: false);
        Peer(memory: false, materialize: true);
        Office(materialize: true);
        Sylvan(materialize: false);
        Sylvan(materialize: true);
#if OFFICEIMO_BENCHMARK_NEW_APIS
        Office(materialize: false);
#endif
        Console.WriteLine($"Qualified pinned real CSV: rowsIncludingHeader={rows}; columns={Columns}; bytes={_bytes.Length}; "
            + $"SHA256={Convert.ToHexString(SHA256.HashData(_bytes))}; UTF8FieldBytes={_byteCount}; "
            + $"UTF16FieldCharacters={_characterCount}; every field/type/order/span validated. "
            + "Peer memory opens ReadOnlyMemory directly; OfficeIMO and Sylvan open memory-backed streams.");
    }
    internal long Peer(bool memory, bool materialize) {
        using var stream = memory ? null : new MemoryStream(_bytes, false);
        using var reader = memory ? ExcelReaderApi.FromCsv(_bytes.AsMemory()) : ExcelReaderApi.FromCsv(stream!);
        int rows = 0;
        long sum = 0;
        foreach (var row in reader.FirstSheet) {
            foreach (var cell in row.Cells) sum += materialize ? cell.Value.GetString().Length : cell.Value.Value.Length;
            rows++;
        }
        return Check(sum, rows, materialize ? _characterCount : _byteCount);
    }
    internal long Office(bool materialize) {
        using var stream = new MemoryStream(_bytes, false);
        using var reader = CsvDocument.OpenDataReader(stream, new CsvLoadOptions { HasHeaderRow = false });
        long sum = 0;
        int rows = 0;
        while (reader.Read()) {
            for (int column = 0; column < reader.FieldCount; column++) {
                if (materialize) sum += reader.GetString(column).Length;
#if OFFICEIMO_BENCHMARK_NEW_APIS
                else if (reader.TryGetUtf8Text(column, out var text)) sum += text.Length;
                else throw new InvalidDataException("The real CSV field must supply borrowed text.");
#else
                else throw new InvalidOperationException("Borrowed real CSV requires the current-API flag.");
#endif
            }
            rows++;
        }
        return Check(sum, rows, materialize ? _characterCount : _byteCount);
    }
    internal long Sylvan(bool materialize) {
        using var stream = new MemoryStream(_bytes, false);
        using var text = new StreamReader(stream);
        using var reader = SylvanCsv.Create(text, Options());
        long sum = 0;
        int rows = 0;
        while (reader.Read()) {
            for (int column = 0; column < reader.FieldCount; column++)
                sum += materialize ? reader.GetString(column).Length : reader.GetFieldSpan(column).Length;
            rows++;
        }
        return Check(sum, rows, _characterCount);
    }
    private static SylvanOptions Options() => new() { HasHeaders = false, Culture = CultureInfo.InvariantCulture };
    private static long Check(long sum, int rows, long expected) => rows == Rows && sum == expected
        ? sum : throw new InvalidDataException("Real CSV checksum/count differs.");
}
