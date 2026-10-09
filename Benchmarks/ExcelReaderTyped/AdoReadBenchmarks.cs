using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using OfficeIMO.Data;
using Sylvan.Data.Excel;
using System.Data;
using System.Globalization;
using System.Text;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;
using PeerDataReader = ExcelReader.Core.Reader.ExcelDataReader;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    public enum AdoAccess {
        GetValue,
        TypedGetters,
#if OFFICEIMO_BENCHMARK_NEW_APIS
        Utf8TextCopy,
#endif
    }

    /// <summary>Consumes the original typed worksheet through IDataReader interface dispatch.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("Ado")]
    public class AdoReadBenchmarks {
        private byte[] _workbook = [];
        private readonly byte[] _buffer = new byte[256];
        private long _expected;
        internal byte[] WorkbookBytes => _workbook;

        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;

        [ParamsSource(nameof(AccessModes))]
        public AdoAccess Access { get; set; }

        public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();
        public IEnumerable<AdoAccess> AccessModes() => Enum.GetValues<AdoAccess>();

        [GlobalSetup]
        public async Task SetupAsync() {
            BenchmarkInput.WriteDescription();
            _workbook = await TypedWorkbookFixture.CreateAsync(RowCount, "Original");
            TypedWorkbookFixture.ValidateShape(_workbook, RowCount, "Original");
            _expected = Access switch {
                AdoAccess.GetValue => (long)RowCount * 4,
#if OFFICEIMO_BENCHMARK_NEW_APIS
                AdoAccess.Utf8TextCopy => Enumerable.Range(1, RowCount).Sum(index => (long)Encoding.UTF8.GetByteCount(TypedWorkbookFixture.ExpectedRecord(index).Name!)),
#endif
                _ => Enumerable.Range(1, RowCount).Sum(index => TypedAccumulator(TypedWorkbookFixture.ExpectedRecord(index)))
            };
            using (IExcelWorkbook workbook = OpenPeer()) {
                using IDataReader reader = new PeerDataReader(workbook.FirstSheet);
                ValidateReader(reader, peerTextBytes: true);
            }
            using (MemoryStream stream = new MemoryStream(_workbook, writable: false))
            using (IDataReader reader = ExcelDocument.OpenDataReader(stream)) ValidateReader(reader, peerTextBytes: false);
            using (MemoryStream stream = new MemoryStream(_workbook, writable: false))
            using (IDataReader reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(stream, ExcelWorkbookType.ExcelXml))
                ValidateReader(reader, peerTextBytes: false);
            Check(ExcelReader());
            Check(OfficeIMO());
            Check(Sylvan());
        }

        [Benchmark(Baseline = true)]
        public long ExcelReader() {
            using IExcelWorkbook workbook = OpenPeer();
            using IDataReader reader = new PeerDataReader(workbook.FirstSheet);
            return Scan(reader, peerTextBytes: true);
        }

        [Benchmark]
        public long OfficeIMO() {
            using MemoryStream stream = new MemoryStream(_workbook, writable: false);
            using IDataReader reader = ExcelDocument.OpenDataReader(stream);
            return Scan(reader, peerTextBytes: false);
        }

        [Benchmark]
        public long Sylvan() {
            using MemoryStream stream = new MemoryStream(_workbook, writable: false);
            using IDataReader reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(stream, ExcelWorkbookType.ExcelXml);
            return Scan(reader, peerTextBytes: false);
        }

        private IExcelWorkbook OpenPeer() => ExcelReaderApi.FromXlsx(new MemoryStream(_workbook, writable: false), leaveOpen: false);

        private long Scan(IDataReader reader, bool peerTextBytes) {
            long accumulator = 0;
            int count = 0;
            while (reader.Read()) {
                if (Access == AdoAccess.GetValue) {
                    for (int column = 0; column < reader.FieldCount; column++)
                        accumulator += reader.GetValue(column) is null ? 0 : 1;
                }
#if OFFICEIMO_BENCHMARK_NEW_APIS
                else if (Access == AdoAccess.Utf8TextCopy) {
                    // ExcelReader defines text copying on GetBytes. OfficeIMO keeps the
                    // provider's binary GetBytes contract and exposes text copying explicitly.
                    accumulator += peerTextBytes
                        ? reader.GetBytes(0, 0, _buffer, 0, _buffer.Length)
                        : reader.GetUtf8Bytes(0, 0, _buffer, 0, _buffer.Length);
                }
#endif
                else {
                    accumulator = unchecked(accumulator + reader.GetString(0).Length + reader.GetInt32(1)
                        + reader.GetDateTime(2).Day + (long)reader.GetDouble(3));
                }
                count++;
            }
            if (count != RowCount) throw new InvalidDataException("ADO scan returned an incorrect row count.");
            return Check(accumulator);
        }

        private void ValidateReader(IDataReader reader, bool peerTextBytes) {
            if (reader.FieldCount != 4) throw new InvalidDataException("ADO input width differs.");
            for (int column = 0; column < 4; column++)
                if (reader.GetName(column) != TypedWorkbookFixture.Headers[column]) throw new InvalidDataException("ADO header differs.");
            int index = 0;
            while (reader.Read()) {
                TypedWorkbookFixture.ValidateRecord(new TypedRecord {
                    Name = reader.GetString(0), Id = reader.GetInt32(1),
                    Date = reader.GetDateTime(2), Value = reader.GetDouble(3)
                }, ++index);
                TypedRecord expected = TypedWorkbookFixture.ExpectedRecord(index);
                object rawDate = reader.GetValue(2);
                DateTime date = rawDate is DateTime nativeDate ? nativeDate
                    : rawDate is string text ? DateTime.Parse(text, CultureInfo.InvariantCulture)
                    : DateTime.FromOADate(Convert.ToDouble(rawDate, CultureInfo.InvariantCulture));
                if (!Equals(reader.GetValue(0), expected.Name) || Convert.ToInt32(reader.GetValue(1), CultureInfo.InvariantCulture) != expected.Id
                    || date != expected.Date || Convert.ToDouble(reader.GetValue(3), CultureInfo.InvariantCulture) != expected.Value)
                    throw new InvalidDataException("ADO GetValue fields differ.");
#if OFFICEIMO_BENCHMARK_NEW_APIS
                if (Access == AdoAccess.Utf8TextCopy) {
                    byte[] expectedUtf8 = Encoding.UTF8.GetBytes(expected.Name!);
                    _buffer.AsSpan().Fill(0xA5);
                    long copied = peerTextBytes
                        ? reader.GetBytes(0, 0, _buffer, 0, _buffer.Length)
                        : reader.GetUtf8Bytes(0, 0, _buffer, 0, _buffer.Length);
                    if (copied != expectedUtf8.Length || !_buffer.AsSpan(0, expectedUtf8.Length).SequenceEqual(expectedUtf8))
                        throw new InvalidDataException($"ADO UTF8 copy bytes/count differ at row {index}.");
                }
#endif
            }
            if (index != RowCount || reader.NextResult()) throw new InvalidDataException("ADO shape differs.");
        }

        private static long TypedAccumulator(TypedRecord record) => record.Name!.Length + record.Id + record.Date.Day + (long)record.Value;
        private long Check(long result) => result == _expected ? result : throw new InvalidDataException("ADO accumulator differs.");
    }
}
