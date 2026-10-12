#if OFFICEIMO_BENCHMARK_NEW_APIS
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using OfficeIMO.CSV;
using OfficeIMO.Data;
using System.Globalization;
using System.Text;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Asynchronous scalar reduction over borrowed text through each public API.</summary>
    /// <remarks>
    /// OfficeIMO consumes decoded UTF-16 spans in a synchronous callback at DOP 1;
    /// the peer consumes UTF-8 spans in its native asynchronous row loop. The timed
    /// fixture has ASCII names, so both length reductions have the same meaning.
    /// Headerless reader initialization and canonical fallback may materialize text.
    /// </remarks>
    [MemoryDiagnoser]
    [BenchmarkCategory("CsvRawAsyncBorrowed")]
    public class CsvBorrowedAsyncReadBenchmarks {
        private byte[] _data = [];
        private long _expected;

        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 50_000;
        public IEnumerable<int> RowCounts() => new TypedReadBenchmarks().RowCounts();

        [GlobalSetup]
        public async Task SetupAsync() {
            _data = CsvBenchmarkFixture.Create(RowCount, headers: false);
            _expected = TypedWorkbookFixture.ExpectedChecksum(RowCount);
            CsvBenchmarkFixture.Describe(_data, RowCount, "raw async borrowed; OfficeIMO UTF16 callback DOP1; peer UTF8 loop; ASCII names");
            await QualifyAsync(_data, RowCount, TypedWorkbookFixture.ExpectedRecord);
            (byte[] control, TypedRecord[] records) = CreateNameControl();
            await QualifyAsync(control, records.Length, index => records[index - 1]);
            await ExcelReaderBorrowedAsync();
            await OfficeIMOBorrowedAggregateAsync();
            Console.WriteLine($"CSV borrowed async qualification passed: rows={RowCount}; Unicode/quoted/multiline controlRows={records.Length}; every name byte/span, scalar and row order validated. "
                + "OfficeIMO borrows UTF16; peer borrows UTF8; headerless opening may materialize the first row; MemoryStream does not establish async disk I/O.");
        }

        [Benchmark(Baseline = true)]
        public async Task<long> ExcelReaderBorrowedAsync() {
            await using MemoryStream stream = new MemoryStream(_data, writable: false);
            await using CsvReader workbook = ExcelReaderApi.FromCsv(stream);
            await using CsvReader.Enumerator rows = workbook.FirstSheet.GetAsyncEnumerator();
            AggregateState state = default;
            while (await rows.MoveNextAsync()) {
                Row row = rows.Current;
                if (!row[1].TryParse(CultureInfo.InvariantCulture, out int id)
                    || !row[2].TryGetDateTime(out DateTime date)
                    || !row[3].TryParse(CultureInfo.InvariantCulture, out double value))
                    throw new InvalidDataException("Borrowed CSV scalar parsing failed.");
                state.Sum = unchecked(state.Sum + row[0].Value.Length + id + date.Ticks + (long)value);
                state.Count++;
            }
            return CsvBenchmarkFixture.Check(state.Sum, state.Count, RowCount, _expected);
        }

        [Benchmark]
        public async Task<long> OfficeIMOBorrowedAggregateAsync() {
            await using MemoryStream stream = new MemoryStream(_data, writable: false);
            AggregateState state = await CsvDocument.AggregateRowsAsParallelAsync(
                stream, static () => default(AggregateState), static _ => AccumulateOffice,
                static (left, right) => new AggregateState { Count = left.Count + right.Count, Sum = unchecked(left.Sum + right.Sum) },
                loadOptions: new CsvLoadOptions { HasHeaderRow = false },
                parallelOptions: new ParallelRowMappingOptions { MaxDegreeOfParallelism = 1 });
            return CsvBenchmarkFixture.Check(state.Sum, state.Count, RowCount, _expected);
        }

        private static void AccumulateOffice(ref AggregateState state, CsvRecord row) {
            state.Sum = unchecked(state.Sum + row.GetSpan(0).Length + row.GetInt32(1)
                + row.GetDateTime(2).Ticks + (long)row.GetDouble(3));
            state.Count++;
        }

        private static async Task QualifyAsync(byte[] data, int rowCount, Func<int, TypedRecord> expectedRecord) {
            using (MemoryStream stream = new MemoryStream(data, writable: false)) {
                await using CsvReader workbook = ExcelReaderApi.FromCsv(stream);
                await using CsvReader.Enumerator rows = workbook.FirstSheet.GetAsyncEnumerator();
                int count = 0;
                while (await rows.MoveNextAsync()) {
                    if (count >= rowCount) throw new InvalidDataException("Peer borrowed CSV has extra rows.");
                    Row row = rows.Current;
                    TypedRecord expected = expectedRecord(count + 1);
                    if (row.ColumnCount != 4 || !row[0].Value.SequenceEqual(Encoding.UTF8.GetBytes(expected.Name!))
                        || !row[1].TryParse(CultureInfo.InvariantCulture, out int id) || id != expected.Id
                        || !row[2].TryGetDateTime(out DateTime date) || date != expected.Date
                        || !row[3].TryParse(CultureInfo.InvariantCulture, out double value) || value != expected.Value)
                        throw new InvalidDataException($"Peer borrowed CSV field/order differs at row {count + 1}.");
                    count++;
                }
                if (count != rowCount) throw new InvalidDataException("Peer borrowed CSV row count differs.");
            }

            using (MemoryStream stream = new MemoryStream(data, writable: false)) {
                AggregateState state = await CsvDocument.AggregateRowsAsParallelAsync(
                    stream, static () => default(AggregateState), header => {
                        if (header.Count != 4) throw new InvalidDataException("OfficeIMO borrowed CSV width differs.");
                        return (ref AggregateState result, CsvRecord row) => {
                            if (result.Count >= rowCount) throw new InvalidDataException("OfficeIMO borrowed CSV has extra rows.");
                            TypedRecord expected = expectedRecord(result.Count + 1);
                            if (row.FieldCount != 4 || !row.GetSpan(0).SequenceEqual(expected.Name.AsSpan())
                                || row.GetInt32(1) != expected.Id || row.GetDateTime(2) != expected.Date
                                || row.GetDouble(3) != expected.Value)
                                throw new InvalidDataException($"OfficeIMO borrowed CSV field/order differs at row {result.Count + 1}.");
                            result.Count++;
                        };
                    }, static (left, right) => new AggregateState { Count = left.Count + right.Count },
                    loadOptions: new CsvLoadOptions { HasHeaderRow = false },
                    parallelOptions: new ParallelRowMappingOptions { MaxDegreeOfParallelism = 1 });
                if (state.Count != rowCount) throw new InvalidDataException("OfficeIMO borrowed CSV row count differs.");
            }
            BenchmarkInput.WriteFixtureIdentity($"read/CsvBorrowedAsync/qualifiedRows={rowCount}", data);
        }

        private static (byte[] Data, TypedRecord[] Records) CreateNameControl() {
            string[] names = ["Zażółć 😀", "comma,value", "quote\"value", "line\r\nvalue", "line\nvalue", ""];
            TypedRecord[] records = new TypedRecord[names.Length];
            StringBuilder csv = new StringBuilder();
            for (int index = 0; index < names.Length; index++) {
                TypedRecord record = TypedWorkbookFixture.ExpectedRecord(index + 1);
                record.Name = names[index];
                records[index] = record;
                csv.Append('"').Append(names[index].Replace("\"", "\"\"")).Append("\",")
                    .Append(record.Id.ToString(CultureInfo.InvariantCulture)).Append(',')
                    .Append(record.Date.ToString("O", CultureInfo.InvariantCulture)).Append(',')
                    .Append(record.Value.ToString(CultureInfo.InvariantCulture)).Append('\n');
            }
            return (Encoding.UTF8.GetBytes(csv.ToString()), records);
        }

        private struct AggregateState {
            internal int Count;
            internal long Sum;
        }
    }
}
#endif
