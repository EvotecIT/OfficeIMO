using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Parser;
using ExcelReader.Core.Reader;
using ExcelReader.Core.Reader.Csv;
using OfficeIMO.CSV;
using CsvReaderOptions = OfficeIMO.CSV.CsvReaderOptions;
using OfficeIMO.Data;
using Sylvan.Data;
using Sylvan.Data.Csv;
using System.Data.Common;
using System.Globalization;
using nietras.SeparatedValues;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Full-file typed parallel reads, including source decoding for decoded-text routes.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("CsvParallelTyped")]
    public class CsvParallelTypedBenchmarks {
        private CsvParallelFixture _fixture = null!;
        private bool _validate;
        [ParamsSource(nameof(Cases))]
        public CsvParallelInput Input { get; set; }
        [ParamsSource(nameof(Degrees))]
        public int Dop { get; set; }
        public IEnumerable<CsvParallelInput> Cases() => CsvParallelFixture.Cases();
        public IEnumerable<int> Degrees() => CsvParallelFixture.Degrees();

        [GlobalSetup]
        public async Task SetupAsync() {
            _fixture = new(Input);
            CsvBenchmarkFixture.DescribeEngines();
            Console.WriteLine($"CSV parallel {Input}: bytes={_fixture.Bytes}, sha256={_fixture.Fingerprint}, DOP={Dop}, expectedSum={_fixture.ExpectedSum}; full-file input/decoding is timed.");
            _validate = true;
            try {
                await ExcelReaderParallel(); ExcelReaderSequential(); OfficeIMOCoreParallel(); SylvanSequential();
                SepParallel(); SepSequential();
#if OFFICEIMO_BENCHMARK_NEW_APIS
                OfficeIMOTextParallel();
#endif
            } catch { Cleanup(); throw; }
            finally { _validate = false; }
        }
        [GlobalCleanup]
        public void Cleanup() => _fixture?.Dispose();

        [Benchmark(Baseline = true)]
        public async Task<long> ExcelReaderParallel() {
            int count = 0;
            long sum = 0;
            SheetParallelOptions options = new SheetParallelOptions { DegreeOfParallelism = Dop };
            await using CsvReader workbook = ExcelReaderApi.FromCsvFile(_fixture.Path);
            if (_fixture.Heavy) {
                await foreach (CsvHeavyRecord row in SheetParallel.ParseAsync(workbook.FirstSheet, ExcelParser.FromAttributes<CsvHeavyRecord>(), options)) {
                    if (_validate) _fixture.Validate(row, count);
                    sum += row.Units; count++;
                }
            } else {
                await foreach (CsvNarrowRecord row in SheetParallel.ParseAsync(workbook.FirstSheet, ExcelParser.FromAttributes<CsvNarrowRecord>(), options)) {
                    if (_validate) _fixture.Validate(row, count);
                    sum += row.A; count++;
                }
            }
            return _fixture.Check(sum, count);
        }

        [Benchmark]
        public long ExcelReaderSequential() {
            using CsvReader workbook = ExcelReaderApi.FromCsvFile(_fixture.Path);
            return _fixture.Heavy
                ? Consume(ExcelParser.FromAttributes<CsvHeavyRecord>().Parse(workbook.FirstSheet))
                : Consume(ExcelParser.FromAttributes<CsvNarrowRecord>().Parse(workbook.FirstSheet));
        }

        [Benchmark]
        public long OfficeIMOCoreParallel() {
            using DbDataReader reader = CsvDocument.OpenReader(_fixture.Path, options: LoadOptions());
            ParallelRowMappingOptions options = ParallelOptions();
            return _fixture.Heavy ? Consume(reader.RowsAsParallel<CsvHeavyRecord>(options))
                : Consume(reader.RowsAsParallel<CsvNarrowRecord>(options));
        }

#if OFFICEIMO_BENCHMARK_NEW_APIS
        [Benchmark]
        public long OfficeIMOTextParallel() {
            // The public span API accepts decoded text. File I/O and UTF-8 decoding belong
            // to this operation, rather than giving one library a pre-decoded setup input.
            string text = File.ReadAllText(_fixture.Path);
            return _fixture.Heavy ? Consume(CsvDocument.ReadTextRowsAsParallel(text,
            HeavyFactory,
            options: LoadOptions(),
            parallelOptions: ParallelOptions()))
                : Consume(CsvDocument.ReadTextRowsAsParallel(text,
            NarrowFactory,
            options: LoadOptions(),
            parallelOptions: ParallelOptions()));
        }
#endif

        [Benchmark]
        public long SylvanSequential() {
            using StreamReader text = new StreamReader(_fixture.Path);
            using global::Sylvan.Data.Csv.CsvDataReader reader = global::Sylvan.Data.Csv.CsvDataReader.Create(text,
                new global::Sylvan.Data.Csv.CsvDataReaderOptions { Culture = CultureInfo.InvariantCulture });
            return _fixture.Heavy ? Consume(reader.GetRecords<CsvHeavyRecord>()) : Consume(reader.GetRecords<CsvNarrowRecord>());
        }

        [Benchmark]
        public long SepParallel() {
            using SepReader reader = Sep.New(',').Reader(options => options with {
                CultureInfo = CultureInfo.InvariantCulture, Unescape = true,
            }).FromFile(_fixture.Path);
            if (_fixture.Heavy)
                return Consume(Dop == 0 ? reader.ParallelEnumerate(ParseSepHeavy)
                    : reader.ParallelEnumerate(ParseSepHeavy, Dop));
            return Consume(Dop == 0 ? reader.ParallelEnumerate(ParseSepNarrow)
                : reader.ParallelEnumerate(ParseSepNarrow, Dop));
        }

        [Benchmark]
        public long SepSequential() {
            using SepReader reader = Sep.New(',').Reader(options => options with {
                CultureInfo = CultureInfo.InvariantCulture, Unescape = true,
            }).FromFile(_fixture.Path);
            long sum = 0;
            int count = 0;
            if (_fixture.Heavy) {
                foreach (SepReader.Row row in reader) {
                    CsvHeavyRecord record = ParseSepHeavy(row);
                    if (_validate) _fixture.Validate(record, count);
                    sum += record.Units;
                    count++;
                }
            } else {
                foreach (SepReader.Row row in reader) {
                    CsvNarrowRecord record = ParseSepNarrow(row);
                    if (_validate) _fixture.Validate(record, count);
                    sum += record.A;
                    count++;
                }
            }
            return _fixture.Check(sum, count);
        }

        private static CsvHeavyRecord ParseSepHeavy(SepReader.Row row) => new() {
            Region = row[0].ToString(), Country = row[1].ToString(), OrderDate = row[2].Parse<DateTime>(),
            UnitPrice = row[3].Parse<decimal>(), TotalRevenue = row[4].Parse<decimal>(), Units = row[5].Parse<int>(),
        };
        private static CsvNarrowRecord ParseSepNarrow(SepReader.Row row) => new() {
            A = row[0].Parse<int>(), B = row[1].Parse<int>(), C = row[2].Parse<int>(),
        };

        private long Consume(IEnumerable<CsvHeavyRecord> records) {
            long sum = 0;
            int count = 0;
            foreach (CsvHeavyRecord row in records) {
                if (_validate) _fixture.Validate(row, count);
                sum += row.Units; count++;
            }
            return _fixture.Check(sum, count);
        }
        private long Consume(IEnumerable<CsvNarrowRecord> records) {
            long sum = 0;
            int count = 0;
            foreach (CsvNarrowRecord row in records) {
                if (_validate) _fixture.Validate(row, count);
                sum += row.A; count++;
            }
            return _fixture.Check(sum, count);
        }

        private CsvReaderOptions LoadOptions() => new() { MaxInputBytes = Math.Max(CsvReaderOptions.DefaultMaxInputBytes, _fixture.Bytes) };
        private ParallelRowMappingOptions ParallelOptions() => new() { MaxDegreeOfParallelism = Dop == 0 ? null : Dop };

#if OFFICEIMO_BENCHMARK_NEW_APIS
        private static CsvRecordFactory<CsvHeavyRecord> HeavyFactory(CsvRecordHeader header) {
            int region = header.GetOrdinal("Region"), country = header.GetOrdinal("Country"), date = header.GetOrdinal("OrderDate"),
                price = header.GetOrdinal("UnitPrice"), revenue = header.GetOrdinal("TotalRevenue"), units = header.GetOrdinal("Units");
            return row => new CsvHeavyRecord { Region = row.GetString(region), Country = row.GetString(country), OrderDate = row.GetDateTime(date),
                UnitPrice = row.GetDecimal(price), TotalRevenue = row.GetDecimal(revenue), Units = row.GetInt32(units) };
        }
        private static CsvRecordFactory<CsvNarrowRecord> NarrowFactory(CsvRecordHeader header) {
            int a = header.GetOrdinal("A"), b = header.GetOrdinal("B"), c = header.GetOrdinal("C");
            return row => new CsvNarrowRecord { A = row.GetInt32(a), B = row.GetInt32(b), C = row.GetInt32(c) };
        }
#endif
    }
}
