#if OFFICEIMO_BENCHMARK_NEW_APIS
using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>Compares complete, authenticated field scans on identical ciphertext.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("EncryptedVerifiedFullFields")]
    public class EncryptedVerifiedReadBenchmarks {
        private readonly EncryptedReadWorkload _workload = new();
        [Params(EncryptedWorkbookInput.OriginalSmallXlsx, EncryptedWorkbookInput.GeneratedLargeXlsx)]
        public EncryptedWorkbookInput Input { get; set; }
        [Params(false, true)]
        public bool MemoryInput { get; set; }
        [GlobalSetup]
        public Task SetupAsync() => _workload.SetupAsync(EncryptedWorkbookFixture.Load(Input));
        [Benchmark(Baseline = true)]
        public long ExcelReaderVerifiedFullFields() => _workload.Peer(MemoryInput);
        [Benchmark]
        public long OfficeIMOVerifiedFullFields() => _workload.Office(MemoryInput);
    }

    /// <summary>Includes each engine's public asynchronous stream opening and complete authenticated scan.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("EncryptedVerifiedStreamAsync")]
    public class EncryptedVerifiedAsyncReadBenchmarks {
        private readonly EncryptedReadWorkload _workload = new();
        [Params(EncryptedWorkbookInput.OriginalSmallXlsx, EncryptedWorkbookInput.GeneratedLargeXlsx)]
        public EncryptedWorkbookInput Input { get; set; }
        [GlobalSetup]
        public Task SetupAsync() => _workload.SetupAsync(EncryptedWorkbookFixture.Load(Input));
        [Benchmark(Baseline = true)]
        public Task<long> ExcelReaderVerifiedStreamAsync() => _workload.PeerAsync();
        [Benchmark]
        public Task<long> OfficeIMOVerifiedStreamAsync() => _workload.OfficeAsync();
    }

    /// <summary>Preserves the exact original small-input operations; these do not consume field values.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("EncryptedOriginalSmallDiagnostic")]
    public class EncryptedOriginalSmallDiagnostics {
        private EncryptedWorkbookFixture _fixture = null!;
        [GlobalSetup]
        public async Task SetupAsync() {
            _fixture = EncryptedWorkbookFixture.Load(EncryptedWorkbookInput.OriginalSmallXlsx);
            await new EncryptedReadWorkload().SetupAsync(_fixture);
            if (Plain_Stream() != 4 || Encrypted_Stream() != 4 || Encrypted_Stream_VerifyIntegrity() != 4
                || Encrypted_Memory() != 4 || Encrypted_OpenOnly() != 1)
                throw new InvalidDataException("Original small diagnostic observation differs.");
        }
        [Benchmark]
        public long Plain_Stream() {
            using Stream stream = _fixture.OpenPlainStream();
            using IExcelWorkbook workbook = ExcelReaderApi.Open(stream, leaveOpen: true);
            return Widths(workbook.FirstSheet);
        }
        [Benchmark]
        public long Encrypted_Stream() {
            using Stream stream = _fixture.OpenEncryptedStream();
            using IExcelWorkbook workbook = ExcelReaderApi.Open(stream, leaveOpen: true, EncryptedWorkbookFixture.PeerOptions(false));
            return Widths(workbook.FirstSheet);
        }
        [Benchmark]
        public long Encrypted_Stream_VerifyIntegrity() {
            using Stream stream = _fixture.OpenEncryptedStream();
            using IExcelWorkbook workbook = ExcelReaderApi.Open(stream, leaveOpen: true, EncryptedWorkbookFixture.PeerOptions());
            return Widths(workbook.FirstSheet);
        }
        [Benchmark]
        public long Encrypted_Memory() {
            // Public memory opening always authenticates, even when the option is false.
            using IExcelWorkbook workbook = ExcelReaderApi.Open(_fixture.Encrypted.AsMemory(), EncryptedWorkbookFixture.PeerOptions(false));
            return Widths(workbook.FirstSheet);
        }
        [Benchmark]
        public int Encrypted_OpenOnly() {
            using MemoryStream stream = new MemoryStream(_fixture.Encrypted, writable: false);
            using IExcelWorkbook workbook = ExcelReaderApi.Open(stream, leaveOpen: true, EncryptedWorkbookFixture.PeerOptions(false));
            return workbook.SheetCount;
        }
        internal static long Widths(IExcelSheet sheet) {
            long sum = 0;
            using IExcelRowEnumerator rows = sheet.GetEnumerator();
            while (rows.MoveNext()) sum += rows.Current.ColumnCount;
            return sum;
        }
    }

    /// <summary>Preserves the upstream large row-width and lazy open-only methods without paired speed ratios.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("EncryptedOriginalLargeDiagnostic")]
    public class EncryptedOriginalLargeDiagnostics {
        private EncryptedWorkbookFixture _fixture = null!;
        [GlobalSetup]
        public async Task SetupAsync() {
            _fixture = EncryptedWorkbookFixture.Load(EncryptedWorkbookInput.GeneratedLargeXlsx);
            await new EncryptedReadWorkload().SetupAsync(_fixture);
            if (Large_Plain_Stream() != 4L * _fixture.Rows || Large_Encrypted_Stream() != 4L * _fixture.Rows
                || Large_Encrypted_OpenOnly() != 1)
                throw new InvalidDataException("Original large diagnostic observation differs.");
        }
        [Benchmark]
        public long Large_Plain_Stream() {
            using MemoryStream stream = new MemoryStream(_fixture.Plain, writable: false);
            using IExcelWorkbook workbook = ExcelReaderApi.Open(stream, leaveOpen: true);
            return EncryptedOriginalSmallDiagnostics.Widths(workbook.FirstSheet);
        }
        [Benchmark]
        public long Large_Encrypted_Stream() {
            using MemoryStream stream = new MemoryStream(_fixture.Encrypted, writable: false);
            using IExcelWorkbook workbook = ExcelReaderApi.Open(stream, leaveOpen: true, EncryptedWorkbookFixture.PeerOptions(false));
            return EncryptedOriginalSmallDiagnostics.Widths(workbook.FirstSheet);
        }
        [Benchmark]
        public int Large_Encrypted_OpenOnly() {
            using MemoryStream stream = new MemoryStream(_fixture.Encrypted, writable: false);
            using IExcelWorkbook workbook = ExcelReaderApi.Open(stream, leaveOpen: true, EncryptedWorkbookFixture.PeerOptions(false));
            return workbook.SheetCount;
        }
    }
}
#endif