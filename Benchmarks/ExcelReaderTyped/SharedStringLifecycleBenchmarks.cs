using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Reader;
using OfficeIMO.Data;
using Sylvan.Data.Excel;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.ReaderComparison.Benchmarks {
    /// <summary>The two ZIP storage policies in the upstream SST lifecycle diagnostics.</summary>
    public enum SharedStringZipStorage { Deflated, Stored }

    /// <summary>Shares byte-identical inflated parts and complete value proof between SST stage cases.</summary>
    internal sealed class SharedStringLifecycleWorkload {
        private byte[] _bytes = [];
        private readonly WorkbookScanWorkload _scan = new();
        private int _rows;
        private ComparisonWorkbookFormat _format;
        private long _headerChecksum;
#if OFFICEIMO_BENCHMARK_NEW_APIS
        private long _borrowedChecksum;
#endif

        internal async Task SetupAsync(int rows, SharedStringZipStorage storage,
            ComparisonWorkbookFormat format = ComparisonWorkbookFormat.Xlsx) {
            BenchmarkInput.WriteDescription();
            _format = format;
            string? configured = format == ComparisonWorkbookFormat.Xlsx
                ? Environment.GetEnvironmentVariable("OFFICEIMO_SHARED_XLSX_FIXTURE") : null;
            byte[] deflated = configured != null ? File.ReadAllBytes(configured) : await (format switch {
                ComparisonWorkbookFormat.Xlsx => StringHeavyWorkbookGenerator.BuildXlsxAsync(rows),
                ComparisonWorkbookFormat.Xlsb => StringHeavyWorkbookGenerator.BuildXlsbAsync(rows),
                _ => throw new ArgumentOutOfRangeException(nameof(format)),
            });
            if (configured != null) WrittenWorkbookValidation.ValidateSavedXlsxRowCount(deflated, rows, configured);
            _bytes = storage == SharedStringZipStorage.Stored ? Store(deflated) : deflated;
            _rows = rows + 1;
            if (storage == SharedStringZipStorage.Stored) ValidateInflatedParts(deflated, _bytes);
            using (MemoryStream packageStream = new MemoryStream(_bytes, writable: false)) {
                using ZipArchive package = new ZipArchive(packageStream, ZipArchiveMode.Read);
                if (storage == SharedStringZipStorage.Deflated) WrittenWorkbookValidation.ValidateEntries(package);
                WrittenWorkbookValidation.ValidateRelationships(package);
            }
            object[] headers = StringHeavyWorkbookGenerator.Headers.Cast<object>().ToArray();
            _scan.Setup(_bytes, format, rows + 1, headers.Length,
                index => index == 0 ? headers : StringHeavyWorkbookGenerator.ExpectedValues(index));
            _headerChecksum = StringHeavyWorkbookGenerator.Headers.Sum(static text => (long)text.Length);
            if (!ExcelReaderFirstRow() || !OfficeIMOFirstRow() || !SylvanFirstRow())
                throw new InvalidDataException("The SST lifecycle case did not produce its first header row.");
            ValidateMaterializedHeaders();
            Console.WriteLine($"Qualified shared-string lifecycle {format}/{storage}: rows={rows + 1}, bytes={_bytes.Length}; "
                + $"SHA256={Convert.ToHexString(SHA256.HashData(_bytes))}; "
                + "stored variant changes compression of every ZIP part, preserving all inflated bytes.");
        }

        internal long ExcelReaderFullMaterialized() => _scan.ExcelReader(materializeStrings: true);
        internal long OfficeIMOFullMaterialized() => _scan.OfficeIMO();
        internal long SylvanFullMaterialized() => _scan.Sylvan();
        internal long ExcelReaderFullOriginal(bool prefetch = false) => _scan.ExcelReader(prefetch: prefetch);

#if OFFICEIMO_BENCHMARK_NEW_APIS
        internal async Task SetupBorrowedAsync(int rows, SharedStringZipStorage storage, ComparisonWorkbookFormat format) {
            await SetupAsync(rows, storage, format);
            using MemoryStream stream = new MemoryStream(_bytes, writable: false);
            using IExcelWorkbook peer = OpenPeer(stream);
            using IExcelRowEnumerator peerRows = peer.FirstSheet.GetEnumerator();
            using ExcelWorkbookDataReader office = ExcelDocument.OpenDataReader(_bytes, new ExcelReadOptions { HasHeaderRow = false });
            int count = 0;
            long checksum = 0;
            while (office.Read()) {
                if (!peerRows.MoveNext() || office.FieldCount != StringHeavyWorkbookGenerator.Headers.Length)
                    throw new InvalidDataException("Borrowed SST row shape differs.");
                object[] expected = count == 0 ? StringHeavyWorkbookGenerator.Headers.Cast<object>().ToArray()
                    : StringHeavyWorkbookGenerator.ExpectedValues(count);
                for (int column = 0; column < expected.Length; column++) {
                    if (!Equals(office.GetValue(column), expected[column]))
                        throw new InvalidDataException("Borrowed SST native field/type differs.");
                    if (expected[column] is string text) {
                        byte[] utf8 = Encoding.UTF8.GetBytes(text);
                        if (!office.TryGetUtf8Text(column, out ReadOnlySpan<byte> borrowed) || !borrowed.SequenceEqual(utf8)
                            || !peerRows.Current[column].Value.SequenceEqual(utf8))
                            throw new InvalidDataException($"Borrowed SST bytes differ at row {count + 1}, column {column + 1}.");
                        checksum = unchecked(checksum + utf8.Length);
                    } else if (expected[column] is double number) checksum = unchecked(checksum + (long)number);
                    else if (expected[column] is DateTime date) checksum = unchecked(checksum + date.Ticks);
                    else throw new InvalidDataException("Unexpected SST fixture field type.");
                }
                count++;
            }
            if (count != _rows || peerRows.MoveNext() || office.NextResult())
                throw new InvalidDataException("Borrowed SST row/result count differs.");
            _borrowedChecksum = checksum;
            if (OfficeIMOFullBorrowed() != checksum || ExcelReaderFullOriginal() != checksum
                || ExcelReaderFullOriginal(prefetch: true) != checksum)
                throw new InvalidDataException("Borrowed SST full scan checksum differs.");
            Console.WriteLine($"Qualified public shared-string UTF8 {format}/{storage}: rows={_rows}, checksum={checksum}; "
                + "every borrowed byte/native field verified. Every measured scan opens a new reader; "
                + "OfficeIMO normalized-string UTF8 encoding and its bounded reader-owned cache are included.");
        }

        internal long OfficeIMOFullBorrowed() {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(_bytes, new ExcelReadOptions { HasHeaderRow = false });
            long sum = 0;
            int rows = 0;
            while (reader.Read()) {
                for (int column = 0; column < reader.FieldCount; column++) {
                    if (reader.TryGetUtf8Text(column, out ReadOnlySpan<byte> text)) sum = unchecked(sum + text.Length);
                    else switch (reader.GetValue(column)) {
                        case double number: sum = unchecked(sum + (long)number); break;
                        case DateTime date: sum = unchecked(sum + date.Ticks); break;
                        default: throw new InvalidDataException("The SST text field must provide borrowed UTF8.");
                    }
                }
                rows++;
            }
            return rows == _rows && sum == _borrowedChecksum ? sum
                : throw new InvalidDataException("Borrowed SST field checksum/count differs.");
        }
#endif

        internal bool ExcelReaderFirstRow() {
            using MemoryStream stream = new MemoryStream(_bytes, writable: false);
            using IExcelWorkbook workbook = OpenPeer(stream);
            using IExcelRowEnumerator rows = workbook.FirstSheet.GetEnumerator();
            return rows.MoveNext();
        }

        internal bool OfficeIMOFirstRow() {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(_bytes, new ExcelReadOptions { HasHeaderRow = false });
            return reader.Read();
        }

        internal bool SylvanFirstRow() {
            using MemoryStream stream = new MemoryStream(_bytes, writable: false);
            using Sylvan.Data.Excel.ExcelDataReader reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(stream,
                _format == ComparisonWorkbookFormat.Xlsx ? ExcelWorkbookType.ExcelXml : ExcelWorkbookType.ExcelBinary,
                new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders });
            return reader.Read();
        }

        internal long ExcelReaderMaterializedHeaders() {
            using MemoryStream stream = new MemoryStream(_bytes, writable: false);
            using IExcelWorkbook workbook = OpenPeer(stream);
            using IExcelRowEnumerator rows = workbook.FirstSheet.GetEnumerator();
            if (!rows.MoveNext()) throw new InvalidDataException("The SST header row is missing.");
            long checksum = 0;
            for (int column = 0; column < StringHeavyWorkbookGenerator.Headers.Length; column++)
                checksum += rows.Current[column].GetString().Length;
            return ValidateHeaderChecksum(checksum);
        }

        internal long OfficeIMOMaterializedHeaders() {
            using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(_bytes, new ExcelReadOptions { HasHeaderRow = false });
            if (!reader.Read() || reader.FieldCount != StringHeavyWorkbookGenerator.Headers.Length)
                throw new InvalidDataException("The SST header row shape differs.");
            long checksum = 0;
            for (int column = 0; column < reader.FieldCount; column++) checksum += reader.GetString(column).Length;
            return ValidateHeaderChecksum(checksum);
        }

        internal long SylvanMaterializedHeaders() {
            using MemoryStream stream = new MemoryStream(_bytes, writable: false);
            using Sylvan.Data.Excel.ExcelDataReader reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(stream,
                _format == ComparisonWorkbookFormat.Xlsx ? ExcelWorkbookType.ExcelXml : ExcelWorkbookType.ExcelBinary,
                new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders });
            if (!reader.Read() || reader.FieldCount != StringHeavyWorkbookGenerator.Headers.Length)
                throw new InvalidDataException("The SST header row shape differs.");
            long checksum = 0;
            for (int column = 0; column < reader.FieldCount; column++) checksum += reader.GetString(column).Length;
            return ValidateHeaderChecksum(checksum);
        }

        private long ValidateHeaderChecksum(long checksum) => checksum == _headerChecksum ? checksum
            : throw new InvalidDataException("The SST materialized header checksum differs.");

        private void ValidateMaterializedHeaders() {
            using MemoryStream stream = new MemoryStream(_bytes, writable: false);
            using IExcelWorkbook peer = OpenPeer(stream);
            using IExcelRowEnumerator peerRows = peer.FirstSheet.GetEnumerator();
            using ExcelWorkbookDataReader office = ExcelDocument.OpenDataReader(_bytes, new ExcelReadOptions { HasHeaderRow = false });
            using MemoryStream sylvanStream = new MemoryStream(_bytes, writable: false);
            using Sylvan.Data.Excel.ExcelDataReader sylvan = global::Sylvan.Data.Excel.ExcelDataReader.Create(sylvanStream,
                _format == ComparisonWorkbookFormat.Xlsx ? ExcelWorkbookType.ExcelXml : ExcelWorkbookType.ExcelBinary,
                new ExcelDataReaderOptions { Schema = ExcelSchema.NoHeaders });
            if (!peerRows.MoveNext() || !office.Read() || !sylvan.Read()
                || office.FieldCount != StringHeavyWorkbookGenerator.Headers.Length || sylvan.FieldCount != office.FieldCount)
                throw new InvalidDataException("The SST materialized header row shape differs.");
            for (int column = 0; column < office.FieldCount; column++) {
                string expected = StringHeavyWorkbookGenerator.Headers[column];
                if (peerRows.Current[column].GetString() != expected || office.GetString(column) != expected || sylvan.GetString(column) != expected)
                    throw new InvalidDataException($"The SST materialized header differs at column {column + 1}.");
            }
            if (ExcelReaderMaterializedHeaders() != _headerChecksum || OfficeIMOMaterializedHeaders() != _headerChecksum
                || SylvanMaterializedHeaders() != _headerChecksum)
                throw new InvalidDataException("The SST materialized header operations differ.");
        }

        private IExcelWorkbook OpenPeer(Stream stream) => _format switch {
            ComparisonWorkbookFormat.Xlsx => ExcelReaderApi.FromXlsx(stream),
            ComparisonWorkbookFormat.Xlsb => ExcelReaderApi.FromXlsb(stream),
            _ => throw new ArgumentOutOfRangeException(nameof(_format)),
        };

        /// <summary>Rewrites every ZIP part without compression, preserving its inflated payload.</summary>
        internal static byte[] Store(byte[] source) {
            using MemoryStream sourceStream = new MemoryStream(source, writable: false);
            using ZipArchive sourcePackage = new ZipArchive(sourceStream, ZipArchiveMode.Read);
            using MemoryStream output = new MemoryStream();
            using (ZipArchive target = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                foreach (ZipArchiveEntry entry in sourcePackage.Entries) {
                    using Stream input = entry.Open();
                    ZipArchiveEntry storedEntry = target.CreateEntry(entry.FullName, CompressionLevel.NoCompression);
                    storedEntry.LastWriteTime = entry.LastWriteTime;
                    using Stream destination = storedEntry.Open();
                    input.CopyTo(destination);
                }
            }
            return output.ToArray();
        }

        /// <summary>Checks both package CRCs and byte-identical inflated parts in a stored variant.</summary>
        internal static void ValidateInflatedParts(byte[] deflated, byte[] stored) {
            using MemoryStream originalStream = new MemoryStream(deflated, writable: false);
            using MemoryStream storedStream = new MemoryStream(stored, writable: false);
            using ZipArchive originalPackage = new ZipArchive(originalStream, ZipArchiveMode.Read);
            using ZipArchive storedPackage = new ZipArchive(storedStream, ZipArchiveMode.Read);
            WrittenWorkbookValidation.ValidateEntries(originalPackage);
            WrittenWorkbookValidation.ValidateEntries(storedPackage);
            if (originalPackage.Entries.Count != storedPackage.Entries.Count)
                throw new InvalidDataException("The stored ZIP variant changed the part count.");
            foreach (ZipArchiveEntry entry in originalPackage.Entries) {
                ZipArchiveEntry counterpart = storedPackage.GetEntry(entry.FullName)
                    ?? throw new InvalidDataException("The stored ZIP variant omitted a part.");
                if (entry.Length != counterpart.Length || counterpart.CompressedLength != counterpart.Length)
                    throw new InvalidDataException("A stored ZIP part has unexpected size or compression.");
                using Stream originalInput = entry.Open();
                using Stream storedInput = counterpart.Open();
                if (!SHA256.HashData(originalInput).AsSpan().SequenceEqual(SHA256.HashData(storedInput)))
                    throw new InvalidDataException("The stored ZIP variant changed an inflated part.");
            }
        }
    }

    /// <summary>Measures opening through the first row, including public reader initialization and disposal.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("SharedStringsFirstRow")]
    public class SharedStringFirstRowBenchmarks {
        private readonly SharedStringLifecycleWorkload _workload = new();
        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 65_536;
        [ParamsAllValues]
        public SharedStringZipStorage Storage { get; set; }
        public IEnumerable<int> RowCounts() => new StringHeavyReadBenchmarks().RowCounts();
        [GlobalSetup]
        public Task SetupAsync() => _workload.SetupAsync(RowCount, Storage);
        [Benchmark(Baseline = true)]
        public bool ExcelReaderOpenThroughFirstRow() => _workload.ExcelReaderFirstRow();
        [Benchmark]
        public bool OfficeIMOOpenThroughFirstRow() => _workload.OfficeIMOFirstRow();
        [Benchmark]
        public bool SylvanOpenThroughFirstRow() => _workload.SylvanFirstRow();
    }

    /// <summary>Opens a fresh reader and materializes only the header strings from a large shared table.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("SharedStringsMaterializedHeaders")]
    public class SharedStringHeaderBenchmarks {
        private readonly SharedStringLifecycleWorkload _workload = new();
        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 65_536;
        [ParamsAllValues]
        public SharedStringZipStorage Storage { get; set; }
        public IEnumerable<int> RowCounts() => new StringHeavyReadBenchmarks().RowCounts();
        [GlobalSetup]
        public Task SetupAsync() => _workload.SetupAsync(RowCount, Storage);
        [Benchmark(Baseline = true)]
        public long ExcelReaderMaterializedHeaders() => _workload.ExcelReaderMaterializedHeaders();
        [Benchmark]
        public long OfficeIMOMaterializedHeaders() => _workload.OfficeIMOMaterializedHeaders();
        [Benchmark]
        public long SylvanMaterializedHeaders() => _workload.SylvanMaterializedHeaders();
    }

    /// <summary>Consumes all materialized fields with compression varied independently of worksheet content.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("SharedStringsFullMaterialized")]
    public class SharedStringFullScanBenchmarks {
        private readonly SharedStringLifecycleWorkload _workload = new();
        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 65_536;
        [ParamsAllValues]
        public SharedStringZipStorage Storage { get; set; }
        public IEnumerable<int> RowCounts() => new StringHeavyReadBenchmarks().RowCounts();
        [GlobalSetup]
        public Task SetupAsync() => _workload.SetupAsync(RowCount, Storage);
        [Benchmark(Baseline = true)]
        public long ExcelReaderStringsMaterialized() => _workload.ExcelReaderFullMaterialized();
        [Benchmark]
        public long OfficeIMOStringsMaterialized() => _workload.OfficeIMOFullMaterialized();
        [Benchmark]
        public long SylvanStringsMaterialized() => _workload.SylvanFullMaterialized();
    }

    /// <summary>Retains the original borrowed-text end-to-end operations with deflated and stored ZIPs.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("SharedStringsFullBorrowed")]
    public class SharedStringBorrowedScanBenchmarks {
        private readonly SharedStringLifecycleWorkload _workload = new();
        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 65_536;
        [ParamsAllValues]
        public SharedStringZipStorage Storage { get; set; }
        public IEnumerable<int> RowCounts() => new StringHeavyReadBenchmarks().RowCounts();
        [GlobalSetup]
        public Task SetupAsync() => _workload.SetupAsync(RowCount, Storage);
        [Benchmark(Baseline = true)]
        public long ExcelReaderBorrowed() => _workload.ExcelReaderFullOriginal();
        [Benchmark]
        public long ExcelReaderBorrowedPrefetch() => _workload.ExcelReaderFullOriginal(prefetch: true);
    }

#if OFFICEIMO_BENCHMARK_NEW_APIS
    /// <summary>A public borrowed UTF8 scan with fresh reader caches for every invocation.</summary>
    [MemoryDiagnoser]
    [BenchmarkCategory("SharedStringsPublicUtf8")]
    public class SharedStringUtf8ReadBenchmarks {
        private readonly SharedStringLifecycleWorkload _workload = new();
        [ParamsSource(nameof(RowCounts))]
        public int RowCount { get; set; } = 100_000;
        [Params(ComparisonWorkbookFormat.Xlsx, ComparisonWorkbookFormat.Xlsb)]
        public ComparisonWorkbookFormat Format { get; set; } = ComparisonWorkbookFormat.Xlsx;
        [ParamsAllValues]
        public SharedStringZipStorage Storage { get; set; }
        public IEnumerable<int> RowCounts() {
            string? configured = Environment.GetEnvironmentVariable("OFFICEIMO_SHARED_UTF8_BENCHMARK_ROWS");
            int rows = configured is null ? 100_000 : int.Parse(configured, System.Globalization.CultureInfo.InvariantCulture);
            if (rows < 1 || rows > 1_048_575) throw new ArgumentOutOfRangeException(nameof(configured));
            yield return rows;
        }
        [GlobalSetup]
        public Task SetupAsync() => _workload.SetupBorrowedAsync(RowCount, Storage, Format);
        [Benchmark(Baseline = true)]
        public long ExcelReaderBorrowed() => _workload.ExcelReaderFullOriginal();
        [Benchmark]
        public long ExcelReaderBorrowedPrefetch() => _workload.ExcelReaderFullOriginal(prefetch: true);
        [Benchmark]
        public long OfficeIMOBorrowedSharedUtf8() => _workload.OfficeIMOFullBorrowed();
    }
#endif
}
