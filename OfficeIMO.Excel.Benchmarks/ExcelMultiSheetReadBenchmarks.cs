#if NET10_0_OR_GREATER
using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Runtime.CompilerServices;
using System.Text;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Configs;
using ClosedXML.Excel;
using ExcelReader.Core.Reader;
using GeneratedRecord = OfficeIMO.Excel.Benchmarks.ExcelReaderGeneratedBenchmarks.GeneratedRecord;
using ExcelReaderApi = ExcelReader.Core.Reader.Excel;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Reads every row of a workbook with several equally sized four-column sheets, one sheet after
/// another and, where a library can open sheets independently, all sheets at once. Sheet
/// <c>s</c> holds generated records <c>s * RowsPerSheet</c> onwards, and every lane checks every
/// sheet's row count and checksum.
/// </summary>
[MemoryDiagnoser]
[CategoriesColumn]
[GroupBenchmarksBy(BenchmarkLogicalGroupRule.ByCategory)]
public class ExcelMultiSheetReadBenchmarks {
    private string _path = string.Empty;
    private long[] _expected = [];

    [Params(ExcelFileFormat.Xlsx, ExcelFileFormat.Xlsb, ExcelFileFormat.Xls)]
    public ExcelFileFormat Format { get; set; }

    [Params(8)]
    public int SheetCount { get; set; }

    [Params(25_000)]
    public int RowsPerSheet { get; set; }

    [GlobalSetup]
    public void Setup() {
        OfficeIMO.Benchmarks.BenchmarkProcessPriority.ApplyConfiguredPriority();
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        string root = Environment.GetEnvironmentVariable("OFFICEIMO_BENCHMARK_DATA") ?? Path.GetTempPath();
        Directory.CreateDirectory(root);
        _path = Path.Combine(root, $"multi-sheet-{SheetCount}x{RowsPerSheet}-{Guid.NewGuid():N}.{Format.ToString().ToLowerInvariant()}");
        File.WriteAllBytes(_path, Write(Format, SheetCount, RowsPerSheet));
        _expected = new long[SheetCount];
        for (int sheet = 0; sheet < SheetCount; sheet++) {
            for (int row = 0; row < RowsPerSheet; row++) _expected[sheet] = Add(_expected[sheet], Record(sheet, row));
        }
        try {
            OfficeIMO_Sequential();
            OfficeIMO_Parallel();
            ExcelReaderNet_Sequential();
            ExcelReaderNet_Parallel();
            Sylvan_Sequential();
            ExcelDataReader_Sequential();
        } catch { Cleanup(); throw; }
    }

    [GlobalCleanup]
    public void Cleanup() { if (_path.Length > 0) File.Delete(_path); }

    internal static byte[] Write(ExcelFileFormat format, int sheetCount, int rowsPerSheet) {
        using var output = new MemoryStream();
        using (var writer = ExcelDocument.CreateWriter(format, output,
            new ExcelWriterOptions { IncludeCellReferences = false, UseSharedStrings = false }, leaveOpen: true)) {
            for (int sheet = 0; sheet < sheetCount; sheet++) {
                writer.WriteTable(SheetName(sheet), Generate(sheet, rowsPerSheet), ExcelReaderGeneratedBenchmarks.Headers,
                    static (row, value) => row.Write(value.Name).Write(value.Id).Write(value.CreatedOn).Write(value.Amount));
            }
        }
        return output.ToArray();
    }

    internal static string SheetName(int sheet) => "Data" + (sheet + 1).ToString(CultureInfo.InvariantCulture);

    private static IEnumerable<GeneratedRecord> Generate(int sheet, int rows) {
        for (int row = 0; row < rows; row++) yield return Record(sheet, row, rows);
    }

    private GeneratedRecord Record(int sheet, int row) => Record(sheet, row, RowsPerSheet);

    private static GeneratedRecord Record(int sheet, int row, int rowsPerSheet) =>
        ExcelReaderGeneratedBenchmarks.CreateRecord(sheet * rowsPerSheet + row);

    // Streaming readers ---------------------------------------------------------------------------

    [Benchmark(Baseline = true), BenchmarkCategory("Stream")]
    public long OfficeIMO_Sequential() {
        using var reader = ExcelDocument.OpenReader(_path);
        int sheet = 0;
        long total = 0;
        do {
            total += Check(sheet, ReadSheet(reader, sheet));
            sheet++;
        } while (reader.NextResult());
        return CheckSheets(sheet, total);
    }

    [Benchmark, BenchmarkCategory("Stream")]
    public long OfficeIMO_Parallel() {
        long[] results = new long[SheetCount];
        Parallel.For(0, SheetCount, sheet => {
            using var reader = ExcelDocument.OpenReader(_path, new ExcelReaderOptions { SheetIndex = sheet });
            results[sheet] = Check(sheet, ReadSheet(reader, sheet));
        });
        return CheckSheets(SheetCount, results.Sum());
    }

    [Benchmark, BenchmarkCategory("Stream")]
    public long ExcelReaderNet_Sequential() {
        long total = 0;
        for (int sheet = 0; sheet < SheetCount; sheet++) total += Check(sheet, ReadExcelReaderSheet(_path, sheet));
        return CheckSheets(SheetCount, total);
    }

    [Benchmark, BenchmarkCategory("Stream")]
    public long ExcelReaderNet_Parallel() {
        long[] results = new long[SheetCount];
        Parallel.For(0, SheetCount, sheet => results[sheet] = Check(sheet, ReadExcelReaderSheet(_path, sheet)));
        return CheckSheets(SheetCount, results.Sum());
    }

    [Benchmark, BenchmarkCategory("Stream")]
    public long Sylvan_Sequential() {
        using var reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(_path);
        int sheet = 0;
        long total = 0;
        do {
            total += Check(sheet, ReadSheet(reader, sheet));
            sheet++;
        } while (reader.NextResult());
        return CheckSheets(sheet, total);
    }

    [Benchmark, BenchmarkCategory("Stream")]
    public long ExcelDataReader_Sequential() {
        using var stream = File.OpenRead(_path);
        using var reader = global::ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
        int sheet = 0;
        long total = 0;
        do {
            if (!reader.Read()) throw new InvalidDataException("Missing header.");
            total += Check(sheet, ReadSheet(reader, sheet));
            sheet++;
        } while (reader.NextResult());
        return CheckSheets(sheet, total);
    }

    // Reading -------------------------------------------------------------------------------------

    [MethodImpl(MethodImplOptions.NoInlining)]
    internal (long Checksum, int Rows) ReadSheet(IDataReader reader, int sheet) {
        bool typed = reader is DbDataReader;
        ExcelReaderGeneratedBenchmarks.ValidateHeaders(reader, typed);
        long checksum = 0;
        int count = 0;
        while (reader.Read()) {
            string name = reader.GetString(0);
            int id = typed ? reader.GetInt32(1) : checked((int)reader.GetDouble(1));
            checksum = Add(checksum, name, id, reader.GetDateTime(2), reader.GetDouble(3));
            count++;
        }
        return (checksum, count);
    }

    private (long Checksum, int Rows) ReadExcelReaderSheet(string path, int sheet) {
        switch (Format) {
            case ExcelFileFormat.Xlsx:
                using (var workbook = ExcelReaderApi.FromXlsxFile(path)) {
                    var state = new RowState(workbook.IsDate1904);
                    foreach (Row row in workbook.Sheets[sheet]) state.Add(row);
                    return state.Result;
                }
            case ExcelFileFormat.Xlsb:
                using (var workbook = ExcelReaderApi.FromXlsbFile(path)) {
                    var state = new RowState(workbook.IsDate1904);
                    foreach (Row row in workbook.Sheets[sheet]) state.Add(row);
                    return state.Result;
                }
            case ExcelFileFormat.Xls:
                using (var workbook = ExcelReaderApi.FromXlsFile(path)) {
                    var state = new RowState(workbook.IsDate1904);
                    foreach (Row row in workbook.Sheets[sheet]) state.Add(row);
                    return state.Result;
                }
            default: throw new ArgumentOutOfRangeException();
        }
    }

    private struct RowState(bool date1904) {
        private long _checksum;
        private int _count = -1;

        internal readonly (long Checksum, int Rows) Result => (_checksum, _count);

        internal void Add(Row row) {
            if (_count++ < 0) {
                for (int i = 0; i < ExcelReaderGeneratedBenchmarks.Headers.Length; i++)
                    if (row[i].GetString() != ExcelReaderGeneratedBenchmarks.Headers[i]) throw new InvalidDataException("Incorrect header.");
                return;
            }
            if (!row[1].TryParse(CultureInfo.InvariantCulture, out int id)
                || !row[2].TryGetDateTime(date1904, out DateTime date)
                || !row[3].TryParse(CultureInfo.InvariantCulture, out double amount)) throw new InvalidDataException("Incorrect typed value.");
            _checksum = ExcelMultiSheetReadBenchmarks.Add(_checksum, row[0].GetString(), id, date, amount);
        }
    }
    internal long Check(int sheet, (long Checksum, int Rows) result) {
        if (sheet >= SheetCount || result.Rows != RowsPerSheet || result.Checksum != _expected[sheet])
            throw new InvalidDataException($"{Format} sheet {sheet}: unexpected rows or checksum ({result.Rows}, {result.Checksum}).");
        return result.Checksum;
    }

    internal long CheckSheets(int sheets, long total) {
        if (sheets != SheetCount) throw new InvalidDataException($"{Format}: read {sheets} of {SheetCount} sheets.");
        return total;
    }

    private static long Add(long checksum, GeneratedRecord row) => Add(checksum, row.Name, row.Id, row.CreatedOn, row.Amount);

    private static long Add(long checksum, string name, int id, DateTime date, double amount) {
        unchecked {
            foreach (char c in name) checksum = checksum * 31 + c;
            return ((checksum * 31 + id) * 31 + date.Ticks) * 31 + BitConverter.DoubleToInt64Bits(amount);
        }
    }

    // Editable workbooks (XLSX) -------------------------------------------------------------------

    /// <summary>Loads the XLSX workbook into an editable model and reads every sheet.</summary>
    [MemoryDiagnoser]
    public class Workbook {
        private readonly ExcelMultiSheetReadBenchmarks _owner = new() { Format = ExcelFileFormat.Xlsx };

        [Params(8)]
        public int SheetCount { get; set; }

        [Params(25_000)]
        public int RowsPerSheet { get; set; }

        [GlobalSetup]
        public void Setup() {
            _owner.SheetCount = SheetCount;
            _owner.RowsPerSheet = RowsPerSheet;
            _owner.Setup();
            OfficeIMO_Sequential();
            ClosedXml_Sequential();
        }

        [GlobalCleanup]
        public void Cleanup() => _owner.Cleanup();

        [Benchmark(Baseline = true)]
        public long OfficeIMO_Sequential() {
            using var document = ExcelDocument.Load(_owner._path);
            long total = 0;
            for (int sheet = 0; sheet < SheetCount; sheet++) {
                using var reader = document.Sheets[sheet].CreateDataReader();
                total += _owner.Check(sheet, _owner.ReadSheet(reader, sheet));
            }
            return _owner.CheckSheets(document.Sheets.Count, total);
        }

        [Benchmark]
        public long ClosedXml_Sequential() {
            using var workbook = new XLWorkbook(_owner._path);
            long total = 0;
            int sheet = 0;
            foreach (IXLWorksheet worksheet in workbook.Worksheets) {
                long checksum = 0;
                int count = -1;
                foreach (IXLRow row in worksheet.RowsUsed()) {
                    if (count++ < 0) continue;
                    checksum = Add(checksum, row.Cell(1).GetString(), row.Cell(2).GetValue<int>(),
                        row.Cell(3).GetDateTime(), row.Cell(4).GetDouble());
                }
                total += _owner.Check(sheet++, (checksum, count));
            }
            return _owner.CheckSheets(sheet, total);
        }
    }
}
#endif
