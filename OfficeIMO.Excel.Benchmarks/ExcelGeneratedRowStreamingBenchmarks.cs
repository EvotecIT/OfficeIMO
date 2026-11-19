using BenchmarkDotNet.Attributes;
using ExcelReader.Core.Writer;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Guards the single-pass typed-row contract with a generated source large enough to make
/// accidental row buffering visible in both execution time and managed allocations.
/// </summary>
[MemoryDiagnoser]
public class ExcelGeneratedRowStreamingBenchmarks {
    private static readonly string[] Headers = ["Id", "Amount", "CreatedOn", "Active"];
    private static readonly ExcelWriterOptions CompactWriteOptions = new() {
        IncludeCellReferences = false,
        UseSharedStrings = false
    };

    [GlobalSetup]
    public void Setup() {
        string officePath = CreatePreflightPath("officeimo");
        string excelReaderPath = CreatePreflightPath("excelreader");
        string officeAsyncPath = CreatePreflightPath("officeimo-async");
        string excelReaderAsyncPath = CreatePreflightPath("excelreader-async");
        try {
            using (var officeStream = new FileStream(
                officePath,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None,
                256 * 1024,
                FileOptions.SequentialScan)) {
                ExcelDataSetImportResult result;

                var destination1 = officeStream;

                var writeOptions1 = CompactWriteOptions ?? new ExcelWriterOptions {
                    UseSharedStrings = false,
                    CreateTable = false
                };

                using (var workbookWriter1 = ExcelDocument.CreateWriter(destination1, writeOptions1, leaveOpen: true)) {
                    result = workbookWriter1.WriteTable(writeOptions1.SheetName, GenerateRows(RowCount), Headers, static (writer, row) => writer.Write(row.Id).Write(row.Amount).Write(row.CreatedOn).Write(row.Active));
                }

                if (destination1.CanSeek)
                    destination1.Position = 0;
                if (result.RowCount != RowCount) {
                    throw new InvalidDataException(
                        $"OfficeIMO reported {result.RowCount} rows instead of {RowCount} during write preflight.");
                }
            }
            Record("OfficeIMO", nameof(WriteRowsGenerated), () => ValidateWorkbook(nameof(WriteRowsGenerated), officePath, RowCount));

            using (var excelReaderStream = new FileStream(
                excelReaderPath,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None,
                256 * 1024,
                FileOptions.SequentialScan)) {
                int excelReaderRows = WriteExcelReaderRows(
                    excelReaderStream,
                    GenerateRows(RowCount));
                if (excelReaderRows != RowCount) {
                    throw new InvalidDataException(
                        $"ExcelReader.NET reported {excelReaderRows} rows instead of {RowCount} during write preflight.");
                }
            }
            Record("ExcelReader.NET", nameof(WriteRowsGenerated), () => ValidateWorkbook(nameof(ExcelReaderNetWriteRowsGenerated), excelReaderPath, RowCount));

            using (var officeAsyncStream = new FileStream(
                officeAsyncPath,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None,
                256 * 1024,
                FileOptions.Asynchronous | FileOptions.SequentialScan)) {
                ExcelDataSetImportResult result;

                var destination2 = officeAsyncStream;

                var writeOptions2 = CompactWriteOptions ?? new ExcelWriterOptions {
                    UseSharedStrings = false,
                    CreateTable = false
                };

                using (var workbookWriter2 = ExcelDocument.CreateWriter(destination2, writeOptions2, leaveOpen: true)) {
                    result = workbookWriter2.WriteTable(writeOptions2.SheetName, GenerateRowsAsync(RowCount), Headers, static (writer, row) => writer.Write(row.Id).Write(row.Amount).Write(row.CreatedOn).Write(row.Active)).GetAwaiter().GetResult();
                }

                if (destination2.CanSeek)
                    destination2.Position = 0;
                if (result.RowCount != RowCount) {
                    throw new InvalidDataException(
                        $"OfficeIMO async reported {result.RowCount} rows instead of {RowCount} during write preflight.");
                }
            }
            Record("OfficeIMO", nameof(WriteRowsAsyncGenerated), () => ValidateWorkbook(nameof(WriteRowsAsyncGenerated), officeAsyncPath, RowCount));

            using (var excelReaderAsyncStream = new FileStream(
                excelReaderAsyncPath,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None,
                256 * 1024,
                FileOptions.Asynchronous | FileOptions.SequentialScan)) {
                int excelReaderRows = WriteExcelReaderRowsAsync(
                    excelReaderAsyncStream,
                    GenerateRowsAsync(RowCount)).GetAwaiter().GetResult();
                if (excelReaderRows != RowCount) {
                    throw new InvalidDataException(
                        $"ExcelReader.NET async reported {excelReaderRows} rows instead of {RowCount} during write preflight.");
                }
            }
            Record("ExcelReader.NET", nameof(WriteRowsAsyncGenerated), () => ValidateWorkbook(nameof(ExcelReaderNetWriteRowsAsyncGenerated), excelReaderAsyncPath, RowCount));
        }
        finally {
            File.Delete(officePath);
            File.Delete(excelReaderPath);
            File.Delete(officeAsyncPath);
            File.Delete(excelReaderAsyncPath);
        }
    }

    [Params(1_000_000)]
    public int RowCount { get; set; }

    // The preflight writes the same rows the timed lanes write to Stream.Null, to a file that two readers recheck.
    private void Record(string library, string caseName, Action validate) =>
        ExcelOutputValidation.Check(library, caseName,
            $"RowCount={RowCount}; OfficeIMO and ExcelDataReader reread every row's id, amount, date and flag", validate);

    [Benchmark]
    [BenchmarkCategory("case:WriteRowsGenerated")]
    public int WriteRowsGenerated() {
        ExcelDataSetImportResult result;

        var destination3 = Stream.Null;

        var writeOptions3 = CompactWriteOptions ?? new ExcelWriterOptions {
            UseSharedStrings = false,
            CreateTable = false
        };

        using (var workbookWriter3 = ExcelDocument.CreateWriter(destination3, writeOptions3, leaveOpen: true)) {
            result = workbookWriter3.WriteTable(writeOptions3.SheetName, GenerateRows(RowCount), Headers, static (writer, row) => writer.Write(row.Id).Write(row.Amount).Write(row.CreatedOn).Write(row.Active));
        }

        if (destination3.CanSeek)
            destination3.Position = 0;
        return result.RowCount;
    }

    [Benchmark]
    [BenchmarkCategory("case:WriteRowsGenerated")]
    public int ExcelReaderNetWriteRowsGenerated() =>
        WriteExcelReaderRows(Stream.Null, GenerateRows(RowCount));

    [Benchmark]
    [BenchmarkCategory("case:WriteRowsAsyncGenerated")]
    public async Task<int> WriteRowsAsyncGenerated() {
        ExcelDataSetImportResult result;

        var destination4 = Stream.Null;

        var writeOptions4 = CompactWriteOptions ?? new ExcelWriterOptions {
            UseSharedStrings = false,
            CreateTable = false
        };

        using (var workbookWriter4 = ExcelDocument.CreateWriter(destination4, writeOptions4, leaveOpen: true)) {
            result = await workbookWriter4.WriteTable(writeOptions4.SheetName, GenerateRowsAsync(RowCount), Headers, static (writer, row) => writer.Write(row.Id).Write(row.Amount).Write(row.CreatedOn).Write(row.Active));
        }

        if (destination4.CanSeek)
            destination4.Position = 0;
        return result.RowCount;
    }

    [Benchmark]
    [BenchmarkCategory("case:WriteRowsAsyncGenerated")]
    public Task<int> ExcelReaderNetWriteRowsAsyncGenerated() =>
        WriteExcelReaderRowsAsync(Stream.Null, GenerateRowsAsync(RowCount));

    private static int WriteExcelReaderRows(Stream stream, IEnumerable<GeneratedRow> rows) {
        using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(stream, leaveOpen: true);
#if !NET10_0_OR_GREATER
        workbook.Start();
#endif
        int count = 0;
        using (XlsxSheetWriter sheet = workbook.AddSheet("Data")) {
#if !NET10_0_OR_GREATER
            sheet.Start();
#endif
            using (XlsxRowWriter header = sheet.StartRow()) {
                WriteHeaders(header);
            }

            foreach (GeneratedRow value in rows) {
                using XlsxRowWriter row = sheet.StartRow();
                WriteRow(row, value);
                count++;
            }
            sheet.End();
        }
        workbook.End();
        return count;
    }

    private static async Task<int> WriteExcelReaderRowsAsync(
        Stream stream,
        IAsyncEnumerable<GeneratedRow> rows) {
#if NET10_0_OR_GREATER
        await using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(stream, leaveOpen: true);
#else
        await using XlsxWorkbookWriter workbook = await XlsxWorkbookWriter.CreateAsync(
            stream,
            leaveOpen: true);
        await workbook.StartAsync();
#endif
        int count = 0;
        await using (XlsxSheetWriter sheet = workbook.AddSheet("Data")) {
#if !NET10_0_OR_GREATER
            await sheet.StartAsync();
#endif
            await using (XlsxRowWriter header = await sheet.StartRowAsync()) {
                WriteHeaders(header);
            }

            await foreach (GeneratedRow value in rows) {
                await using XlsxRowWriter row = await sheet.StartRowAsync();
                WriteRow(row, value);
                count++;
            }
            await sheet.EndAsync();
        }
        await workbook.EndAsync();
        return count;
    }

    private static void WriteHeaders<TRow>(TRow row)
        where TRow : IRowWriter {
        foreach (string header in Headers) {
            row.Write(header);
        }
    }

    private static void WriteRow<TRow>(TRow row, GeneratedRow value)
        where TRow : IRowWriter {
        row.Write(value.Id);
        row.Write(value.Amount);
        row.Write(value.CreatedOn);
        row.Write(value.Active);
    }

    private static void ValidateWorkbook(
        string method,
        string workbookPath,
        int expectedRowCount) {
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenReader(
            workbookPath,
            new ExcelReaderOptions { NumericAsDecimal = true });
        if (reader.FieldCount != Headers.Length) {
            throw new InvalidDataException(
                $"{method} exposed {reader.FieldCount} columns instead of {Headers.Length}.");
        }

        for (int column = 0; column < Headers.Length; column++) {
            if (!string.Equals(reader.GetName(column), Headers[column], StringComparison.Ordinal)) {
                throw new InvalidDataException($"{method} header {column + 1} did not round-trip.");
            }
        }

        int index = 0;
        while (reader.Read()) {
            if (index >= expectedRowCount) {
                throw new InvalidDataException($"{method} emitted extra rows.");
            }

            GeneratedRow expected = CreateRow(index);
            if (reader.GetInt32(0) != expected.Id
                || reader.GetDecimal(1) != expected.Amount
                || reader.GetDateTime(2) != expected.CreatedOn
                || reader.GetBoolean(3) != expected.Active) {
                throw new InvalidDataException($"{method} row {index + 2} did not round-trip.");
            }
            index++;
        }

        if (index != expectedRowCount || reader.NextResult()) {
            throw new InvalidDataException($"{method} did not round-trip the exact single-sheet row set.");
        }

        ValidateIndependently(method, workbookPath, expectedRowCount);
    }

    // ExcelDataReader returns numbers as doubles, styled dates as DateTime and booleans as bool; a writer that stores
    // the flag as a number or the date without a style is still accepted when the stored value is the same.
    private static void ValidateIndependently(string method, string workbookPath, int expectedRowCount) {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        using FileStream stream = File.OpenRead(workbookPath);
        using var reader = global::ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
        if (!reader.Read() || reader.FieldCount != Headers.Length) {
            throw new InvalidDataException($"{method} has no header row for ExcelDataReader.");
        }
        for (int column = 0; column < Headers.Length; column++) {
            if (!Equals(reader.GetValue(column), Headers[column])) {
                throw new InvalidDataException($"{method} header {column + 1} reads as '{reader.GetValue(column)}' in ExcelDataReader.");
            }
        }
        int index = 0;
        while (reader.Read()) {
            if (index >= expectedRowCount) throw new InvalidDataException($"{method} emitted extra rows for ExcelDataReader.");
            GeneratedRow expected = CreateRow(index);
            object id = reader.GetValue(0), amount = reader.GetValue(1), date = reader.GetValue(2), active = reader.GetValue(3);
            bool dateMatches = date switch {
                DateTime value => value == expected.CreatedOn,
                double serial => serial == expected.CreatedOn.ToOADate(),
                _ => false
            };
            bool activeMatches = active switch {
                bool value => value == expected.Active,
                double number => number == (expected.Active ? 1 : 0),
                _ => false
            };
            if (id is not double idValue || idValue != expected.Id
                || amount is not double amountValue || amountValue != (double)expected.Amount
                || !dateMatches || !activeMatches) {
                throw new InvalidDataException($"{method} row {index + 2} reads as {id}/{amount}/{date}/{active} in ExcelDataReader.");
            }
            index++;
        }
        if (index != expectedRowCount) {
            throw new InvalidDataException($"{method} has {index} data rows in ExcelDataReader, expected {expectedRowCount}.");
        }
    }

    internal static IEnumerable<GeneratedRow> GenerateRows(int count) {
        for (int index = 0; index < count; index++) {
            yield return CreateRow(index);
        }
    }

    private static async IAsyncEnumerable<GeneratedRow> GenerateRowsAsync(int count) {
        await Task.CompletedTask;
        for (int index = 0; index < count; index++) {
            yield return CreateRow(index);
        }
    }

    internal static GeneratedRow CreateRow(int index) => new(
        index + 1,
        index * 1.25m,
        new DateTime(2026, 1, 1).AddMinutes(index),
        (index & 1) == 0);

    private static string CreatePreflightPath(string implementation) => Path.Combine(
        Path.GetTempPath(),
        $"OfficeIMO.Excel.Benchmarks.{implementation}.{Guid.NewGuid():N}.xlsx");

    internal readonly record struct GeneratedRow(
        int Id,
        decimal Amount,
        DateTime CreatedOn,
        bool Active);
}
