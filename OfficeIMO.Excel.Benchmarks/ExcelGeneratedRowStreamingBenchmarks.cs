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
    private static readonly ExcelTabularWriteOptions CompactWriteOptions = new() {
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
                ExcelDataSetImportResult result = ExcelDocument.WriteRows(
                    officeStream,
                    GenerateRows(RowCount),
                    Headers,
                    static (writer, row) => writer
                        .Write(row.Id)
                        .Write(row.Amount)
                        .Write(row.CreatedOn)
                        .Write(row.Active),
                    CompactWriteOptions);
                if (result.RowCount != RowCount) {
                    throw new InvalidDataException(
                        $"OfficeIMO reported {result.RowCount} rows instead of {RowCount} during write preflight.");
                }
            }
            ValidateWorkbook(nameof(WriteRowsGenerated), officePath, RowCount);

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
            ValidateWorkbook(nameof(ExcelReaderNetWriteRowsGenerated), excelReaderPath, RowCount);

            using (var officeAsyncStream = new FileStream(
                officeAsyncPath,
                FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None,
                256 * 1024,
                FileOptions.Asynchronous | FileOptions.SequentialScan)) {
                ExcelDataSetImportResult result = ExcelDocument.WriteRowsAsync(
                    officeAsyncStream,
                    GenerateRowsAsync(RowCount),
                    Headers,
                    static (writer, row) => writer
                        .Write(row.Id)
                        .Write(row.Amount)
                        .Write(row.CreatedOn)
                        .Write(row.Active),
                    CompactWriteOptions).GetAwaiter().GetResult();
                if (result.RowCount != RowCount) {
                    throw new InvalidDataException(
                        $"OfficeIMO async reported {result.RowCount} rows instead of {RowCount} during write preflight.");
                }
            }
            ValidateWorkbook(nameof(WriteRowsAsyncGenerated), officeAsyncPath, RowCount);

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
            ValidateWorkbook(nameof(ExcelReaderNetWriteRowsAsyncGenerated), excelReaderAsyncPath, RowCount);
        } finally {
            File.Delete(officePath);
            File.Delete(excelReaderPath);
            File.Delete(officeAsyncPath);
            File.Delete(excelReaderAsyncPath);
        }
    }

    [Params(1_000_000)]
    public int RowCount { get; set; }

    [Benchmark]
    public int WriteRowsGenerated() {
        ExcelDataSetImportResult result = ExcelDocument.WriteRows(
            Stream.Null,
            GenerateRows(RowCount),
            Headers,
            static (writer, row) => writer
                .Write(row.Id)
                .Write(row.Amount)
                .Write(row.CreatedOn)
                .Write(row.Active),
            CompactWriteOptions);
        return result.RowCount;
    }

    [Benchmark]
    public int ExcelReaderNetWriteRowsGenerated() =>
        WriteExcelReaderRows(Stream.Null, GenerateRows(RowCount));

    [Benchmark]
    public async Task<int> WriteRowsAsyncGenerated() {
        ExcelDataSetImportResult result = await ExcelDocument.WriteRowsAsync(
            Stream.Null,
            GenerateRowsAsync(RowCount),
            Headers,
            static (writer, row) => writer
                .Write(row.Id)
                .Write(row.Amount)
                .Write(row.CreatedOn)
                .Write(row.Active),
            CompactWriteOptions);
        return result.RowCount;
    }

    [Benchmark]
    public Task<int> ExcelReaderNetWriteRowsAsyncGenerated() =>
        WriteExcelReaderRowsAsync(Stream.Null, GenerateRowsAsync(RowCount));

    private static int WriteExcelReaderRows(Stream stream, IEnumerable<GeneratedRow> rows) {
        using XlsxWorkbookWriter workbook = XlsxWorkbookWriter.Create(stream, leaveOpen: true);
        workbook.Start();
        int count = 0;
        using (XlsxSheetWriter sheet = workbook.AddSheet("Data")) {
            sheet.Start();
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
        await using XlsxWorkbookWriter workbook = await XlsxWorkbookWriter.CreateAsync(
            stream,
            leaveOpen: true);
        await workbook.StartAsync();
        int count = 0;
        await using (XlsxSheetWriter sheet = workbook.AddSheet("Data")) {
            await sheet.StartAsync();
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
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(
            workbookPath,
            new ExcelReadOptions { NumericAsDecimal = true });
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
    }

    private static IEnumerable<GeneratedRow> GenerateRows(int count) {
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

    private static GeneratedRow CreateRow(int index) => new(
        index + 1,
        index * 1.25m,
        new DateTime(2026, 1, 1).AddMinutes(index),
        (index & 1) == 0);

    private static string CreatePreflightPath(string implementation) => Path.Combine(
        Path.GetTempPath(),
        $"OfficeIMO.Excel.Benchmarks.{implementation}.{Guid.NewGuid():N}.xlsx");

    private readonly record struct GeneratedRow(
        int Id,
        decimal Amount,
        DateTime CreatedOn,
        bool Active);
}
