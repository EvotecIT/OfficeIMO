using System.Data.Common;
using OfficeIMO.CSV;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Csv;

namespace OfficeIMO.Tool.Commands.Tabular;

internal static class TabularCommand {
    private static readonly HashSet<string> ExcelExtensions = new(StringComparer.OrdinalIgnoreCase) {
        ".xlsx", ".xlsm", ".xltx", ".xltm", ".xlam", ".xlsb", ".xls"
    };
    private static readonly HashSet<string> WritableExcelExtensions = new(StringComparer.OrdinalIgnoreCase) {
        ".xlsx", ".xlsb", ".xls"
    };
    private static readonly HashSet<string> CsvExtensions = new(StringComparer.OrdinalIgnoreCase) {
        ".csv", ".tsv"
    };

    internal const string Usage = """
OfficeIMO.Tool - tabular data

Usage:
  officeimo tabular sheets <workbook>
  officeimo tabular schema <workbook|csv> [--sheet <name>|--sheet-index <index>]
                           [--delimiter <character|\t>] [--no-header]
  officeimo tabular convert <input> <output> [--sheet <name>|--sheet-index <index>]
                            [--delimiter <character|\t>] [--no-header] [--force]

Workbook input formats: XLSX, XLSM, XLTX, XLTM, XLAM, XLSB, and XLS.
Workbook output formats: XLSX, XLSB, and XLS.
CSV and TSV conversion uses the OfficeIMO.CSV streaming reader and writer.
""";

    internal static async Task<int> RunAsync(
        string[] args,
        TextWriter standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken = default) {
        TabularArguments parsed;
        try {
            parsed = TabularArguments.Parse(args);
        } catch (TabularUsageException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        }

        if (parsed.Command == TabularCommandKind.Help) {
            await standardOutput.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        }
        if (!File.Exists(parsed.InputPath)) {
            await standardError.WriteLineAsync("Input was not found: " + parsed.InputPath).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.InputNotFound;
        }

        try {
            cancellationToken.ThrowIfCancellationRequested();
            ValidatePathOptions(parsed);
            return parsed.Command switch {
                TabularCommandKind.Sheets => await ListSheetsAsync(parsed, standardOutput, cancellationToken).ConfigureAwait(false),
                TabularCommandKind.Schema => await WriteSchemaAsync(parsed, standardOutput, cancellationToken).ConfigureAwait(false),
                TabularCommandKind.Convert => await ConvertAsync(parsed, standardOutput, cancellationToken).ConfigureAwait(false),
                _ => (int)OfficeImoToolExitCode.Usage
            };
        } catch (OperationCanceledException) {
            return (int)OfficeImoToolExitCode.Cancelled;
        } catch (NotSupportedException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.UnsupportedInput;
        } catch (TabularOutputException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OutputFailed;
        } catch (Exception exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OperationFailed;
        }
    }

    private static async Task<int> ListSheetsAsync(
        TabularArguments options,
        TextWriter output,
        CancellationToken cancellationToken) {
        EnsureExcelPath(options.InputPath, "sheets");
        using ExcelWorkbookDataReader reader = ExcelDocument.OpenDataReader(
            options.InputPath,
            new ExcelReadOptions { CancellationToken = cancellationToken });
        foreach (string sheetName in reader.SheetNames) {
            cancellationToken.ThrowIfCancellationRequested();
            await output.WriteLineAsync(sheetName).ConfigureAwait(false);
        }
        return (int)OfficeImoToolExitCode.Success;
    }

    private static async Task<int> WriteSchemaAsync(
        TabularArguments options,
        TextWriter output,
        CancellationToken cancellationToken) {
        using DbDataReader reader = OpenReader(options, inferSchema: true, cancellationToken);
        await output.WriteLineAsync("ordinal\tname\ttype").ConfigureAwait(false);
        for (int ordinal = 0; ordinal < reader.FieldCount; ordinal++) {
            cancellationToken.ThrowIfCancellationRequested();
            await output.WriteLineAsync(
                ordinal + "\t" + reader.GetName(ordinal) + "\t" + reader.GetFieldType(ordinal).FullName)
                .ConfigureAwait(false);
        }
        return (int)OfficeImoToolExitCode.Success;
    }

    private static async Task<int> ConvertAsync(
        TabularArguments options,
        TextWriter output,
        CancellationToken cancellationToken) {
        string outputPath = options.OutputPath!;
        string outputDirectory = Path.GetDirectoryName(outputPath) ?? string.Empty;
        if (outputDirectory.Length != 0 && !Directory.Exists(outputDirectory)) {
            throw new TabularOutputException("Output directory does not exist: " + outputDirectory);
        }
        if (File.Exists(outputPath) && !options.Force) {
            throw new TabularOutputException("Output already exists. Use --force to replace it: " + outputPath);
        }

        string outputExtension = Path.GetExtension(outputPath);
        if (!WritableExcelExtensions.Contains(outputExtension) && !CsvExtensions.Contains(outputExtension)) {
            throw new NotSupportedException("Output must use XLSX, XLSB, XLS, CSV, or TSV.");
        }
        string temporaryPath = Path.Combine(
            outputDirectory,
            Path.GetFileNameWithoutExtension(outputPath) + "." + Guid.NewGuid().ToString("N") + outputExtension);
        try {
            await WriteConversionAsync(options, temporaryPath, cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            File.Move(temporaryPath, outputPath, overwrite: options.Force);
        } catch {
            if (File.Exists(temporaryPath)) File.Delete(temporaryPath);
            throw;
        }

        await output.WriteLineAsync("Converted to " + outputPath).ConfigureAwait(false);
        return (int)OfficeImoToolExitCode.Success;
    }

    private static async Task WriteConversionAsync(
        TabularArguments options,
        string temporaryPath,
        CancellationToken cancellationToken) {
        bool inputIsCsv = IsCsvPath(options.InputPath);
        bool outputIsCsv = IsCsvPath(temporaryPath);
        if (inputIsCsv || outputIsCsv) {
            using DbDataReader reader = OpenReader(
                options,
                inferSchema: inputIsCsv && !outputIsCsv,
                cancellationToken);
            if (outputIsCsv) {
                var saveOptions = new CsvSaveOptions {
                    Delimiter = ResolveDelimiter(temporaryPath, options.Delimiter),
                    IncludeHeader = options.HasHeaderRow,
                    NoClobber = true
                };
                CsvDocument.WriteDataReader(temporaryPath, reader, saveOptions, cancellationToken);
                return;
            }

            string extension = Path.GetExtension(temporaryPath);
            if (string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)) {
                await using var stream = new FileStream(
                    temporaryPath,
                    FileMode.CreateNew,
                    FileAccess.Write,
                    FileShare.None,
                    128 * 1024,
                    FileOptions.Asynchronous | FileOptions.SequentialScan);
                ExcelDocument.WriteDataReader(
                    stream,
                    reader,
                    new ExcelTabularWriteOptions {
                        IncludeHeaders = options.HasHeaderRow,
                        CreateTable = false,
                        UseSharedStrings = false
                    },
                    cancellationToken);
                await stream.FlushAsync(cancellationToken).ConfigureAwait(false);
                return;
            }

            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.InsertDataReader(
                reader,
                includeHeaders: options.HasHeaderRow,
                createTable: false,
                ct: cancellationToken);
            await document.SaveAsync(temporaryPath, cancellationToken: cancellationToken).ConfigureAwait(false);
            return;
        }

        EnsureExcelPath(options.InputPath, "convert");
        if (string.Equals(
                Path.GetExtension(options.InputPath),
                Path.GetExtension(temporaryPath),
                StringComparison.OrdinalIgnoreCase)) {
            await CopyFileAsync(options.InputPath, temporaryPath, cancellationToken).ConfigureAwait(false);
            return;
        }
        await ExcelDocument.ConvertAsync(
            options.InputPath,
            temporaryPath,
            cancellationToken: cancellationToken).ConfigureAwait(false);
    }

    internal static async Task CopyFileAsync(
        string inputPath,
        string outputPath,
        CancellationToken cancellationToken) {
        await using var input = new FileStream(
            inputPath,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            128 * 1024,
            FileOptions.Asynchronous | FileOptions.SequentialScan);
        await using var output = new FileStream(
            outputPath,
            FileMode.CreateNew,
            FileAccess.Write,
            FileShare.None,
            128 * 1024,
            FileOptions.Asynchronous | FileOptions.SequentialScan);
        await input.CopyToAsync(output, 128 * 1024, cancellationToken).ConfigureAwait(false);
        await output.FlushAsync(cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
    }

    private static DbDataReader OpenReader(
        TabularArguments options,
        bool inferSchema,
        CancellationToken cancellationToken) {
        if (IsCsvPath(options.InputPath)) {
            var loadOptions = new CsvLoadOptions {
                HasHeaderRow = options.HasHeaderRow,
                DetectDelimiter = !options.Delimiter.HasValue,
                Delimiter = ResolveDelimiter(options.InputPath, options.Delimiter),
                CancellationToken = cancellationToken
            };
            return CsvDocument.OpenDataReader(
                options.InputPath,
                loadOptions,
                new CsvDataReaderOptions { InferSchema = inferSchema });
        }

        EnsureExcelPath(options.InputPath, "read");
        return ExcelDocument.OpenDataReader(
            options.InputPath,
            new ExcelReadOptions {
                SheetName = options.SheetName,
                SheetIndex = options.SheetIndex,
                HasHeaderRow = options.HasHeaderRow,
                InferSchema = inferSchema,
                CancellationToken = cancellationToken
            });
    }

    private static char ResolveDelimiter(string path, char? delimiter) => delimiter ??
        (string.Equals(Path.GetExtension(path), ".tsv", StringComparison.OrdinalIgnoreCase) ? '\t' : ',');

    private static bool IsCsvPath(string path) => CsvExtensions.Contains(Path.GetExtension(path));

    private static void ValidatePathOptions(TabularArguments options) {
        bool inputIsCsv = IsCsvPath(options.InputPath);
        if (inputIsCsv && (options.SheetName is not null || options.SheetIndex.HasValue)) {
            throw new NotSupportedException("--sheet and --sheet-index apply only to workbook input.");
        }

        if (options.Command == TabularCommandKind.Convert &&
            !inputIsCsv &&
            options.OutputPath is not null &&
            !IsCsvPath(options.OutputPath) &&
            (options.SheetName is not null || options.SheetIndex.HasValue)) {
            throw new NotSupportedException("Sheet selection is available when converting a workbook to CSV or TSV, not workbook-to-workbook conversion.");
        }
    }

    private static void EnsureExcelPath(string path, string operation) {
        if (!ExcelExtensions.Contains(Path.GetExtension(path))) {
            throw new NotSupportedException(operation + " requires a supported Excel workbook.");
        }
    }
}

internal sealed class TabularOutputException : Exception {
    internal TabularOutputException(string message) : base(message) { }
}
