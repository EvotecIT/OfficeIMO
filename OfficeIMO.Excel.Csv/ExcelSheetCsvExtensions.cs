using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.CSV;
using OfficeIMO.Excel;

namespace OfficeIMO.Excel.Csv;

/// <summary>Imports and exports worksheet data through the shared OfficeIMO.CSV pipeline.</summary>
public static class ExcelSheetCsvExtensions {
    /// <summary>Imports a loaded CSV document into an existing worksheet.</summary>
    public static ExcelCsvImportResult ImportCsv(
        this ExcelSheet sheet,
        CsvDocument csv,
        ExcelCsvImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(csv);
        cancellationToken.ThrowIfCancellationRequested();
        ExcelCsvImportOptions resolved = ExcelDocumentCsvExtensions.ResolveOptions(options);
        using DbDataReader reader = csv.CreateDataReader(resolved.ReaderOptions);
        return ImportCsvCore(sheet, reader, resolved, cancellationToken);
    }

    /// <summary>Opens a CSV file and imports it into an existing worksheet.</summary>
    public static ExcelCsvImportResult ImportCsvFile(
        this ExcelSheet sheet,
        string path,
        ExcelCsvImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(sheet);
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        ExcelCsvImportOptions resolved = ExcelDocumentCsvExtensions.ResolveOptions(options);
        using var linkedCancellation = ExcelDocumentCsvExtensions.CreateLinkedLoadOptions(resolved, cancellationToken, out CsvReaderOptions loadOptions);
        using DbDataReader reader = CsvDocument.OpenReader(path, loadOptions);
        return ImportCsvCore(sheet, reader, resolved, cancellationToken);
    }

    /// <summary>Reads a caller-owned CSV stream and imports it into an existing worksheet.</summary>
    public static ExcelCsvImportResult ImportCsv(
        this ExcelSheet sheet,
        Stream stream,
        ExcelCsvImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(sheet);
        ExcelCsvImportOptions resolved = ExcelDocumentCsvExtensions.ResolveOptions(options);
        using var linkedCancellation = ExcelDocumentCsvExtensions.CreateLinkedLoadOptions(resolved, cancellationToken, out CsvReaderOptions loadOptions);
        using DbDataReader reader = CsvDocument.OpenReader(stream, loadOptions);
        return ImportCsvCore(sheet, reader, resolved, cancellationToken);
    }

    /// <summary>
    /// Parses CSV text and imports it into an existing worksheet. Because the input is already decoded,
    /// <see cref="CsvReaderOptions.Encoding"/> applies only to file and stream imports.
    /// </summary>
    public static ExcelCsvImportResult ImportCsvText(
        this ExcelSheet sheet,
        string text,
        ExcelCsvImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(sheet);
        ArgumentNullException.ThrowIfNull(text);
        ExcelCsvImportOptions resolved = ExcelDocumentCsvExtensions.ResolveOptions(options);
        using var linkedCancellation = ExcelDocumentCsvExtensions.CreateLinkedLoadOptions(resolved, cancellationToken, out CsvReaderOptions loadOptions);
        using DbDataReader reader = CsvDocument.OpenReader(
            new StringReader(text), loadOptions);
        return ImportCsvCore(sheet, reader, resolved, cancellationToken);
    }

    /// <summary>Projects the selected worksheet values into a materialized CSV model.</summary>
    public static CsvDocument ToCsv(this ExcelSheet sheet, ExcelCsvExportOptions? options = null,
        CancellationToken cancellationToken = default) => sheet.ToCsv(options, out _, cancellationToken);

    /// <summary>Projects values and reports the limits of CSV fidelity for workbook features.</summary>
    public static CsvDocument ToCsv(this ExcelSheet sheet, ExcelCsvExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(sheet);
        ExcelCsvExportOptions resolved = ResolveExportOptions(options);
        cancellationToken.ThrowIfCancellationRequested();
        report = CreateExportReport(sheet);
        if (resolved.FailOnLoss) report.ThrowIfLoss();
        return ToCsvDocumentCore(sheet, resolved.Range, resolved.HeadersInFirstRow, resolved.ReadOptions, cancellationToken);
    }

    /// <summary>Projects worksheet values asynchronously and returns the model and fidelity report.</summary>
    public static Task<OfficeConversionResult<CsvDocument>> ToCsvAsync(this ExcelSheet sheet,
        ExcelCsvExportOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeConversionOutput.RunAsync(() => (sheet.ToCsv(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);

    /// <summary>Streams worksheet values as CSV without materializing the target document.</summary>
    public static OfficeConversionReport SaveAsCsv(this ExcelSheet sheet,
        string path, ExcelCsvExportOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(sheet);
        ExcelCsvExportOptions resolved = ResolveExportOptions(options);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeConversionReport report = CreateExportReport(sheet);
        if (resolved.FailOnLoss) report.ThrowIfLoss();
        using ExcelWorkbookDataReader reader = CreateSheetReader(sheet, resolved.Range, resolved.HeadersInFirstRow, resolved.ReadOptions, cancellationToken);
        OfficeConversionFailures.Run(report, () => CsvDocument.WriteDataReader(path, reader,
            ResolveSaveOptions(resolved.SaveOptions, resolved.HeadersInFirstRow), cancellationToken));
        return OfficeConversionOutput.WithPath(report, Path.GetFullPath(path));
    }

    /// <summary>Streams worksheet values as CSV without materializing the target document.</summary>
    public static OfficeConversionReport SaveAsCsv(this ExcelSheet sheet,
        Stream stream, ExcelCsvExportOptions? options = null, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(sheet);
        ExcelCsvExportOptions resolved = ResolveExportOptions(options);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeConversionReport report = CreateExportReport(sheet);
        if (resolved.FailOnLoss) report.ThrowIfLoss();
        using ExcelWorkbookDataReader reader = CreateSheetReader(sheet, resolved.Range, resolved.HeadersInFirstRow, resolved.ReadOptions, cancellationToken);
        OfficeConversionFailures.Run(report, () => CsvDocument.WriteDataReader(stream, reader,
            ResolveSaveOptions(resolved.SaveOptions, resolved.HeadersInFirstRow), leaveOpen: true, cancellationToken: cancellationToken));
        return report;
    }

    /// <summary>Asynchronously streams worksheet values as CSV without materializing the target document.</summary>
    public static Task<OfficeConversionReport> SaveAsCsvAsync(this ExcelSheet sheet,
        string path, ExcelCsvExportOptions? options = null, CancellationToken cancellationToken = default) {
        return Task.Run(() => sheet.SaveAsCsv(path, options, cancellationToken), cancellationToken);
    }

    /// <summary>Asynchronously streams worksheet values as CSV without materializing the target document.</summary>
    public static Task<OfficeConversionReport> SaveAsCsvAsync(this ExcelSheet sheet,
        Stream stream, ExcelCsvExportOptions? options = null, CancellationToken cancellationToken = default) {
        return Task.Run(() => sheet.SaveAsCsv(stream, options, cancellationToken), cancellationToken);
    }

    private static ExcelCsvExportOptions ResolveExportOptions(ExcelCsvExportOptions? options) {
        ExcelCsvExportOptions resolved = options ?? new ExcelCsvExportOptions();
        if (resolved.Range != null && string.IsNullOrWhiteSpace(resolved.Range)) throw new ArgumentException("Range cannot be empty.", nameof(options));
        return resolved;
    }

    private static OfficeConversionReport CreateExportReport(ExcelSheet sheet) => new OfficeConversionReport("Excel", "CSV",
        diagnostics: new[] { new OfficeDiagnostic("EXCEL_CSV_VALUES_ONLY", OfficeDiagnosticSeverity.Warning,
            "CSV projects selected cell values. Workbook formatting, formulas, drawings and other native features are not assessed for preservation.",
            sheet.Name, "Excel.Csv", OfficeConversionLossKind.Unassessed) });

    internal static ExcelCsvImportResult ImportCsvCore(
        ExcelSheet sheet,
        IDataReader reader,
        ExcelCsvImportOptions options,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        char delimiter = reader is ICsvDataReaderMetadata metadata
            ? metadata.Delimiter
            : options.ReaderOptions.Delimiter;
        string delimiterText = reader is ICsvDataReaderDialectMetadata dialect
            ? dialect.DelimiterText
            : options.ReaderOptions.DelimiterText ?? delimiter.ToString();
        if (reader.FieldCount == 0) {
            return new ExcelCsvImportResult(sheet.Name, null, string.Empty, delimiterText);
        }

        string tableName = string.IsNullOrWhiteSpace(options.TableName)
            ? sheet.Name
            : options.TableName!.Trim();
        ExcelDataReaderInsertResult imported = sheet.InsertDataReaderWithResult(
            reader,
            options.StartRow,
            options.StartColumn,
            options.IncludeHeaders,
            tableName,
            options.TableStyle,
            options.IncludeAutoFilter,
            options.CreateTable,
            options.AutoFit,
            cancellationToken);

        return new ExcelCsvImportResult(imported.SheetName, imported.TableName, imported.Range, delimiterText);
    }

    private static CsvDocument ToCsvDocumentCore(
        ExcelSheet sheet,
        string? a1Range,
        bool headersInFirstRow,
        ExcelReaderOptions? readOptions,
        CancellationToken cancellationToken) {
        ArgumentNullException.ThrowIfNull(sheet);
        using ExcelWorkbookDataReader reader = CreateSheetReader(
            sheet, a1Range, headersInFirstRow, readOptions, cancellationToken);
        var document = CsvDocument.Create();
        if (reader.FieldCount == 0) return document;

        var headers = new string[reader.FieldCount];
        for (int index = 0; index < headers.Length; index++) headers[index] = reader.GetName(index);
        document.WithHeader(headers);

        var values = new object?[reader.FieldCount];
        while (reader.Read()) {
            cancellationToken.ThrowIfCancellationRequested();
            for (int index = 0; index < values.Length; index++) {
                object value = reader.GetValue(index);
                values[index] = ReferenceEquals(value, DBNull.Value) ? null : value;
            }
            document.AddRow((object?[])values.Clone());
        }
        return document;
    }

    private static ExcelWorkbookDataReader CreateSheetReader(
        ExcelSheet sheet,
        string? a1Range,
        bool headersInFirstRow,
        ExcelReaderOptions? readOptions,
        CancellationToken cancellationToken) {
        ExcelReaderOptions source = readOptions ?? new ExcelReaderOptions();
        var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken,
            source.CancellationToken);
        try {
            ExcelReaderOptions effective = source.ForSheet(
                sheet.Name,
                a1Range,
                headersInFirstRow,
                linkedCancellation.Token);
            return sheet.CreateDataReader(effective).OwnLifetime(linkedCancellation);
        } catch {
            linkedCancellation.Dispose();
            throw;
        }
    }

    private static CsvSaveOptions ResolveSaveOptions(CsvSaveOptions? options, bool headersInFirstRow) {
        CsvSaveOptions resolved = options?.Clone() ?? new CsvSaveOptions();
        resolved.IncludeHeader = headersInFirstRow && resolved.IncludeHeader;
        return resolved;
    }
}
