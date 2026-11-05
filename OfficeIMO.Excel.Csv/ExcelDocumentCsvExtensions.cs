using System.Data;
using System.Collections.Generic;
using System.Data.Common;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Core.Internal;
using OfficeIMO.CSV;
using OfficeIMO.Excel;

namespace OfficeIMO.Excel.Csv;

/// <summary>Imports CSV data through the shared OfficeIMO.CSV reader pipeline.</summary>
public static class ExcelDocumentCsvExtensions {
    /// <summary>Imports a loaded CSV document into a new worksheet.</summary>
    public static ExcelCsvImportResult ImportCsv(
        this ExcelDocument document,
        CsvDocument csv,
        ExcelCsvImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(csv);
        cancellationToken.ThrowIfCancellationRequested();
        ExcelCsvImportOptions resolved = ResolveOptions(options);
        using DbDataReader reader = csv.CreateDataReader(resolved.ReaderOptions);
        return ImportIntoNewWorksheet(document, reader, resolved, cancellationToken);
    }

    /// <summary>Opens a CSV file and imports it into a new worksheet.</summary>
    public static ExcelCsvImportResult ImportCsvFile(
        this ExcelDocument document,
        string path,
        ExcelCsvImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        ExcelCsvImportOptions resolved = ResolveOptions(options);
        using var linkedCancellation = CreateLinkedLoadOptions(resolved, cancellationToken, out CsvReaderOptions loadOptions);
        using DbDataReader reader = CsvDocument.OpenReader(path, loadOptions);
        return ImportIntoNewWorksheet(document, reader, resolved, cancellationToken);
    }

    /// <summary>Reads a caller-owned CSV stream and imports it into a new worksheet.</summary>
    public static ExcelCsvImportResult ImportCsv(
        this ExcelDocument document,
        Stream stream,
        ExcelCsvImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        ExcelCsvImportOptions resolved = ResolveOptions(options);
        using var linkedCancellation = CreateLinkedLoadOptions(resolved, cancellationToken, out CsvReaderOptions loadOptions);
        using DbDataReader reader = CsvDocument.OpenReader(stream, loadOptions);
        return ImportIntoNewWorksheet(document, reader, resolved, cancellationToken);
    }

    /// <summary>
    /// Parses CSV text and imports it into a new worksheet. Because the input is already decoded,
    /// <see cref="CsvReaderOptions.Encoding"/> applies only to file and stream imports.
    /// </summary>
    public static ExcelCsvImportResult ImportCsvText(
        this ExcelDocument document,
        string text,
        ExcelCsvImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        ArgumentNullException.ThrowIfNull(text);
        ExcelCsvImportOptions resolved = ResolveOptions(options);
        using var linkedCancellation = CreateLinkedLoadOptions(resolved, cancellationToken, out CsvReaderOptions loadOptions);
        using DbDataReader reader = CsvDocument.OpenReader(
            new StringReader(text), loadOptions);
        return ImportIntoNewWorksheet(document, reader, resolved, cancellationToken);
    }

    /// <summary>Converts CSV into a detached editable workbook owned by the caller.</summary>
    public static ExcelDocument ToExcel(this CsvDocument csv, CsvExcelExportOptions? options = null,
        CancellationToken cancellationToken = default) => csv.ToExcel(options, out _, cancellationToken);

    /// <summary>Converts CSV and returns the completed data-projection report.</summary>
    public static ExcelDocument ToExcel(this CsvDocument csv, CsvExcelExportOptions? options,
        out OfficeConversionReport report, CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(csv);
        cancellationToken.ThrowIfCancellationRequested();
        CsvExcelExportOptions resolved = options ?? new CsvExcelExportOptions();
        ArgumentNullException.ThrowIfNull(resolved.ImportOptions);
        ExcelDocument document = ExcelDocument.Create(resolved.CreateOptions);
        try {
            document.ImportCsv(csv, resolved.ImportOptions, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            var diagnostics = new List<OfficeDiagnostic>();
            if (!resolved.ImportOptions.IncludeHeaders) diagnostics.Add(new OfficeDiagnostic(
                "CSV_EXCEL_HEADERS_OMITTED", OfficeDiagnosticSeverity.Warning, "CSV field names are not written to the workbook.",
                format: "Excel.Csv", lossKind: OfficeConversionLossKind.Omission));
            if (resolved.ImportOptions.ReaderOptions.InferSchema || resolved.ImportOptions.ReaderOptions.Schema != null) diagnostics.Add(new OfficeDiagnostic(
                "CSV_EXCEL_SCHEMA_PROJECTION", OfficeDiagnosticSeverity.Warning,
                "CSV values use the requested typed schema projection; preservation of their original lexical spelling is not assessed.",
                format: "Excel.Csv", lossKind: OfficeConversionLossKind.Unassessed));
            report = new OfficeConversionReport("CSV", "Excel", diagnostics: diagnostics);
            if (resolved.FailOnLoss) report.ThrowIfLoss();
            return document;
        } catch { document.Dispose(); throw; }
    }

    /// <summary>Converts CSV asynchronously and returns the workbook with its report.</summary>
    public static Task<OfficeConversionResult<ExcelDocument>> ToExcelAsync(this CsvDocument csv,
        CsvExcelExportOptions? options = null, CancellationToken cancellationToken = default) =>
        OfficeConversionOutput.RunAsync(() => (csv.ToExcel(options, out OfficeConversionReport report, cancellationToken), report), cancellationToken);

    /// <summary>Writes CSV data as a workbook and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsExcel(this CsvDocument csv, string path,
        CsvExcelExportOptions? options = null, CancellationToken cancellationToken = default) {
        using ExcelDocument document = csv.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        OfficeConversionFailures.Run(report, () => document.SaveAsync(path, options?.SaveOptions, cancellationToken).GetAwaiter().GetResult());
        return OfficeConversionOutput.WithPath(report, Path.GetFullPath(path));
    }

    /// <summary>Writes CSV data as a workbook and returns the conversion report.</summary>
    public static OfficeConversionReport SaveAsExcel(this CsvDocument csv, Stream stream,
        CsvExcelExportOptions? options = null, CancellationToken cancellationToken = default) {
        using ExcelDocument document = csv.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        OfficeConversionFailures.Run(report, () => document.SaveAsync(stream, options?.SaveOptions, cancellationToken).GetAwaiter().GetResult());
        return report;
    }

    /// <summary>Asynchronously writes CSV data as a workbook and returns the conversion report.</summary>
    public static async Task<OfficeConversionReport> SaveAsExcelAsync(this CsvDocument csv, string path,
        CsvExcelExportOptions? options = null, CancellationToken cancellationToken = default) {
        using ExcelDocument document = csv.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        await OfficeConversionFailures.RunAsync(report, () => document.SaveAsync(path, options?.SaveOptions, cancellationToken)).ConfigureAwait(false);
        return OfficeConversionOutput.WithPath(report, Path.GetFullPath(path));
    }

    /// <summary>Asynchronously writes CSV data as a workbook and returns the conversion report.</summary>
    public static async Task<OfficeConversionReport> SaveAsExcelAsync(this CsvDocument csv, Stream stream,
        CsvExcelExportOptions? options = null, CancellationToken cancellationToken = default) {
        using ExcelDocument document = csv.ToExcel(options, out OfficeConversionReport report, cancellationToken);
        await OfficeConversionFailures.RunAsync(report, () => document.SaveAsync(stream, options?.SaveOptions, cancellationToken)).ConfigureAwait(false);
        return report;
    }

    private static ExcelCsvImportResult ImportIntoNewWorksheet(
        ExcelDocument document,
        IDataReader reader,
        ExcelCsvImportOptions options,
        CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(options.SheetName)) {
            throw new ArgumentException("SheetName cannot be empty.", nameof(options));
        }

        ExcelSheet sheet = document.AddSheet(options.SheetName.Trim());
        return ExcelSheetCsvExtensions.ImportCsvCore(sheet, reader, options, cancellationToken);
    }

    internal static ExcelCsvImportOptions ResolveOptions(ExcelCsvImportOptions? options) {
        ExcelCsvImportOptions resolved = options ?? new ExcelCsvImportOptions();
        if (resolved.ReaderOptions == null) throw new ArgumentException("ReaderOptions cannot be null.", nameof(options));
        if (resolved.StartRow < 1) throw new ArgumentOutOfRangeException(nameof(options), "StartRow must be at least 1.");
        if (resolved.StartColumn < 1) throw new ArgumentOutOfRangeException(nameof(options), "StartColumn must be at least 1.");
        return resolved;
    }

    internal static CancellationTokenSource CreateLinkedLoadOptions(
        ExcelCsvImportOptions options,
        CancellationToken cancellationToken,
        out CsvReaderOptions loadOptions) {
        var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken,
            options.ReaderOptions.CancellationToken);
        loadOptions = options.ReaderOptions.Clone();
        loadOptions.CancellationToken = linkedCancellation.Token;
        return linkedCancellation;
    }
}
