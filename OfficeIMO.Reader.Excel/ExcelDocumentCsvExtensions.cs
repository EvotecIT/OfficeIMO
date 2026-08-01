using System.Data;
using System.Data.Common;
using System.Threading;
using OfficeIMO.CSV;
using OfficeIMO.Excel;

namespace OfficeIMO.Reader.Excel;

/// <summary>Imports CSV data through the shared OfficeIMO.CSV reader pipeline.</summary>
public static class ExcelDocumentCsvExtensions {
    /// <summary>Imports a loaded CSV document into a new worksheet.</summary>
    public static ExcelCsvImportResult ImportCsv(
        this ExcelDocument document,
        CsvDocument csv,
        ExcelCsvImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (csv == null) throw new ArgumentNullException(nameof(csv));
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
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        ExcelCsvImportOptions resolved = ResolveOptions(options);
        using var linkedCancellation = CreateLinkedLoadOptions(resolved, cancellationToken, out CsvLoadOptions loadOptions);
        using DbDataReader reader = CsvDocument.OpenDataReader(path, loadOptions, resolved.ReaderOptions);
        return ImportIntoNewWorksheet(document, reader, resolved, cancellationToken);
    }

    /// <summary>Reads a caller-owned CSV stream and imports it into a new worksheet.</summary>
    public static ExcelCsvImportResult ImportCsv(
        this ExcelDocument document,
        Stream stream,
        ExcelCsvImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        ExcelCsvImportOptions resolved = ResolveOptions(options);
        using var linkedCancellation = CreateLinkedLoadOptions(resolved, cancellationToken, out CsvLoadOptions loadOptions);
        using DbDataReader reader = CsvDocument.OpenDataReader(stream, loadOptions, resolved.ReaderOptions);
        return ImportIntoNewWorksheet(document, reader, resolved, cancellationToken);
    }

    /// <summary>
    /// Parses CSV text and imports it into a new worksheet. Because the input is already decoded,
    /// <see cref="CsvLoadOptions.Encoding"/> applies only to file and stream imports.
    /// </summary>
    public static ExcelCsvImportResult ImportCsvText(
        this ExcelDocument document,
        string text,
        ExcelCsvImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (text == null) throw new ArgumentNullException(nameof(text));
        ExcelCsvImportOptions resolved = ResolveOptions(options);
        using var linkedCancellation = CreateLinkedLoadOptions(resolved, cancellationToken, out CsvLoadOptions loadOptions);
        CsvDocument csv = CsvDocument.Parse(text, loadOptions);
        using DbDataReader reader = csv.CreateDataReader(resolved.ReaderOptions);
        return ImportIntoNewWorksheet(document, reader, resolved, cancellationToken);
    }

    /// <summary>Creates a detached Excel workbook from a loaded CSV document.</summary>
    public static ExcelDocument ToExcelDocument(
        this CsvDocument csv,
        ExcelCsvImportOptions? options = null,
        ExcelCreateOptions? createOptions = null,
        CancellationToken cancellationToken = default) {
        if (csv == null) throw new ArgumentNullException(nameof(csv));
        ExcelDocument document = ExcelDocument.Create(createOptions);
        try {
            document.ImportCsv(csv, options, cancellationToken);
            return document;
        } catch {
            document.Dispose();
            throw;
        }
    }

    /// <summary>Converts a loaded CSV document and saves it as an Excel workbook.</summary>
    public static void SaveAsExcel(
        this CsvDocument csv,
        string path,
        ExcelCsvImportOptions? importOptions = null,
        ExcelSaveOptions? saveOptions = null,
        CancellationToken cancellationToken = default) {
        if (csv == null) throw new ArgumentNullException(nameof(csv));
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        using ExcelDocument document = csv.ToExcelDocument(importOptions, cancellationToken: cancellationToken);
        document.SaveAsync(path, saveOptions, cancellationToken).GetAwaiter().GetResult();
    }

    /// <summary>Converts a loaded CSV document and saves it to an Excel workbook stream.</summary>
    public static void SaveAsExcel(
        this CsvDocument csv,
        Stream stream,
        ExcelCsvImportOptions? importOptions = null,
        ExcelSaveOptions? saveOptions = null,
        CancellationToken cancellationToken = default) {
        if (csv == null) throw new ArgumentNullException(nameof(csv));
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        using ExcelDocument document = csv.ToExcelDocument(importOptions, cancellationToken: cancellationToken);
        document.SaveAsync(stream, saveOptions, cancellationToken).GetAwaiter().GetResult();
    }

    private static ExcelCsvImportResult ImportIntoNewWorksheet(
        ExcelDocument document,
        IDataReader reader,
        ExcelCsvImportOptions options,
        CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(options.SheetName)) {
            throw new ArgumentException("SheetName cannot be empty.", nameof(options));
        }

        ExcelSheet sheet = document.AddWorksheet(options.SheetName.Trim());
        return ExcelSheetCsvExtensions.ImportCsvCore(sheet, reader, options, cancellationToken);
    }

    internal static ExcelCsvImportOptions ResolveOptions(ExcelCsvImportOptions? options) {
        ExcelCsvImportOptions resolved = options ?? new ExcelCsvImportOptions();
        if (resolved.LoadOptions == null) throw new ArgumentException("LoadOptions cannot be null.", nameof(options));
        if (resolved.ReaderOptions == null) throw new ArgumentException("ReaderOptions cannot be null.", nameof(options));
        if (resolved.StartRow < 1) throw new ArgumentOutOfRangeException(nameof(options), "StartRow must be at least 1.");
        if (resolved.StartColumn < 1) throw new ArgumentOutOfRangeException(nameof(options), "StartColumn must be at least 1.");
        return resolved;
    }

    internal static CancellationTokenSource CreateLinkedLoadOptions(
        ExcelCsvImportOptions options,
        CancellationToken cancellationToken,
        out CsvLoadOptions loadOptions) {
        var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken,
            options.LoadOptions.CancellationToken);
        loadOptions = options.LoadOptions.Clone();
        loadOptions.CancellationToken = linkedCancellation.Token;
        return linkedCancellation;
    }
}
