using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Threading;
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
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (csv == null) throw new ArgumentNullException(nameof(csv));
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
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("File path cannot be empty.", nameof(path));
        ExcelCsvImportOptions resolved = ExcelDocumentCsvExtensions.ResolveOptions(options);
        using var linkedCancellation = ExcelDocumentCsvExtensions.CreateLinkedLoadOptions(resolved, cancellationToken, out CsvLoadOptions loadOptions);
        using DbDataReader reader = CsvDocument.OpenDataReader(path, loadOptions, resolved.ReaderOptions);
        return ImportCsvCore(sheet, reader, resolved, cancellationToken);
    }

    /// <summary>Reads a caller-owned CSV stream and imports it into an existing worksheet.</summary>
    public static ExcelCsvImportResult ImportCsv(
        this ExcelSheet sheet,
        Stream stream,
        ExcelCsvImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        ExcelCsvImportOptions resolved = ExcelDocumentCsvExtensions.ResolveOptions(options);
        using var linkedCancellation = ExcelDocumentCsvExtensions.CreateLinkedLoadOptions(resolved, cancellationToken, out CsvLoadOptions loadOptions);
        using DbDataReader reader = CsvDocument.OpenDataReader(stream, loadOptions, resolved.ReaderOptions);
        return ImportCsvCore(sheet, reader, resolved, cancellationToken);
    }

    /// <summary>
    /// Parses CSV text and imports it into an existing worksheet. Because the input is already decoded,
    /// <see cref="CsvLoadOptions.Encoding"/> applies only to file and stream imports.
    /// </summary>
    public static ExcelCsvImportResult ImportCsvText(
        this ExcelSheet sheet,
        string text,
        ExcelCsvImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (text == null) throw new ArgumentNullException(nameof(text));
        ExcelCsvImportOptions resolved = ExcelDocumentCsvExtensions.ResolveOptions(options);
        using var linkedCancellation = ExcelDocumentCsvExtensions.CreateLinkedLoadOptions(resolved, cancellationToken, out CsvLoadOptions loadOptions);
        using DbDataReader reader = CsvDocument.OpenTextDataReader(
            text, loadOptions, resolved.ReaderOptions);
        return ImportCsvCore(sheet, reader, resolved, cancellationToken);
    }

    /// <summary>Converts the worksheet used range to CSV text.</summary>
    public static string ToCsv(
        this ExcelSheet sheet,
        bool headersInFirstRow = true,
        CsvSaveOptions? csvOptions = null,
        ExcelReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        return ToCsvCore(sheet, a1Range: null, headersInFirstRow, csvOptions, readOptions, cancellationToken);
    }

    /// <summary>Converts an A1 worksheet range to CSV text.</summary>
    public static string ToCsv(
        this ExcelSheet sheet,
        string a1Range,
        bool headersInFirstRow = true,
        CsvSaveOptions? csvOptions = null,
        ExcelReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentException("Range cannot be empty.", nameof(a1Range));
        return ToCsvCore(sheet, a1Range, headersInFirstRow, csvOptions, readOptions, cancellationToken);
    }

    /// <summary>Converts the worksheet used range to a materialized CSV document.</summary>
    public static CsvDocument ToCsvDocument(
        this ExcelSheet sheet,
        bool headersInFirstRow = true,
        ExcelReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        return ToCsvDocumentCore(sheet, a1Range: null, headersInFirstRow, readOptions, cancellationToken);
    }

    /// <summary>Converts an A1 worksheet range to a materialized CSV document.</summary>
    public static CsvDocument ToCsvDocument(
        this ExcelSheet sheet,
        string a1Range,
        bool headersInFirstRow = true,
        ExcelReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(a1Range)) throw new ArgumentException("Range cannot be empty.", nameof(a1Range));
        return ToCsvDocumentCore(sheet, a1Range, headersInFirstRow, readOptions, cancellationToken);
    }

    /// <summary>Saves the worksheet used range as CSV.</summary>
    public static void SaveAsCsv(
        this ExcelSheet sheet,
        string path,
        bool headersInFirstRow = true,
        CsvSaveOptions? csvOptions = null,
        ExcelReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        using ExcelWorkbookDataReader reader = CreateSheetReader(
            sheet, a1Range: null, headersInFirstRow, readOptions, cancellationToken);
        CsvDocument.WriteDataReader(
            path,
            reader,
            ResolveSaveOptions(csvOptions, headersInFirstRow),
            cancellationToken);
    }

    /// <summary>Saves the worksheet used range to a caller-owned CSV stream.</summary>
    public static void SaveAsCsv(
        this ExcelSheet sheet,
        Stream stream,
        bool headersInFirstRow = true,
        CsvSaveOptions? csvOptions = null,
        ExcelReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        using ExcelWorkbookDataReader reader = CreateSheetReader(
            sheet, a1Range: null, headersInFirstRow, readOptions, cancellationToken);
        CsvDocument.WriteDataReader(
            stream,
            reader,
            ResolveSaveOptions(csvOptions, headersInFirstRow),
            leaveOpen: true,
            cancellationToken: cancellationToken);
    }

    internal static ExcelCsvImportResult ImportCsvCore(
        ExcelSheet sheet,
        IDataReader reader,
        ExcelCsvImportOptions options,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        char delimiter = reader is ICsvDataReaderMetadata metadata
            ? metadata.Delimiter
            : options.LoadOptions.Delimiter;
        if (reader.FieldCount == 0) {
            return new ExcelCsvImportResult(sheet.Name, null, string.Empty, delimiter);
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

        return new ExcelCsvImportResult(imported.SheetName, imported.TableName, imported.Range, delimiter);
    }

    private static string ToCsvCore(
        ExcelSheet sheet,
        string? a1Range,
        bool headersInFirstRow,
        CsvSaveOptions? csvOptions,
        ExcelReadOptions? readOptions,
        CancellationToken cancellationToken) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        using ExcelWorkbookDataReader reader = CreateSheetReader(
            sheet, a1Range, headersInFirstRow, readOptions, cancellationToken);
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        CsvDocument.WriteDataReader(
            writer,
            reader,
            ResolveSaveOptions(csvOptions, headersInFirstRow),
            cancellationToken);
        return writer.ToString();
    }

    private static CsvDocument ToCsvDocumentCore(
        ExcelSheet sheet,
        string? a1Range,
        bool headersInFirstRow,
        ExcelReadOptions? readOptions,
        CancellationToken cancellationToken) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        using ExcelWorkbookDataReader reader = CreateSheetReader(
            sheet, a1Range, headersInFirstRow, readOptions, cancellationToken);
        var document = new CsvDocument();
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
        ExcelReadOptions? readOptions,
        CancellationToken cancellationToken) {
        ExcelReadOptions source = readOptions ?? new ExcelReadOptions();
        var linkedCancellation = CancellationTokenSource.CreateLinkedTokenSource(
            cancellationToken,
            source.CancellationToken);
        try {
            ExcelReadOptions effective = source.ForSheet(
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
