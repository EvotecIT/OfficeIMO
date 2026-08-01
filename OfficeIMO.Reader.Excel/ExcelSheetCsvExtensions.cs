using System.Data;
using System.Data.Common;
using System.Globalization;
using System.Text;
using System.Threading;
using OfficeIMO.CSV;
using OfficeIMO.Excel;

namespace OfficeIMO.Reader.Excel;

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
        using DbDataReader reader = CsvDocument.OpenDataReader(path, resolved.LoadOptions, resolved.ReaderOptions);
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
        using DbDataReader reader = CsvDocument.OpenDataReader(stream, resolved.LoadOptions, resolved.ReaderOptions);
        return ImportCsvCore(sheet, reader, resolved, cancellationToken);
    }

    /// <summary>Parses CSV text and imports it into an existing worksheet.</summary>
    public static ExcelCsvImportResult ImportCsvText(
        this ExcelSheet sheet,
        string text,
        ExcelCsvImportOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        if (text == null) throw new ArgumentNullException(nameof(text));
        ExcelCsvImportOptions resolved = ExcelDocumentCsvExtensions.ResolveOptions(options);
        Encoding encoding = resolved.LoadOptions.Encoding ?? new UTF8Encoding(false);
        using var stream = new MemoryStream(encoding.GetBytes(text), writable: false);
        using DbDataReader reader = CsvDocument.OpenDataReader(stream, resolved.LoadOptions, resolved.ReaderOptions);
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
        return CsvDocument.Parse(
            sheet.ToCsv(headersInFirstRow, readOptions: readOptions, cancellationToken: cancellationToken),
            new CsvLoadOptions { HasHeaderRow = headersInFirstRow });
    }

    /// <summary>Converts an A1 worksheet range to a materialized CSV document.</summary>
    public static CsvDocument ToCsvDocument(
        this ExcelSheet sheet,
        string a1Range,
        bool headersInFirstRow = true,
        ExcelReadOptions? readOptions = null,
        CancellationToken cancellationToken = default) {
        return CsvDocument.Parse(
            sheet.ToCsv(a1Range, headersInFirstRow, readOptions: readOptions, cancellationToken: cancellationToken),
            new CsvLoadOptions { HasHeaderRow = headersInFirstRow });
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
        using DataTable table = sheet.ToDataTable(headersInFirstRow, options: readOptions, ct: cancellationToken);
        using IDataReader reader = table.CreateDataReader();
        CsvDocument.WriteDataReader(path, reader, ResolveSaveOptions(csvOptions, headersInFirstRow));
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
        using DataTable table = sheet.ToDataTable(headersInFirstRow, options: readOptions, ct: cancellationToken);
        using IDataReader reader = table.CreateDataReader();
        CsvDocument.WriteDataReader(
            stream,
            reader,
            ResolveSaveOptions(csvOptions, headersInFirstRow),
            leaveOpen: true);
    }

    internal static ExcelCsvImportResult ImportCsvCore(
        ExcelSheet sheet,
        IDataReader reader,
        ExcelCsvImportOptions options,
        CancellationToken cancellationToken) {
        string range = sheet.InsertDataReader(
            reader,
            options.StartRow,
            options.StartColumn,
            options.IncludeHeaders,
            options.TableName,
            options.TableStyle,
            options.IncludeAutoFilter,
            options.CreateTable,
            options.AutoFit,
            cancellationToken);

        return new ExcelCsvImportResult(sheet.Name, range);
    }

    private static string ToCsvCore(
        ExcelSheet sheet,
        string? a1Range,
        bool headersInFirstRow,
        CsvSaveOptions? csvOptions,
        ExcelReadOptions? readOptions,
        CancellationToken cancellationToken) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        using DataTable table = a1Range == null
            ? sheet.ToDataTable(headersInFirstRow, options: readOptions, ct: cancellationToken)
            : sheet.ToDataTable(a1Range, headersInFirstRow, options: readOptions, ct: cancellationToken);
        using IDataReader reader = table.CreateDataReader();
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        CsvDocument.WriteDataReader(writer, reader, ResolveSaveOptions(csvOptions, headersInFirstRow));
        return writer.ToString();
    }

    private static CsvSaveOptions ResolveSaveOptions(CsvSaveOptions? options, bool headersInFirstRow) {
        CsvSaveOptions resolved = options?.Clone() ?? new CsvSaveOptions();
        resolved.IncludeHeader = headersInFirstRow && resolved.IncludeHeader;
        return resolved;
    }
}
