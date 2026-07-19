using System.Data;
using System.Globalization;
using System.Threading;
using OfficeIMO.CSV;
using OfficeIMO.Excel;

namespace OfficeIMO.Reader.Excel;

/// <summary>
/// Converts worksheet ranges and Excel tables to and from CSV text.
/// </summary>
public static class ExcelSheetCsvExtensions {
    /// <summary>Reads an A1 range and returns CSV text.</summary>
    public static string ToCsv(this ExcelSheet sheet, string a1Range, bool headersInFirstRow = true, ExcelReadOptions? options = null, ExecutionMode? mode = null, CancellationToken ct = default) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        using DataTable table = sheet.ToDataTable(a1Range, headersInFirstRow, options, mode, ct);
        return ToCsv(table, headersInFirstRow);
    }

    /// <summary>Reads the worksheet used range and returns CSV text.</summary>
    public static string ToCsv(this ExcelSheet sheet, bool headersInFirstRow = true, ExcelReadOptions? options = null, ExecutionMode? mode = null, CancellationToken ct = default) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        using DataTable table = sheet.ToDataTable(headersInFirstRow, options, mode, ct);
        return ToCsv(table, headersInFirstRow);
    }

    /// <summary>Reads an Excel table and returns CSV text.</summary>
    public static string TableToCsv(this ExcelSheet sheet, string tableName, bool? headersInFirstRow = null, ExcelReadOptions? options = null, ExecutionMode? mode = null, CancellationToken ct = default) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        bool includeHeaders = headersInFirstRow ?? true;
        using DataTable table = sheet.TableToDataTable(tableName, headersInFirstRow, options, mode, ct);
        return ToCsv(table, includeHeaders);
    }

    /// <summary>Inserts CSV text into the worksheet and returns the inserted range.</summary>
    public static string FromCsv(this ExcelSheet sheet, string csv, int startRow = 1, int startColumn = 1, bool firstRowIsHeader = true, bool includeHeaders = true, ExecutionMode? mode = null, CancellationToken ct = default) {
        if (sheet == null) throw new ArgumentNullException(nameof(sheet));
        using DataTable table = ExcelCsvDataTableBuilder.FromText(
            csv,
            delimiter: ',',
            headersInFirstRow: firstRowIsHeader,
            skipInitialRecords: 0,
            culture: CultureInfo.InvariantCulture,
            convertNumbersAndDates: false);
        sheet.InsertDataTable(table, startRow, startColumn, includeHeaders, mode, ct);
        return BuildInsertedRange(table, startRow, startColumn, includeHeaders);
    }

    private static string ToCsv(DataTable table, bool includeHeaders) {
        using var reader = table.CreateDataReader();
        using var writer = new StringWriter(CultureInfo.InvariantCulture);
        CsvDocument.WriteDataReader(writer, reader, new CsvSaveOptions {
            IncludeHeader = includeHeaders,
            Culture = CultureInfo.InvariantCulture
        });
        return writer.ToString();
    }

    private static string BuildInsertedRange(DataTable table, int startRow, int startColumn, bool includeHeaders) {
        int rowCount = table.Rows.Count + (includeHeaders ? 1 : 0);
        if (table.Columns.Count == 0 || rowCount == 0) return string.Empty;

        return A1.CellReference(startRow, startColumn) + ":" +
               A1.CellReference(startRow + rowCount - 1, startColumn + table.Columns.Count - 1);
    }
}
