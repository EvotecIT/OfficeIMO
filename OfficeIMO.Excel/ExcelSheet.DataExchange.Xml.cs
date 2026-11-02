using System.Threading;

namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    /// <summary>
    /// Inserts a positional or explicitly mapped XML table and returns the inserted A1 range.
    /// Text remains literal, including formula-looking strings. The XML is fully admitted and converted before worksheet mutation.
    /// </summary>
    /// <param name="xml">XML table text, separate from an XML Spreadsheet workbook.</param>
    /// <param name="startRow">First destination row, one-based.</param>
    /// <param name="startColumn">First destination column, one-based.</param>
    /// <param name="includeHeaders">Writes mapped column names, or positional Column1, Column2 names.</param>
    /// <param name="options">Explicit schema and resource limits.</param>
    /// <param name="mode">Existing worksheet insertion execution mode.</param>
    public string FromXml(string xml, int startRow = 1, int startColumn = 1, bool includeHeaders = false,
        ExcelXmlTableOptions? options = null, ExcelExecutionMode? mode = null) {
        if (xml == null) throw new ArgumentNullException(nameof(xml));
        var resolved = (options ?? new ExcelXmlTableOptions()).Snapshot();
        using var table = XmlTables.ExcelXmlTableCodec.Read(xml, resolved);
        return InsertXmlTable(table, startRow, startColumn, includeHeaders, resolved, mode);
    }

    /// <summary>Inserts a bounded XML table from a caller-owned stream; the stream remains open.</summary>
    /// <param name="input">Readable XML table stream.</param>
    /// <param name="startRow">First destination row, one-based.</param>
    /// <param name="startColumn">First destination column, one-based.</param>
    /// <param name="includeHeaders">Writes mapped or positional column headers.</param>
    /// <param name="options">Explicit schema and resource limits.</param>
    /// <param name="mode">Existing worksheet insertion execution mode.</param>
    public string FromXml(Stream input, int startRow = 1, int startColumn = 1, bool includeHeaders = false,
        ExcelXmlTableOptions? options = null, ExcelExecutionMode? mode = null) =>
        InsertXmlTable(input, startRow, startColumn, includeHeaders, (options ?? new ExcelXmlTableOptions()).Snapshot(), mode);

    private string InsertXmlTable(Stream input, int row, int column, bool headers, ExcelXmlTableOptions options, ExcelExecutionMode? mode) {
        using var table = XmlTables.ExcelXmlTableCodec.Read(input, options);
        return InsertXmlTable(table, row, column, headers, options, mode);
    }

    private string InsertXmlTable(System.Data.DataTable table, int row, int column, bool headers, ExcelXmlTableOptions options, ExcelExecutionMode? mode) {
        if (row < 1 || column < 1 || (long)row + table.Rows.Count + (headers ? 1 : 0) - 1 > ExcelA1.MaxRows
            || (long)column + table.Columns.Count - 1 > ExcelA1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(row), "XML table does not fit within the destination worksheet.");
        options.CancellationToken.ThrowIfCancellationRequested();
        InsertDataTable(table, row, column, headers, mode, options.CancellationToken);
        return BuildInsertedRange(table, row, column, headers);
    }

    /// <summary>
    /// Exports a bounded A1 range as a positional or explicitly mapped XML table using cached cell values.
    /// This does not export workbook formatting or calculate formulas.
    /// </summary>
    /// <param name="a1Range">Explicit worksheet range, bounded before reading.</param>
    /// <param name="headersInFirstRow">Uses the first range row as DataTable headers rather than data.</param>
    /// <param name="options">Explicit schema and resource limits.</param>
    /// <param name="readOptions">Existing cell value, date and formula-cache read policy.</param>
    /// <param name="mode">Existing worksheet reading execution mode.</param>
    public string ToXml(string a1Range, bool headersInFirstRow = false, ExcelXmlTableOptions? options = null,
        ExcelReaderOptions? readOptions = null, ExcelExecutionMode? mode = null) {
        var resolved = (options ?? new ExcelXmlTableOptions()).Snapshot();
        if (!ExcelA1.TryParseRange(a1Range, out int firstRow, out int firstColumn, out int lastRow, out int lastColumn)
            || (long)lastRow - firstRow + 1 - (headersInFirstRow ? 1 : 0) > resolved.MaxRows || (long)lastColumn - firstColumn + 1 > resolved.MaxColumns
            || ((long)lastRow - firstRow + 1 - (headersInFirstRow ? 1 : 0)) * ((long)lastColumn - firstColumn + 1) > resolved.MaxCells)
            throw new ArgumentException("Worksheet range is invalid or exceeds XML table limits.", nameof(a1Range));
        using var table = ToDataTable(a1Range, headersInFirstRow, readOptions, mode, resolved.CancellationToken);
        return Encoding.UTF8.GetString(XmlTables.ExcelXmlTableCodec.Write(table, resolved));
    }
}
