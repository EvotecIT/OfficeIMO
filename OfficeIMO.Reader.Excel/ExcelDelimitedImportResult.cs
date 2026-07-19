namespace OfficeIMO.Reader.Excel;

/// <summary>Describes a CSV or TSV import into an Excel workbook.</summary>
public sealed class ExcelDelimitedImportResult {
    internal ExcelDelimitedImportResult(string sheetName, string? tableName, string range, int rowCount, int columnCount, char delimiter, IReadOnlyList<string> warnings) {
        SheetName = sheetName;
        TableName = tableName;
        Range = range;
        RowCount = rowCount;
        ColumnCount = columnCount;
        Delimiter = delimiter;
        Warnings = warnings;
    }

    /// <summary>Gets the created worksheet name.</summary>
    public string SheetName { get; }

    /// <summary>Gets the created Excel table name, when a table was requested.</summary>
    public string? TableName { get; }

    /// <summary>Gets the imported A1 range.</summary>
    public string Range { get; }

    /// <summary>Gets the number of imported data rows.</summary>
    public int RowCount { get; }

    /// <summary>Gets the number of imported columns.</summary>
    public int ColumnCount { get; }

    /// <summary>Gets the delimiter used for parsing.</summary>
    public char Delimiter { get; }

    /// <summary>Gets non-fatal import warnings.</summary>
    public IReadOnlyList<string> Warnings { get; }
}
