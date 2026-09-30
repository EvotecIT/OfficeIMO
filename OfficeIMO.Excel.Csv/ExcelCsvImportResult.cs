namespace OfficeIMO.Excel.Csv;

/// <summary>Describes CSV content imported into an Excel worksheet.</summary>
public sealed class ExcelCsvImportResult {
    internal ExcelCsvImportResult(string sheetName, string? tableName, string range, string delimiterText) {
        SheetName = sheetName;
        TableName = tableName;
        Range = range;
        DelimiterText = delimiterText;
    }

    /// <summary>Gets the worksheet containing the imported rows.</summary>
    public string SheetName { get; }

    /// <summary>Gets the actual Excel table name, when a table was created.</summary>
    public string? TableName { get; }

    /// <summary>Gets the occupied A1 range, or an empty string when no cells were written.</summary>
    public string Range { get; }

    /// <summary>Gets the delimiter used to parse the CSV input, including a detected delimiter.</summary>
    public char Delimiter => DelimiterText[0];

    /// <summary>Gets the complete delimiter used to parse the input.</summary>
    public string DelimiterText { get; }
}
