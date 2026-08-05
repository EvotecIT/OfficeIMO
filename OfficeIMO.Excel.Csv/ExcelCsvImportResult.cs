namespace OfficeIMO.Excel.Csv;

/// <summary>Describes CSV content imported into an Excel worksheet.</summary>
public sealed class ExcelCsvImportResult {
    internal ExcelCsvImportResult(string sheetName, string range, char delimiter) {
        SheetName = sheetName;
        Range = range;
        Delimiter = delimiter;
    }

    /// <summary>Gets the worksheet containing the imported rows.</summary>
    public string SheetName { get; }

    /// <summary>Gets the occupied A1 range, or an empty string when no cells were written.</summary>
    public string Range { get; }

    /// <summary>Gets the delimiter used to parse the CSV input, including a detected delimiter.</summary>
    public char Delimiter { get; }
}
