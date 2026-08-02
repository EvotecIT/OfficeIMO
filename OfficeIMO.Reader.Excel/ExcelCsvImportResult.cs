namespace OfficeIMO.Reader.Excel;

/// <summary>Describes CSV content imported into an Excel worksheet.</summary>
public sealed class ExcelCsvImportResult {
    internal ExcelCsvImportResult(string sheetName, string range) {
        SheetName = sheetName;
        Range = range;
    }

    /// <summary>Gets the worksheet containing the imported rows.</summary>
    public string SheetName { get; }

    /// <summary>Gets the occupied A1 range, or an empty string when no cells were written.</summary>
    public string Range { get; }
}
