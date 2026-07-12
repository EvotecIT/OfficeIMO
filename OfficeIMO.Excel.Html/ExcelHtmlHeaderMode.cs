namespace OfficeIMO.Excel.Html;

/// <summary>
/// Controls which worksheet rows are represented as HTML column headers.
/// </summary>
public enum ExcelHtmlHeaderMode {
    /// <summary>Exports every worksheet row as table data.</summary>
    None,

    /// <summary>Exports the first used-range row as a column header row.</summary>
    FirstRow
}
