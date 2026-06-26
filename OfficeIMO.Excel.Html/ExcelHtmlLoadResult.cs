namespace OfficeIMO.Excel.Html;

/// <summary>
/// Summary of a semantic Excel HTML import.
/// </summary>
public sealed class ExcelHtmlLoadResult {
    internal ExcelHtmlLoadResult(ExcelDocument workbook) {
        Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
    }

    /// <summary>
    /// Imported workbook.
    /// </summary>
    public ExcelDocument Workbook { get; }

    /// <summary>
    /// Number of imported worksheets.
    /// </summary>
    public int Sheets { get; internal set; }

    /// <summary>
    /// Number of imported worksheet table cells.
    /// </summary>
    public int Cells { get; internal set; }

    /// <summary>
    /// Number of formulas restored from semantic formula inventory.
    /// </summary>
    public int Formulas { get; internal set; }

    /// <summary>
    /// Number of comments restored from semantic comment inventory.
    /// </summary>
    public int Comments { get; internal set; }

    /// <summary>
    /// Number of embedded images restored from semantic image inventory.
    /// </summary>
    public int Images { get; internal set; }

    /// <summary>
    /// Number of chart inventory items restored as native charts.
    /// </summary>
    public int Charts { get; internal set; }

    /// <summary>
    /// Import diagnostics for skipped or approximate rich content.
    /// </summary>
    public IList<string> Diagnostics { get; } = new List<string>();
}
