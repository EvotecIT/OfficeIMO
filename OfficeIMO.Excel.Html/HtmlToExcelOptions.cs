namespace OfficeIMO.Excel.Html;

/// <summary>
/// Options for importing semantic OfficeIMO Excel HTML into a workbook.
/// </summary>
public sealed class HtmlToExcelOptions {
    /// <summary>
    /// Maximum number of HTML table grid slots imported across each worksheet table, including merged spans.
    /// </summary>
    public int MaxTableCells { get; set; } = 50_000;

    /// <summary>
    /// Imports embedded data URI images from the semantic image inventory.
    /// </summary>
    public bool ImportImages { get; set; } = true;

    /// <summary>
    /// Imports chart inventory items as native charts when the worksheet table contains enough data.
    /// </summary>
    public bool ImportChartInventory { get; set; } = true;

    /// <summary>
    /// Imports comment inventory items as native cell comments.
    /// </summary>
    public bool ImportComments { get; set; } = true;

    /// <summary>
    /// Imports formula inventory items as native cell formulas.
    /// </summary>
    public bool ImportFormulas { get; set; } = true;
}
