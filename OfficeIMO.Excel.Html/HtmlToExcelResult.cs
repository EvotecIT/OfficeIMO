using OfficeIMO.Html;

namespace OfficeIMO.Excel.Html;

/// <summary>
/// Summary of a semantic Excel HTML import.
/// </summary>
public sealed class HtmlToExcelResult : HtmlConversionResult<ExcelDocument> {
    internal HtmlToExcelResult(ExcelDocument workbook) : base(workbook) { }

    internal void AddImportDiagnostic(HtmlDiagnostic diagnostic) => AddDiagnostic(diagnostic);

    /// <summary>
    /// Number of imported worksheets.
    /// </summary>
    public int Sheets { get; internal set; }

    /// <summary>
    /// Number of imported worksheet table cells.
    /// </summary>
    public int Cells { get; internal set; }

    /// <summary>
    /// Number of merged worksheet ranges restored from the HTML table grid.
    /// </summary>
    public int MergedRanges { get; internal set; }

    /// <summary>
    /// Number of formulas restored from semantic cell metadata or formula inventory.
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

}
