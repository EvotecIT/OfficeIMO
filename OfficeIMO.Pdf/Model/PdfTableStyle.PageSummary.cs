namespace OfficeIMO.Pdf;

public partial class PdfTableStyle {
    /// <summary>
    /// When true, the table's trailing footer rows (<see cref="FooterRowCount"/>) close every page fragment of a
    /// table that splits across pages, not only the last one. Defaults to false.
    /// </summary>
    public bool RepeatFooterRows { get; set; }

    /// <summary>
    /// Optional per-page aggregates and summary rows, such as page subtotals and brought-forward or carried-forward
    /// amounts. Applies to flow tables; deferred, positioned and nested tables reject it.
    /// </summary>
    public PdfTablePageSummary? PageSummary { get; set; }
}
