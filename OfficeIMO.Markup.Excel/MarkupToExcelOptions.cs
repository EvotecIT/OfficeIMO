namespace OfficeIMO.Markup.Excel;

/// <summary>Controls conversion of an Office Markup document to an Excel workbook.</summary>
public sealed class MarkupToExcelOptions {
    /// <summary>Whether non-tabular Markdown should be included as worksheet text.</summary>
    public bool IncludeMarkdownAsWorksheetText { get; set; } = true;
    /// <summary>Worksheet name used when the markup does not provide one.</summary>
    public string DefaultSheetName { get; set; } = "Sheet1";
    /// <summary>Default one-based row for chart placement.</summary>
    public int DefaultChartRow { get; set; } = 2;
    /// <summary>Default one-based column for chart placement.</summary>
    public int DefaultChartColumn { get; set; } = 5;
    /// <summary>Default rendered chart width in pixels.</summary>
    public int DefaultChartWidthPixels { get; set; } = 640;
    /// <summary>Default rendered chart height in pixels.</summary>
    public int DefaultChartHeightPixels { get; set; } = 360;
    /// <summary>Whether converted worksheet columns should be auto-fitted.</summary>
    public bool AutoFitColumns { get; set; } = true;
    /// <summary>Whether table header rows should be frozen.</summary>
    public bool FreezeTableHeaderRows { get; set; } = true;
    /// <summary>Whether worksheet gridlines should be hidden.</summary>
    public bool HideGridlines { get; set; } = true;
    /// <summary>Whether Markdown headings should receive worksheet heading styles.</summary>
    public bool StyleMarkdownHeadings { get; set; } = true;
    /// <summary>Whether charts should receive the standard OfficeIMO style.</summary>
    public bool StyleCharts { get; set; } = true;
}
