namespace OfficeIMO.Markup.Excel;

public sealed class OfficeMarkupExcelExportOptions {
    public string OutputPath { get; set; } = string.Empty;
    public bool IncludeMarkdownAsWorksheetText { get; set; } = true;
    public string DefaultSheetName { get; set; } = "Sheet1";
    public int DefaultChartRow { get; set; } = 2;
    public int DefaultChartColumn { get; set; } = 5;
    public int DefaultChartWidthPixels { get; set; } = 640;
    public int DefaultChartHeightPixels { get; set; } = 360;
    public bool AutoFitColumns { get; set; } = true;
    public bool FreezeTableHeaderRows { get; set; } = true;
    public bool HideGridlines { get; set; } = true;
    public bool StyleMarkdownHeadings { get; set; } = true;
    public bool StyleCharts { get; set; } = true;
    public bool SafePreflight { get; set; } = true;
    public bool ValidateOpenXml { get; set; } = true;
    public bool SafeRepairDefinedNames { get; set; } = true;
}
