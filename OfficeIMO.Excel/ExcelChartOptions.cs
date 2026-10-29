namespace OfficeIMO.Excel;

/// <summary>Controls the title, size and worksheet anchor of an editable Excel chart.</summary>
public sealed class ExcelChartOptions : OfficeChartOptions {
    /// <summary>Gets or sets the one-based anchor row. Defaults to 1.</summary>
    public int Row { get; set; } = 1;

    /// <summary>Gets or sets the one-based anchor column. Defaults to 1.</summary>
    public int Column { get; set; } = 1;
}
