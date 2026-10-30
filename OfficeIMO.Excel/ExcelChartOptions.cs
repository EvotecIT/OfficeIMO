namespace OfficeIMO.Excel;

/// <summary>Controls the title, size, worksheet anchor and native range source of an editable Excel chart.</summary>
public sealed class ExcelChartOptions : OfficeChartOptions {
    /// <summary>Gets or sets the one-based anchor row. Defaults to 1.</summary>
    public int Row { get; set; } = 1;

    /// <summary>Gets or sets the one-based anchor column. Defaults to 1.</summary>
    public int Column { get; set; } = 1;
    /// <summary>Gets or sets whether an A1 source includes a header row. Table sources use their native header setting.</summary>
    public bool HasHeaders { get; set; } = true;

    /// <summary>Gets or sets whether a named range or table source stores a snapshot cache. Formulas remain linked either way.</summary>
    public bool IncludeCachedData { get; set; } = true;

    /// <summary>Gets or sets an explicit cache for a typed native range descriptor. Null preserves a formula-only descriptor.</summary>
    public ExcelChartData? CachedData { get; set; }
}
