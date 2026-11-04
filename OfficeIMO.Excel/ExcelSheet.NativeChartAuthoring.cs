namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    /// <summary>Adds an editable native chart with per-series chart kinds, axes, styles and optional explicit X values.</summary>
    public ExcelChart AddChart(OfficeChartType type, ExcelChartData data, OfficeChartOptions? options = null) {
        var native = options as ExcelChartOptions;
        return AddChart(data, native?.Row ?? 1, native?.Column ?? 1,
            options?.Width, options?.Height, type, options?.Title);
    }

    /// <summary>Adds an editable chart linked to a worksheet range, including horizontal ranges and ranges on other sheets.</summary>
    /// <remarks>The descriptor defines headers and orientation. ExcelChartOptions.CachedData supplies an optional explicit cache.</remarks>
    public ExcelChart AddChart(OfficeChartType type, ExcelChartDataRange data, OfficeChartOptions? options = null) {
        var native = options as ExcelChartOptions;
        return AddChart(data, native?.Row ?? 1, native?.Column ?? 1,
            options?.Width, options?.Height, type, native?.CachedData, options?.Title);
    }

    /// <summary>Adds an editable chart linked to an A1 range or a native table on this worksheet.</summary>
    /// <remarks>A1 references are parsed first; other strings resolve native table names. Table totals are excluded and native hidden headers are respected.</remarks>
    public ExcelChart AddChart(OfficeChartType type, string rangeOrTable, OfficeChartOptions? options = null) {
        if (string.IsNullOrWhiteSpace(rangeOrTable)) throw new ArgumentNullException(nameof(rangeOrTable));
        var native = options as ExcelChartOptions;
        int row = native?.Row ?? 1;
        int column = native?.Column ?? 1;
        bool includeCache = native?.IncludeCachedData ?? true;
        if (ExcelA1.TryParseRange(rangeOrTable, out _, out _, out _, out _)) {
            return AddChartFromRange(rangeOrTable, row, column, options?.Width, options?.Height,
                type, native?.HasHeaders ?? true, options?.Title, includeCache);
        }
        return AddChartFromTable(rangeOrTable, row, column, options?.Width, options?.Height,
            type, options?.Title, includeCache);
    }
}
