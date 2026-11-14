namespace OfficeIMO.Excel;

public partial class ExcelSheet {
    /// <summary>Adds an editable chart whose series reference their own worksheet ranges, which need not be adjacent.</summary>
    /// <remarks>
    /// Category charts take the categories from each series' X range; every series must use the same category range and a
    /// values range with the same number of cells. Ranges without a sheet name refer to this worksheet. A cached snapshot is
    /// read from the referenced cells unless ExcelChartOptions.CachedData supplies one or IncludeCachedData is false.
    /// Scatter and bubble charts use the X, Y and bubble-size ranges as given and store formulas only.
    /// </remarks>
    public ExcelChart AddChartFromRanges(OfficeChartType type, IEnumerable<ExcelChartSeriesRange> series, OfficeChartOptions? options = null) {
        ArgumentNullException.ThrowIfNull(series);
        var native = options as ExcelChartOptions;
        int row = native?.Row ?? 1;
        int column = native?.Column ?? 1;
        OfficeLength width = options?.Width ?? OfficeLength.Pixels(640);
        OfficeLength height = options?.Height ?? OfficeLength.Pixels(360);
        var ranges = series as IReadOnlyList<ExcelChartSeriesRange> ?? series.ToList();
        if (type == OfficeChartType.Scatter || type == OfficeChartType.Bubble) {
            return AddChartFromSeriesRanges(ranges, row, column, width, height, type, options?.Title);
        }

        ValidateChartPlacement(row, column, width, height);
        // Like the chart-data sheet path, only a drawing is added, so bulk-written rows keep their direct save model.
        using var preserveFastSaveState = _excelDocument.PreserveDirectDataSetFastSaveStateDuringDirtyMarks();
        ExcelChartRangeSeries rangeSeries = ExcelChartRangeSeries.Resolve(this, ranges);
        ExcelChartData? data = native?.CachedData;
        if (data != null) {
            rangeSeries.ValidateCache(data);
        } else if (native?.IncludeCachedData ?? true) {
            data = rangeSeries.ReadCache();
        }

        // The range descriptor only shapes the series; ApplyFormulas then points each series at its own ranges.
        var shape = new ExcelChartDataRange(Name, 1, 1, rangeSeries.PointCount, ranges.Count, hasHeaderRow: false);
        return AddChartInternal(shape, row, column, width, height, type, data, options?.Title, preserveDeferredDataSetCandidate: true,
            nativeSeriesNames: rangeSeries.ResolveNames(data), rangeSeries: rangeSeries);
    }
}
