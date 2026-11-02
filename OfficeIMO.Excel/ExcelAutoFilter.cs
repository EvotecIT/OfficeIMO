namespace OfficeIMO.Excel;

/// <summary>An AutoFilter attached to a worksheet range or its table.</summary>
public sealed class ExcelAutoFilter {
    internal ExcelAutoFilter(ExcelSheet sheet, string range) {
        Sheet = sheet;
        Range = range;
    }

    /// <summary>Gets the worksheet that owns the filter.</summary>
    public ExcelSheet Sheet { get; }

    /// <summary>Gets the A1 range covered by the filter.</summary>
    public string Range { get; }

    /// <summary>Reads current criteria and table ownership; returns null after the range's filter is removed.</summary>
    /// <remarks>Creating the handle retains deferred workbook exports. Inspection materializes worksheet state when needed.</remarks>
    public ExcelAutoFilterInfo? Inspect() => Sheet.GetAutoFilters()
        .FirstOrDefault(filter => string.Equals(filter.Range, Range, StringComparison.OrdinalIgnoreCase));
}
