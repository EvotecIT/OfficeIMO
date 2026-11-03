using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel;

/// <summary>An AutoFilter attached to a worksheet range or its table.</summary>
/// <remarks>Get it from <see cref="ExcelRange.AutoFilter"/>; set criteria per column with <see cref="Column(string)"/>.</remarks>
public sealed class ExcelAutoFilter {
    internal ExcelAutoFilter(ExcelSheet sheet, string range) {
        Sheet = sheet;
        Range = range;
    }

    /// <summary>Gets the worksheet that owns the filter.</summary>
    public ExcelSheet Sheet { get; }

    /// <summary>Gets the A1 range covered by the filter.</summary>
    public string Range { get; }

    /// <summary>Selects the column whose header, in the first row of the filter range, matches <paramref name="header"/>.</summary>
    /// <exception cref="ArgumentException">No header in the first row of the range matches.</exception>
    public ExcelAutoFilterColumn Column(string header) => new ExcelAutoFilterColumn(this, Sheet.ResolveAutoFilterColumnOffset(Range, header));

    /// <summary>Selects a column by its zero-based offset inside the filter range.</summary>
    /// <exception cref="ArgumentOutOfRangeException">The offset is outside the filter range.</exception>
    public ExcelAutoFilterColumn Column(uint columnOffset) {
        var (_, firstColumn, _, lastColumn) = ExcelA1.ParseRange(Range);
        if (columnOffset > (uint)(lastColumn - firstColumn)) throw new ArgumentOutOfRangeException(nameof(columnOffset), "AutoFilter column offset must be inside the filter range.");
        return new ExcelAutoFilterColumn(this, columnOffset);
    }

    /// <summary>Removes the AutoFilter, including its criteria. Returns false when no filter covers the range.</summary>
    public bool Clear() => Sheet.RemoveAutoFilter(Range);

    /// <summary>Reads current criteria and table ownership; returns null after the range's filter is removed.</summary>
    /// <remarks>Creating the handle retains deferred workbook exports. Inspection materializes worksheet state when needed.</remarks>
    public ExcelAutoFilterInfo? Inspect() => Sheet.GetAutoFilters()
        .FirstOrDefault(filter => string.Equals(filter.Range, Range, StringComparison.OrdinalIgnoreCase));
}

/// <summary>One column of an <see cref="ExcelAutoFilter"/>. Each criterion replaces the column's previous criterion.</summary>
public sealed class ExcelAutoFilterColumn {
    internal ExcelAutoFilterColumn(ExcelAutoFilter filter, uint offset) {
        Filter = filter;
        Offset = offset;
    }

    /// <summary>Gets the filter that owns the column.</summary>
    public ExcelAutoFilter Filter { get; }

    /// <summary>Gets the zero-based column offset inside the filter range.</summary>
    public uint Offset { get; }

    /// <summary>Shows rows whose value equals one of <paramref name="values"/> (case-insensitive, duplicates ignored).</summary>
    public ExcelAutoFilter Equals(params string[] values) => Equals((IEnumerable<string>)values);

    /// <summary>Shows rows whose value equals one of <paramref name="values"/> (case-insensitive, duplicates ignored).</summary>
    public ExcelAutoFilter Equals(IEnumerable<string> values) {
        Filter.Sheet.ApplyAutoFilterValueCriteria(Filter.Range, Offset, values);
        return Filter;
    }

    /// <summary>Shows rows whose text contains <paramref name="text"/>.</summary>
    public ExcelAutoFilter Contains(string text) => Custom(FilterOperatorValues.Equal, "*" + RequireText(text) + "*");

    /// <summary>Shows rows whose text does not contain <paramref name="text"/>.</summary>
    public ExcelAutoFilter DoesNotContain(string text) => Custom(FilterOperatorValues.NotEqual, "*" + RequireText(text) + "*");

    /// <summary>Shows rows whose text begins with <paramref name="text"/>.</summary>
    public ExcelAutoFilter BeginsWith(string text) => Custom(FilterOperatorValues.Equal, RequireText(text) + "*");

    /// <summary>Shows rows whose text ends with <paramref name="text"/>.</summary>
    public ExcelAutoFilter EndsWith(string text) => Custom(FilterOperatorValues.Equal, "*" + RequireText(text));

    /// <summary>Shows rows whose number differs from <paramref name="value"/>.</summary>
    public ExcelAutoFilter NotEqual(double value) => Custom(FilterOperatorValues.NotEqual, InvariantNumberText.Get(value));

    /// <summary>Shows rows whose number is at least <paramref name="value"/>.</summary>
    public ExcelAutoFilter GreaterThanOrEqual(double value) => Custom(FilterOperatorValues.GreaterThanOrEqual, InvariantNumberText.Get(value));

    /// <summary>Shows rows whose number is at most <paramref name="value"/>.</summary>
    public ExcelAutoFilter LessThanOrEqual(double value) => Custom(FilterOperatorValues.LessThanOrEqual, InvariantNumberText.Get(value));

    /// <summary>Shows rows whose number lies between the inclusive bounds.</summary>
    public ExcelAutoFilter Between(double minimumInclusive, double maximumInclusive) {
        if (maximumInclusive < minimumInclusive) throw new ArgumentOutOfRangeException(nameof(maximumInclusive), "The maximum value must be greater than or equal to the minimum value.");
        Filter.Sheet.ApplyAutoFilterCustomCriteria(Filter.Range, Offset, matchAll: true, new[] {
            (FilterOperatorValues.GreaterThanOrEqual, InvariantNumberText.Get(minimumInclusive)),
            (FilterOperatorValues.LessThanOrEqual, InvariantNumberText.Get(maximumInclusive))
        });
        return Filter;
    }

    /// <summary>Shows rows whose number lies below <paramref name="minimumExclusive"/> or above <paramref name="maximumExclusive"/>.</summary>
    public ExcelAutoFilter NotBetween(double minimumExclusive, double maximumExclusive) {
        if (maximumExclusive < minimumExclusive) throw new ArgumentOutOfRangeException(nameof(maximumExclusive), "The maximum value must be greater than or equal to the minimum value.");
        Filter.Sheet.ApplyAutoFilterCustomCriteria(Filter.Range, Offset, matchAll: false, new[] {
            (FilterOperatorValues.LessThan, InvariantNumberText.Get(minimumExclusive)),
            (FilterOperatorValues.GreaterThan, InvariantNumberText.Get(maximumExclusive))
        });
        return Filter;
    }

    /// <summary>Shows only rows whose cell in this column is blank.</summary>
    public ExcelAutoFilter Blanks() {
        Filter.Sheet.ApplyAutoFilterBlankCriteria(Filter.Range, Offset);
        return Filter;
    }

    /// <summary>Shows the top or bottom <paramref name="value"/> items, or percent with <paramref name="percent"/>.</summary>
    /// <param name="value">Item count (1-500) or percentage (1-100).</param>
    /// <param name="bottom">True for the lowest values; false for the highest.</param>
    /// <param name="percent">Whether <paramref name="value"/> is a percentage.</param>
    public ExcelAutoFilter TopBottom(ushort value, bool bottom = false, bool percent = false) {
        Filter.Sheet.ApplyAutoFilterTop10Criteria(Filter.Range, Offset, value, isTop: !bottom, isPercent: percent);
        return Filter;
    }

    /// <summary>Removes this column's criterion; the filter stays.</summary>
    public ExcelAutoFilter Clear() {
        Filter.Sheet.ClearAutoFilterColumn(Filter.Range, Offset);
        return Filter;
    }

    private ExcelAutoFilter Custom(FilterOperatorValues filterOperator, string value) {
        Filter.Sheet.ApplyAutoFilterCustomCriteria(Filter.Range, Offset, matchAll: false, new[] { (filterOperator, value) });
        return Filter;
    }

    private static string RequireText(string text) => text ?? throw new ArgumentNullException(nameof(text));
}
