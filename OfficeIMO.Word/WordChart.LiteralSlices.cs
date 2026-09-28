using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word;

public partial class WordChart {
    private Chart? _validatedSliceChart;
    private OpenXmlCompositeElement? _validatedSliceFamily;
    private OpenXmlCompositeElement? _validatedCategoryLiteral;
    private NumberLiteral? _validatedValueLiteral;
    private OpenXmlElement? _lastAppendedCategoryPoint;
    private NumericPoint? _lastAppendedValuePoint;
    private bool _validatedDoughnutFamily;

    private static void AppendLiteralSlicePoint(OpenXmlCompositeElement literal, OpenXmlElement point) {
        // The extension list is the final schema child. Checking the tail keeps
        // each append constant-time even for large authored literal charts.
        OpenXmlElement? last = literal.LastChild;
        if (last is not ExtensionList && last is not StrDataExtensionList) literal.Append(point);
        else literal.InsertBefore(point, last);
    }

    // Literal slice appends must use logical cache positions, not the number of stored points.
    // A missing point is a gap; category and value caches can omit different positions.
    private void PrepareLiteralSliceAppend(string category, bool doughnut) {
        _chart ??= _chartPart?.ChartSpace?.GetFirstChild<Chart>();
        if (_chart == null) {
            _currentIndexCategory = 0U;
            _currentIndexValues = 0U;
            return;
        }
        if (_chart.PlotArea?.ChildElements.Where(element =>
                element.LocalName.EndsWith("Chart", StringComparison.Ordinal)).Take(2).Count() != 1)
            throw new NotSupportedException("Slice appends require a plot area with exactly one chart group.");
        OpenXmlCompositeElement? family = doughnut
            ? _chart.PlotArea?.GetFirstChild<DoughnutChart>()
            : (OpenXmlCompositeElement?)_chart.PlotArea?.GetFirstChild<PieChart>() ??
                _chart.PlotArea?.GetFirstChild<Pie3DChart>();
        if (family == null)
            throw new NotSupportedException("An existing chart of another family cannot accept these slices.");
        if (family.Elements<PieChartSeries>().Skip(1).Any())
            throw new NotSupportedException("Use cached-data mutation APIs to edit a chart with multiple rings or series.");
        PieChartSeries? series = family.GetFirstChild<PieChartSeries>();
        CategoryAxisData? categories = series?.GetFirstChild<CategoryAxisData>();
        Values? values = series?.GetFirstChild<Values>();
        if (categories?.GetFirstChild<StringReference>() != null ||
            categories?.GetFirstChild<NumberReference>() != null ||
            categories?.GetFirstChild<MultiLevelStringReference>() != null ||
            values?.GetFirstChild<NumberReference>() != null)
            throw new NotSupportedException("Use cached-data mutation APIs when chart data uses worksheet references.");
        OpenXmlCompositeElement? categoryLiteral = (OpenXmlCompositeElement?)categories?.GetFirstChild<StringLiteral>() ??
            categories?.GetFirstChild<NumberLiteral>();
        NumberLiteral? valueLiteral = values?.GetFirstChild<NumberLiteral>();
        uint nextCategory = _currentIndexCategory;
        if (nextCategory > 0U && nextCategory == _currentIndexValues &&
            ReferenceEquals(_chart, _validatedSliceChart) &&
            ReferenceEquals(family, _validatedSliceFamily) &&
            ReferenceEquals(categoryLiteral, _validatedCategoryLiteral) &&
            ReferenceEquals(valueLiteral, _validatedValueLiteral) &&
            _validatedDoughnutFamily == doughnut &&
            nextCategory < MaxCachedChartPoints &&
            HasAppendedTail(categoryLiteral, _lastAppendedCategoryPoint, nextCategory) &&
            HasAppendedTail(valueLiteral, _lastAppendedValuePoint, nextCategory)) {
            if (categoryLiteral is NumberLiteral && !IsNumericSliceCategory(category))
                throw new ArgumentException("A numeric category literal requires a finite invariant numeric category.", nameof(category));
            return;
        }
        uint categoryLength = ReadLiteralSliceLength(categoryLiteral);
        uint valueLength = ReadLiteralSliceLength(valueLiteral);
        if (categoryLength != valueLength)
            throw new InvalidOperationException("Slice category and value caches must have equal logical lengths.");
        if (categoryLength >= MaxCachedChartPoints)
            throw new InvalidOperationException("Slice caches must contain fewer than 10000 logical points before appending.");
        if (categoryLiteral is NumberLiteral && !IsNumericSliceCategory(category))
            throw new ArgumentException("A numeric category literal requires a finite invariant numeric category.", nameof(category));
        _currentIndexCategory = categoryLength;
        _currentIndexValues = valueLength;
        _validatedSliceChart = _chart;
        _validatedSliceFamily = family;
        _validatedCategoryLiteral = categoryLiteral;
        _validatedValueLiteral = valueLiteral;
        _validatedDoughnutFamily = doughnut;
    }

    private static bool HasAppendedTail(OpenXmlCompositeElement? literal, OpenXmlElement? point,
        uint nextIndex) =>
        literal?.GetFirstChild<PointCount>()?.Val?.Value == nextIndex &&
        point != null && ReferenceEquals(point.Parent, literal) &&
        (point is StringPoint text ? text.Index?.Value : (point as NumericPoint)?.Index?.Value) == nextIndex - 1U;

    private static bool IsNumericSliceCategory(string category) =>
        double.TryParse(category, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) &&
        !double.IsNaN(number) && !double.IsInfinity(number);

    private static uint ReadLiteralSliceLength(OpenXmlCompositeElement? literal) {
        var indexes = new HashSet<uint>();
        uint length = literal?.GetFirstChild<PointCount>()?.Val?.Value ?? 0U;
        if (length > MaxCachedChartPoints)
            throw new InvalidOperationException("Slice cache exceeds the supported logical point count.");
        if (literal == null) return length;
        int stored = 0;
        foreach (OpenXmlElement point in literal.ChildElements) {
            if (!(point is StringPoint) && !(point is NumericPoint)) continue;
            if (++stored > MaxCachedChartPoints)
                throw new InvalidOperationException("Slice cache exceeds the supported stored point count.");
            uint? index = point is StringPoint text ? text.Index?.Value : ((NumericPoint)point).Index?.Value;
            if (!index.HasValue || index.Value >= MaxCachedChartPoints || !indexes.Add(index.Value))
                throw new InvalidOperationException("Slice cache contains an invalid or duplicate point index.");
            length = Math.Max(length, index.Value + 1U);
        }
        return length;
    }
}
