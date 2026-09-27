using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word;

public partial class WordChart {
    private static void AppendLiteralSlicePoint(OpenXmlCompositeElement literal, OpenXmlElement point) {
        OpenXmlElement? extensions = literal.ChildElements.FirstOrDefault(element =>
            element is ExtensionList || element is StrDataExtensionList);
        if (extensions == null) literal.Append(point);
        else literal.InsertBefore(point, extensions);
    }

    // Literal slice appends must use logical cache positions, not the number of stored points.
    // A missing point is a gap, and category/value caches must describe the same positions.
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
        (uint categoryLength, HashSet<uint> categoryIndexes) = ReadLiteralSlicePositions(categoryLiteral);
        (uint valueLength, HashSet<uint> valueIndexes) = ReadLiteralSlicePositions(valueLiteral);
        if (categoryLength != valueLength || !categoryIndexes.SetEquals(valueIndexes))
            throw new InvalidOperationException("Slice category and value caches must have aligned logical positions.");
        if (categoryLength >= MaxCachedChartPoints)
            throw new InvalidOperationException("Slice caches must contain fewer than 10000 logical points before appending.");
        if (categoryLiteral is NumberLiteral && !IsNumericSliceCategory(category))
            throw new ArgumentException("A numeric category literal requires a finite invariant numeric category.", nameof(category));
        _currentIndexCategory = categoryLength;
        _currentIndexValues = valueLength;
    }

    private static bool IsNumericSliceCategory(string category) =>
        double.TryParse(category, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) &&
        !double.IsNaN(number) && !double.IsInfinity(number);

    private static (uint Length, HashSet<uint> Indexes) ReadLiteralSlicePositions(OpenXmlCompositeElement? literal) {
        var indexes = new HashSet<uint>();
        uint length = literal?.GetFirstChild<PointCount>()?.Val?.Value ?? 0U;
        if (length > MaxCachedChartPoints)
            throw new InvalidOperationException("Slice cache exceeds the supported logical point count.");
        if (literal == null) return (length, indexes);
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
        return (length, indexes);
    }
}
