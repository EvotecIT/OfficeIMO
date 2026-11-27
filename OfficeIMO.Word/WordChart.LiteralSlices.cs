using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordChart {
    private OfficeXmlNode _validatedSliceChart;
    private OfficeXmlNode _validatedSliceFamily;
    private OfficeXmlNode _validatedCategoryLiteral;
    private OfficeXmlNode _validatedValueLiteral;
    private OfficeXmlNode _lastAppendedCategoryPoint;
    private OfficeXmlNode _lastAppendedValuePoint;
    private bool _validatedDoughnutFamily;

    /// <summary>Appends a point to a literal before its extension list, and returns it as it is now held.</summary>
    private static OfficeXmlNode AppendLiteralSlicePoint(OfficeXmlNode literal, OfficeXmlNode point) {
        // The extension list is the final schema child. Checking the tail keeps
        // each append constant-time even for large authored literal charts.
        OfficeXmlNode last = literal.LastChild;
        return last.IsNone || last.LocalName != "extLst" ? literal.Append(point) : literal.InsertBefore(point, last);
    }

    // Literal slice appends must use logical cache positions, not the number of stored points.
    // A missing point is a gap; category and value caches can omit different positions.
    private void PrepareLiteralSliceAppend(string category, bool doughnut) {
        if (_chart.IsNone && !ChartSpaceNode.IsNone) _chart = ChartSpaceNode.Child(NsC, "chart");
        if (_chart.IsNone) {
            _currentIndexCategory = 0U;
            _currentIndexValues = 0U;
            return;
        }
        OfficeXmlNode plot = PlotAreaOf(_chart);
        int groups = 0;
        if (!plot.IsNone) {
            foreach (OfficeXmlNode child in plot.Children()) {
                if (IsChartGroup(child) && ++groups == 2) break;
            }
        }
        if (groups != 1)
            throw new NotSupportedException("Slice appends require a plot area with exactly one chart group.");
        OfficeXmlNode family = doughnut ? plot.Child(NsC, "doughnutChart") : plot.Child(NsC, "pieChart");
        if (!doughnut && family.IsNone) family = plot.Child(NsC, "pie3DChart");
        if (family.IsNone)
            throw new NotSupportedException("An existing chart of another family cannot accept these slices.");
        if (family.Children(NsC, "ser").Count > 1)
            throw new NotSupportedException("Use cached-data mutation APIs to edit a chart with multiple rings or series.");
        OfficeXmlNode series = family.Child(NsC, "ser");
        OfficeXmlNode categories = series.IsNone ? default : series.Child(NsC, "cat");
        OfficeXmlNode values = series.IsNone ? default : series.Child(NsC, "val");
        if (!categories.IsNone && (!categories.Child(NsC, "strRef").IsNone || !categories.Child(NsC, "numRef").IsNone ||
                !categories.Child(NsC, "multiLvlStrRef").IsNone) ||
            !values.IsNone && !values.Child(NsC, "numRef").IsNone)
            throw new NotSupportedException("Use cached-data mutation APIs when chart data uses worksheet references.");
        OfficeXmlNode categoryLiteral = categories.IsNone ? default : categories.Child(NsC, "strLit");
        if (categoryLiteral.IsNone && !categories.IsNone) categoryLiteral = categories.Child(NsC, "numLit");
        OfficeXmlNode valueLiteral = values.IsNone ? default : values.Child(NsC, "numLit");
        uint nextCategory = _currentIndexCategory;
        bool numericCategories = !categoryLiteral.IsNone && categoryLiteral.LocalName == "numLit";
        if (nextCategory > 0U && nextCategory == _currentIndexValues &&
            _chart == _validatedSliceChart &&
            family == _validatedSliceFamily &&
            categoryLiteral == _validatedCategoryLiteral &&
            valueLiteral == _validatedValueLiteral &&
            _validatedDoughnutFamily == doughnut &&
            nextCategory < MaxCachedChartPoints &&
            HasAppendedTail(categoryLiteral, _lastAppendedCategoryPoint, nextCategory) &&
            HasAppendedTail(valueLiteral, _lastAppendedValuePoint, nextCategory)) {
            if (numericCategories && !IsNumericSliceCategory(category))
                throw new ArgumentException("A numeric category literal requires a finite invariant numeric category.", nameof(category));
            return;
        }
        uint categoryLength = ReadLiteralSliceLength(categoryLiteral);
        uint valueLength = ReadLiteralSliceLength(valueLiteral);
        if (categoryLength != valueLength)
            throw new InvalidOperationException("Slice category and value caches must have equal logical lengths.");
        if (categoryLength >= MaxCachedChartPoints)
            throw new InvalidOperationException("Slice caches must contain fewer than 10000 logical points before appending.");
        if (numericCategories && !IsNumericSliceCategory(category))
            throw new ArgumentException("A numeric category literal requires a finite invariant numeric category.", nameof(category));
        _currentIndexCategory = categoryLength;
        _currentIndexValues = valueLength;
        _validatedSliceChart = _chart;
        _validatedSliceFamily = family;
        _validatedCategoryLiteral = categoryLiteral;
        _validatedValueLiteral = valueLiteral;
        _validatedDoughnutFamily = doughnut;
    }

    private static bool HasAppendedTail(OfficeXmlNode literal, OfficeXmlNode point, uint nextIndex) {
        if (literal.IsNone || point.IsNone) return false;
        OfficeXmlNode count = literal.Child(NsC, "ptCount");
        return !count.IsNone && count.UIntVal() == nextIndex && point.Parent == literal &&
            point.Is(NsC, "pt") && point.UIntAttr("idx") == nextIndex - 1U;
    }

    private static bool IsNumericSliceCategory(string category) =>
        OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(category, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) &&
        !double.IsNaN(number) && !double.IsInfinity(number);

    private static uint ReadLiteralSliceLength(OfficeXmlNode literal) {
        var indexes = new HashSet<uint>();
        OfficeXmlNode count = literal.IsNone ? default : literal.Child(NsC, "ptCount");
        uint length = (count.IsNone ? null : count.UIntVal()) ?? 0U;
        if (length > MaxCachedChartPoints)
            throw new InvalidOperationException("Slice cache exceeds the supported logical point count.");
        if (literal.IsNone) return length;
        int stored = 0;
        foreach (OfficeXmlNode point in literal.Children(NsC, "pt")) {
            if (++stored > MaxCachedChartPoints)
                throw new InvalidOperationException("Slice cache exceeds the supported stored point count.");
            uint? index = point.UIntAttr("idx");
            if (!index.HasValue || index.Value >= MaxCachedChartPoints || !indexes.Add(index.Value))
                throw new InvalidOperationException("Slice cache contains an invalid or duplicate point index.");
            length = Math.Max(length, index.Value + 1U);
        }
        return length;
    }
}
