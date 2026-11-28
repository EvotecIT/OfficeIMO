using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordChart {
    /// <summary>Uses percentage-only pie labels or count-only bar labels, and removes the chart background and frame.</summary>
    /// <param name="percentageLabels">True for percentage-only pie labels; false for numeric count labels.</param>
    /// <param name="titleColor">Text color for the chart title.</param>
    /// <returns>The chart.</returns>
    public WordChart SetPaperStyle(bool percentageLabels = false, OfficeColor? titleColor = null) {
        OfficeXmlNode root = ChartSpaceNode;
        if (root.IsNone) return this;
        OfficeXmlNode frame = root.Child(NsC, "spPr");
        if (frame.IsNone) frame = OfficeChartXml.SetOrdered(root, NewC(root, "spPr"), CChartSpaceView.Order);
        foreach (string fill in new[] { "solidFill", "gradFill", "pattFill", "blipFill", "grpFill", "noFill" }) frame.RemoveChildren(NsA, fill);
        OfficeChartXml.SetOrdered(frame, frame.CreateElement(NsA, "noFill"), DmlShapePropertiesView.Order);
        frame.RemoveChildren(NsA, "ln");
        OfficeXmlNode outline = frame.CreateElement(NsA, "ln");
        AddA(outline, "noFill");
        OfficeChartXml.SetOrdered(frame, outline, DmlShapePropertiesView.Order);
        var titleRunProperties = new List<OfficeXmlNode>();
        foreach (OfficeXmlNode title in root.ContentDescendants(NsC, "title")) titleRunProperties.AddRange(title.ContentDescendants(NsA, "defRPr"));
        foreach (OfficeXmlNode properties in titleRunProperties) {
            properties.RemoveChildren(NsA, "solidFill");
            OfficeChartXml.SetOrdered(properties, CreateChartFill(properties, titleColor ?? OfficeColor.Parse("202428")), DmlTextCharacterPropertiesView.Order);
        }
        foreach (OfficeXmlNode labels in root.ContentDescendants(NsC, "dLbls")) {
            if (!percentageLabels) {
                labels.RemoveChildren(NsC, "dLblPos");
                if (SupportsOutsideEndLabels(labels)) {
                    OfficeChartXml.SetOrdered(labels, NewC(labels, "dLblPos", "outEnd"), CDLblsView.Order);
                }
            }
            labels.RemoveChildren(NsC, "showVal");
            labels.RemoveChildren(NsC, "showPercent");
            labels.RemoveChildren(NsC, "numFmt");
            OfficeChartXml.SetOrdered(labels, NewC(labels, "showVal", !percentageLabels), CDLblsView.Order);
            OfficeChartXml.SetOrdered(labels, NewC(labels, "showPercent", percentageLabels), CDLblsView.Order);
            OfficeXmlNode numberFormat = NewC(labels, "numFmt");
            numberFormat.SetAttribute("formatCode", percentageLabels ? "0.#%" : "0");
            numberFormat.SetAttribute("sourceLinked", OfficeChartXml.Text(false));
            OfficeChartXml.SetOrdered(labels, numberFormat, CDLblsView.Order);
        }
        SaveChartSpace();
        return this;
    }

    /// <summary>Whether data labels sit in a pie group or a clustered bar group, which place labels outside the end.</summary>
    private static bool SupportsOutsideEndLabels(OfficeXmlNode labels) {
        for (OfficeXmlNode ancestor = labels.Parent; !ancestor.IsNone; ancestor = ancestor.Parent) {
            if (ancestor.NamespaceId != NsC) continue;
            if (ancestor.LocalName is "pieChart" or "pie3DChart" or "ofPieChart") return true;
            if (ancestor.LocalName == "barChart" && IsGrouping(ancestor, "clustered")) return true;
        }
        return false;
    }

    private static bool IsGrouping(OfficeXmlNode barChart, string grouping) {
        OfficeXmlNode value = barChart.Child(NsC, "grouping");
        return !value.IsNone && value.Val() == grouping;
    }

    /// <summary>Sets count axes to start at zero, use whole-number labels and a positive integral tick interval.</summary>
    /// <param name="majorUnit">Optional distance between ticks in whole counts. When omitted, a readable integral interval is selected from the cached values.</param>
    /// <returns>The chart.</returns>
    public WordChart SetCountAxis(uint? majorUnit = null) {
        if (majorUnit == 0) throw new ArgumentOutOfRangeException(nameof(majorUnit));
        OfficeXmlNode root = ChartSpaceNode;
        if (root.IsNone) return this;
        double interval = majorUnit.HasValue ? majorUnit.Value : SelectCountAxisInterval(root);
        foreach (OfficeXmlNode axis in PlotAreaOf(root.Child(NsC, "chart")).Children(NsC, "valAx")) {
            OfficeXmlNode scaling = axis.Child(NsC, "scaling");
            if (scaling.IsNone) scaling = OfficeChartXml.SetOrdered(axis, NewC(axis, "scaling"), CValAxView.Order);
            OfficeChartXml.SetOrdered(scaling, NewC(scaling, "min", OfficeChartXml.Text(0D)), CScalingView.Order);
            double dataMaximum = ReadCountMaximum(root, includeAuthoredMaximum: false);
            double requiredMaximum = Math.Ceiling(dataMaximum / interval + 1D) * interval;
            OfficeXmlNode existingMaximum = scaling.Child(NsC, "max");
            double authored = (existingMaximum.IsNone ? null : existingMaximum.DoubleVal()) ?? 0D;
            OfficeChartXml.SetOrdered(scaling, NewC(scaling, "max", OfficeChartXml.Text(Math.Max(authored, requiredMaximum))), CScalingView.Order);
            axis.RemoveChildren(NsC, "numFmt");
            OfficeXmlNode numberFormat = NewC(axis, "numFmt");
            numberFormat.SetAttribute("formatCode", "0");
            numberFormat.SetAttribute("sourceLinked", OfficeChartXml.Text(false));
            OfficeChartXml.SetOrdered(axis, numberFormat, CValAxView.Order);
            axis.RemoveChildren(NsC, "majorUnit");
            OfficeChartXml.SetOrdered(axis, NewC(axis, "majorUnit", OfficeChartXml.Text(interval)), CValAxView.Order);
        }
        SaveChartSpace();
        return this;
    }

    private static double SelectCountAxisInterval(OfficeXmlNode root) {
        double maximum = ReadCountMaximum(root, includeAuthoredMaximum: true);
        // Allow the padded upper bound to retain its interval when formatting is repeated.
        // Data targets ten intervals; authored bounds may use up to twelve, below the reader's limit.
        double dataMaximum = ReadCountMaximum(root, includeAuthoredMaximum: false);
        double target = Math.Max(1D, Math.Max(dataMaximum / 10D, maximum / 12D));
        double scale = Math.Pow(10D, Math.Floor(Math.Log10(target)));
        double normalized = target / scale;
        return Math.Max(1D, (normalized <= 1D ? 1D : normalized <= 2D ? 2D : normalized <= 5D ? 5D : 10D) * scale);
    }

    private static double ReadCountPoint(OfficeXmlNode point) =>
        OfficeIMO.Core.Internal.OfficeRoundTripText.TryParse(point.ChildText(NsC, "v"), NumberStyles.Float, CultureInfo.InvariantCulture, out double value) &&
        !double.IsNaN(value) && !double.IsInfinity(value) ? Math.Max(0D, value) : 0D;

    private static double ReadCountMaximum(OfficeXmlNode root, bool includeAuthoredMaximum) {
        // Only plotted values and axis bounds count; values in extension or vendor markup do not.
        OfficeXmlNode plot = PlotAreaOf(root.Child(NsC, "chart"));
        double maximum = 0D;
        foreach (OfficeXmlNode series in PlottedSeries(plot)) {
            foreach (OfficeXmlNode point in CachedNumbers(series.Child(NsC, "val"))) maximum = Math.Max(maximum, ReadCountPoint(point));
        }
        if (includeAuthoredMaximum) {
            for (OfficeXmlNode axis = plot.IsNone ? default : plot.FirstChild; !axis.IsNone; axis = axis.NextSibling) {
                if (!IsAxis(axis)) continue;
                OfficeXmlNode scaling = axis.Child(NsC, "scaling");
                OfficeXmlNode bound = scaling.IsNone ? default : scaling.Child(NsC, "max");
                if (!bound.IsNone) maximum = Math.Max(maximum, bound.DoubleVal() ?? 0D);
            }
        }
        // Stacked bars use the sum at each category, rather than an individual series maximum.
        foreach (OfficeXmlNode bar in plot.Children(NsC, "barChart")) {
            if (!IsGrouping(bar, "stacked")) continue;
            var sums = new Dictionary<uint, double>();
            foreach (OfficeXmlNode series in bar.Children(NsC, "ser")) {
                foreach (OfficeXmlNode point in CachedNumbers(series.Child(NsC, "val"))) {
                    uint index = point.UIntAttr("idx") ?? 0U;
                    sums[index] = (sums.TryGetValue(index, out double sum) ? sum : 0D) + ReadCountPoint(point);
                }
            }
            if (sums.Count > 0) maximum = Math.Max(maximum, sums.Values.Max());
        }
        return maximum;
    }

    /// <summary>The cached points of a <c>c:val</c>: its <c>c:numRef/c:numCache</c> or <c>c:numLit</c> points.</summary>
    private static IEnumerable<OfficeXmlNode> CachedNumbers(OfficeXmlNode values) {
        if (values.IsNone) yield break;
        OfficeXmlNode reference = values.Child(NsC, "numRef");
        OfficeXmlNode cache = reference.IsNone ? values.Child(NsC, "numLit") : reference.Child(NsC, "numCache");
        if (cache.IsNone) yield break;
        foreach (OfficeXmlNode point in cache.Children(NsC, "pt")) yield return point;
    }
}
