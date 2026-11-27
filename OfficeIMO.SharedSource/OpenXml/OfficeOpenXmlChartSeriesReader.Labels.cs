using static OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartTextReader;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using static OfficeIMO.OpenXml.Internal.OfficeChartXml;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

internal static partial class OfficeOpenXmlChartSeriesReader {
    internal sealed class LabelLayout {
        internal bool Values, Categories, SeriesNames, Percentages, LeaderLines;
        /// <summary>Label parts the renderer leaves out; checked once every override has applied.</summary>
        internal bool LegendKeys, BubbleSizes, RangeText;
        internal double? FontSize;
        internal string? FontFamily;
        internal OfficeFontStyle? FontStyle;
        internal OfficeColor? TextColor;
        internal bool Visible => Values || Categories || SeriesNames || Percentages;
        internal string? Separator, NumberFormat;
        internal OfficeChartDataLabelPosition Position;
        /// <summary>Series (by native <c>c:idx</c>) whose labels are drawn; null draws every series' labels.</summary>
        internal HashSet<uint>? LabeledSeries;
        /// <summary>Points (by native series <c>c:idx</c>) whose own label is deleted.</summary>
        internal Dictionary<uint, HashSet<int>>? HiddenPoints;
        /// <summary>Points (by native series <c>c:idx</c>) labelled although their series is not.</summary>
        internal Dictionary<uint, HashSet<int>>? LabeledPoints;
        internal bool SameAs(LabelLayout other) => Values == other.Values && Categories == other.Categories &&
            SeriesNames == other.SeriesNames && Percentages == other.Percentages && Separator == other.Separator &&
            NumberFormat == other.NumberFormat && Position == other.Position &&
            LeaderLines == other.LeaderLines && FontSize == other.FontSize &&
            string.Equals(FontFamily, other.FontFamily, StringComparison.OrdinalIgnoreCase) &&
            FontStyle == other.FontStyle && Nullable.Equals(TextColor, other.TextColor);

        /// <summary>Maps the native series selection onto the positions of the projected series.</summary>
        internal void ResolveSelection(IReadOnlyList<uint> sourceIndexes, out int[]? series,
            out Dictionary<int, IReadOnlyCollection<int>>? points, out Dictionary<int, IReadOnlyCollection<int>>? hidden) {
            series = null; points = null; hidden = null;
            if (!Visible) return;
            if (LabeledSeries != null)
                series = Enumerable.Range(0, sourceIndexes.Count).Where(index => LabeledSeries.Contains(sourceIndexes[index])).ToArray();
            points = Map(LabeledPoints, sourceIndexes);
            hidden = Map(HiddenPoints, sourceIndexes);
        }

        private static Dictionary<int, IReadOnlyCollection<int>>? Map(Dictionary<uint, HashSet<int>>? source, IReadOnlyList<uint> sourceIndexes) {
            if (source == null || source.Count == 0) return null;
            var map = new Dictionary<int, IReadOnlyCollection<int>>();
            for (int index = 0; index < sourceIndexes.Count; index++)
                if (source.TryGetValue(sourceIndexes[index], out HashSet<int>? set) && set.Count > 0) map[index] = set.ToArray();
            return map.Count == 0 ? null : map;
        }
    }

    /// <summary>
    /// Reads the data labels of every chart group into one shared label layout. Series-level labels
    /// override their group's settings element by element; per-point labels that only delete or repeat
    /// the series label are kept as point selections.
    /// </summary>
    internal static LabelLayout ReadLabels(OfficeXmlNode chart, OfficeXmlNode scheme = default) {
        LabelLayout? selected = null;
        var labeledSeries = new HashSet<uint>();
        var hiddenPoints = new Dictionary<uint, HashSet<int>>();
        var labeledPoints = new Dictionary<uint, HashSet<int>>();
        bool everySeriesLabeled = true;
        void Select(LabelLayout layout, bool radial) {
            Qualify(layout, radial);
            if (selected == null) selected = layout;
            else if (!selected.SameAs(layout))
                OfficeOpenXmlChartFidelityScope.Approximate("Different label layouts across chart layers cannot be projected together.",
                    "Data labels with different settings per series are drawn with the first labelled series' settings.");
        }
        foreach (OfficeXmlNode layer in chart.Child(NsC, "plotArea").Children()) {
            if (!layer.LocalName.EndsWith("Chart", StringComparison.Ordinal)) continue;
            bool radial = layer.Is(NsC, "pieChart") || layer.Is(NsC, "doughnutChart");
            OfficeXmlNode groupLabels = layer.Child(NsC, "dLbls");
            LabelLayout group = ReadLabelSet(chart, groupLabels, null, radial, scheme);
            List<OfficeXmlNode> groupPoints = groupLabels.Children(NsC, "dLbl").ToList();
            int position = 0;
            foreach (OfficeXmlNode series in layer.Children()) {
                if (series.LocalName != "ser") continue;
                uint key = series.Child(NsC, "idx").UIntVal() ?? (uint)position;
                position++;
                OfficeXmlNode seriesLabels = series.Child(NsC, "dLbls");
                LabelLayout effective = seriesLabels.IsNone ? group : ReadLabelSet(chart, seriesLabels, group, radial, scheme);
                if (effective.Visible) {
                    Select(effective, radial);
                    labeledSeries.Add(key);
                } else {
                    everySeriesLabeled = false;
                    // Office still shows labels whose only content is a cell range, a legend key or a bubble size.
                    QualifyLeftOutParts(effective);
                }
                // Group point entries apply to every series; the series' own entries for a point replace them.
                var points = new Dictionary<uint, OfficeXmlNode>();
                foreach (OfficeXmlNode point in groupPoints.Concat(seriesLabels.Children(NsC, "dLbl").ToList()))
                    if (point.Child(NsC, "idx").UIntVal() is uint pointIndex && pointIndex <= int.MaxValue) points[pointIndex] = point;
                foreach (KeyValuePair<uint, OfficeXmlNode> entry in points) {
                    OfficeXmlNode point = entry.Value;
                    int pointIndex = (int)entry.Key;
                    if (point.Child(NsC, "delete") is { IsNone: false } deleted && deleted.BoolVal() != false) {
                        if (effective.Visible) Add(hiddenPoints, key, pointIndex);
                        continue;
                    }
                    if (!point.Child(NsC, "tx").IsNone ||
                        !point.Child(NsC, "layout").Child(NsC, "manualLayout").IsNone)
                        OfficeOpenXmlChartFidelityScope.Approximate("Custom data label text or placement cannot be projected.",
                            "A data label with its own text or position is drawn as a standard label.");
                    LabelLayout own = ReadLabelSet(chart, point, effective, radial, scheme);
                    if (own.Visible == effective.Visible) {
                        if (own.Visible) {
                            Qualify(own, radial);
                            if (!own.SameAs(effective))
                                OfficeOpenXmlChartFidelityScope.Approximate("The native data label overrides cannot be projected.",
                                    "A data label formatted on its own is drawn like the other labels of its series.");
                        }
                        continue;
                    }
                    if (!own.Visible) {
                        Add(hiddenPoints, key, pointIndex);
                    } else {
                        Select(own, radial);
                        Add(labeledPoints, key, pointIndex);
                    }
                }
            }
        }
        if (selected == null) return new LabelLayout();
        if (!everySeriesLabeled) selected.LabeledSeries = labeledSeries;
        selected.HiddenPoints = hiddenPoints.Count == 0 ? null : hiddenPoints;
        selected.LabeledPoints = labeledPoints.Count == 0 ? null : labeledPoints;
        return selected;
    }

    private static void Add(Dictionary<uint, HashSet<int>> map, uint series, int point) {
        if (!map.TryGetValue(series, out HashSet<int>? points)) map[series] = points = new HashSet<int>();
        points.Add(point);
    }

    private static void ApproximateRangeText() =>
        OfficeOpenXmlChartFidelityScope.Approximate("Data labels taken from cell ranges cannot be projected.",
            "Data labels taken from cell ranges are left out of the label text.");

    /// <summary>Checks a visible label set once every override has applied.</summary>
    private static void Qualify(LabelLayout layout, bool radial) {
        QualifyLeftOutParts(layout);
        if (radial && layout.Position is not OfficeChartDataLabelPosition.BestFit and not OfficeChartDataLabelPosition.OutsideEnd) {
            OfficeOpenXmlChartFidelityScope.Approximate("The native radial data label position cannot be projected.",
                "Radial data labels in that position are drawn at the best-fit position inside their slices.");
            layout.Position = OfficeChartDataLabelPosition.BestFit;
        }
    }

    // Legend keys, bubble sizes and cell-range text in labels are left out of the drawn labels.
    private static void QualifyLeftOutParts(LabelLayout layout) {
        if (layout.LegendKeys)
            OfficeOpenXmlChartFidelityScope.Approximate("The native data label overrides cannot be projected.",
                "Legend keys inside data labels are left out.");
        if (layout.BubbleSizes)
            OfficeOpenXmlChartFidelityScope.Approximate("The native data label overrides cannot be projected.",
                "Bubble sizes in data labels are left out.");
        if (layout.RangeText) ApproximateRangeText();
    }

    /// <summary>
    /// Reads one <c>c:dLbls</c> or <c>c:dLbl</c>; elements it omits come from <paramref name="inherited"/>. A hidden
    /// set keeps its position, format and text, so a point that shows its label inherits them.
    /// </summary>
    private static LabelLayout ReadLabelSet(OfficeXmlNode chart, OfficeXmlNode labels, LabelLayout? inherited, bool radial,
        OfficeXmlNode scheme) {
        if (labels.IsNone) return inherited ?? new LabelLayout();
        if (labels.Child(NsC, "delete") is { IsNone: false } deleted && deleted.BoolVal() != false) return new LabelLayout();
        OfficeXmlNode separator = labels.Child(NsC, "separator");
        string? separatorText = separator.IsNone ? inherited?.Separator : separator.InnerText;
        if (!separator.IsNone && string.IsNullOrEmpty(separator.InnerText)) {
            OfficeOpenXmlChartFidelityScope.Approximate("An empty native data label separator cannot be projected.",
                "An empty data label separator is drawn as the default separator.");
            separatorText = null;
        }
        if (separatorText?.Length > OfficeChartLayout.MaximumDataLabelSeparatorCharacters) {
            OfficeOpenXmlChartFidelityScope.Approximate("The native data label separator exceeds the supported rendered-text limit.",
                "A data label separator longer than the rendered-text limit is drawn as the default separator.");
            separatorText = null;
        }
        var current = new LabelLayout {
            Values = Flag(labels, "showVal", inherited?.Values),
            Categories = Flag(labels, "showCatName", inherited?.Categories),
            SeriesNames = Flag(labels, "showSerName", inherited?.SeriesNames),
            Percentages = radial && Flag(labels, "showPercent", inherited?.Percentages),
            // Leader lines connect radial labels to their slices; Office ignores them elsewhere.
            LeaderLines = radial && Flag(labels, "showLeaderLines", inherited?.LeaderLines),
            LegendKeys = Flag(labels, "showLegendKey", inherited?.LegendKeys),
            BubbleSizes = Flag(labels, "showBubbleSize", inherited?.BubbleSizes),
            RangeText = ReadRangeTextFlag(labels) ?? inherited?.RangeText ?? false,
            Separator = separatorText,
            NumberFormat = inherited?.NumberFormat,
            Position = inherited?.Position ?? OfficeChartDataLabelPosition.BestFit,
            FontSize = inherited?.FontSize, FontFamily = inherited?.FontFamily, FontStyle = inherited?.FontStyle, TextColor = inherited?.TextColor
        };
        if (labels.Child(NsC, "dLblPos") is { IsNone: false } position)
            current.Position = ReadLabelPosition(position.Val());
        if (current.Visible || !labels.Child(NsC, "txPr").IsNone) {
            int checkpoint = OfficeOpenXmlChartFidelityScope.Checkpoint;
            try {
                ReadLabelText(chart, labels, scheme, current);
            } catch (NotSupportedException exception) when (OfficeOpenXmlChartFidelityScope.AllowsApproximation) {
                OfficeOpenXmlChartFidelityScope.RollBack(checkpoint);
                OfficeOpenXmlChartFidelityScope.Record("The data label text is drawn with its inherited formatting: " + exception.Message);
            }
        }
        if (labels.Child(NsC, "numFmt") is { IsNone: false } numbering)
            current.NumberFormat = numbering.BoolAttr("sourceLinked") == true
                ? ReadSourceLinkedLabelFormat(labels)
                : OfficeOpenXmlChartSecondaryAxis.ReadDrawableNumberFormat(labels, "data label");
        if (!current.Visible) return current;
        foreach (OfficeXmlNode child in labels.Children()) {
            if (child.NamespaceId == NsC) {
                switch (child.LocalName) {
                    case "idx": case "delete": case "layout": case "tx": case "txPr": case "showVal": case "showCatName":
                    case "showSerName": case "showPercent": case "showLegendKey": case "showBubbleSize": case "showLeaderLines":
                    case "separator": case "numFmt": case "dLblPos": case "dLbl": case "leaderLines":
                        continue;
                    case "spPr":
                        if (!IsUnpaintedLabelShape(child))
                            OfficeOpenXmlChartFidelityScope.Approximate("The native data label appearance cannot be projected.",
                                "Data label fills and borders are left out.");
                        continue;
                    case "extLst":
                        continue;
                }
            }
            OfficeOpenXmlChartFidelityScope.Approximate("The native data label appearance cannot be projected.",
                "Data label formatting the renderer cannot draw is left out.");
        }
        return current;
    }

    private static OfficeChartDataLabelPosition ReadLabelPosition(string? value) => value switch {
        null or "bestFit" => OfficeChartDataLabelPosition.BestFit,
        "b" => OfficeChartDataLabelPosition.Bottom, "t" => OfficeChartDataLabelPosition.Top,
        "l" => OfficeChartDataLabelPosition.Left, "r" => OfficeChartDataLabelPosition.Right,
        "ctr" => OfficeChartDataLabelPosition.Center, "inBase" => OfficeChartDataLabelPosition.InsideBase,
        "inEnd" => OfficeChartDataLabelPosition.InsideEnd, "outEnd" => OfficeChartDataLabelPosition.OutsideEnd,
        _ => ApproximatePosition()
    };

    private static OfficeChartDataLabelPosition ApproximatePosition() {
        OfficeOpenXmlChartFidelityScope.Approximate("The native data label position cannot be projected.",
            "Data labels in an unknown position are drawn at the best-fit position.");
        return OfficeChartDataLabelPosition.BestFit;
    }

    // Label text comes from the label's text properties only; a custom rich-text label is drawn as a standard one.
    // A label that names some text properties keeps the rest from its group or series.
    private static void ReadLabelText(OfficeXmlNode chart, OfficeXmlNode labels, OfficeXmlNode scheme, LabelLayout current) {
        OfficeXmlNode properties = labels.Child(NsC, "txPr");
        bool inherits = current.FontSize.HasValue || current.FontFamily != null || current.FontStyle.HasValue || current.TextColor.HasValue;
        if (properties.IsNone && inherits) return;
        IOfficeChartPart? chartPart = ChartPartOf(chart);
        OfficeXmlNode colorMap = chartPart == null ? default : OfficeOpenXmlThemeColorResolver.ResolveChartColorMap(chartPart);
        if (inherits) {
            NativeText parent = new NativeText(current.FontSize, current.FontStyle, current.TextColor, current.FontFamily);
            NativeText own = ReadTextPropertyDefaults(parent, properties, scheme, colorMap);
            current.FontSize = own.Size;
            current.FontFamily = own.Family;
            current.FontStyle = own.Style;
            current.TextColor = own.Color;
            return;
        }
        NativeText chartText = ReadTextPropertyDefaults(default, chart.Parent.Child(NsC, "txPr"), scheme, colorMap);
        NativeText text = ReadTextPropertyDefaults(chartText, properties, scheme, colorMap);
        // A hidden label set only passes its text on to points that show their labels; the chart style's
        // label text applies to labels that are drawn.
        NativeText styled = chartPart == null || !current.Visible ? default : ReadChartStyleLabelText(chartPart, scheme);
        if (current.Visible && !properties.IsNone && chartPart != null && chartPart.ChartStyles(1).Count > 0) {
            NativeText explicitText = ReadTextPropertyDefaults(default, properties, scheme, colorMap);
            NativeText inheritedText = ReadTextPropertyDefaults(styled, properties, scheme, colorMap);
            if (styled.Size.HasValue && !explicitText.Size.HasValue ||
                styled.Style.HasValue && !explicitText.Style.HasValue ||
                styled.Style.HasValue && inheritedText.Style != explicitText.Style ||
                styled.Color.HasValue && !explicitText.Color.HasValue ||
                styled.Family != null && explicitText.Family == null)
                OfficeOpenXmlChartFidelityScope.Approximate("Styled chart data-label text inheritance cannot be projected.",
                    "Data label text inherited from the chart style is drawn with the label's own text settings.");
        }
        current.FontSize = text.Size ?? styled.Size;
        current.FontFamily = text.Family ?? styled.Family;
        current.FontStyle = text.Style ?? styled.Style;
        current.TextColor = text.Color ?? styled.Color;
    }

    private static string? ReadSourceLinkedLabelFormat(OfficeXmlNode labels) {
        // Linked labels show the plotted cells' format, which the series value cache records.
        OfficeXmlNode series = default;
        for (OfficeXmlNode ancestor = labels.Parent; !ancestor.IsNone; ancestor = ancestor.Parent) {
            if (ancestor.LocalName == "ser") {
                series = ancestor;
                break;
            }
        }
        IEnumerable<OfficeXmlNode> scopes = !series.IsNone ? new[] { series } :
            labels.Parent.Children().ToList().Where(element => element.LocalName == "ser");
        string[] codes = scopes.Select(item => {
                OfficeXmlNode values = item.Child(NsC, "val");
                if (values.IsNone) values = item.Child(NsC, "yVal");
                OfficeXmlNode code = values.FirstDescendant(NsC, "formatCode");
                return code.IsNone ? null : code.InnerText;
            })
            .Select(code => string.IsNullOrWhiteSpace(code) ? "General" : code!).Distinct(StringComparer.Ordinal).ToArray();
        if (codes.Length > 1)
            OfficeOpenXmlChartFidelityScope.Approximate("Source-linked data label formats require qualified workbook formatting.",
                "Linked data labels whose series use different formats are drawn with the first series' format.");
        string format = codes.FirstOrDefault() ?? "General";
        if (string.Equals(format, "General", StringComparison.OrdinalIgnoreCase)) return null;
        if (OfficeOpenXmlChartSecondaryAxis.IsDrawableNumberFormat(format)) return format;
        OfficeOpenXmlChartFidelityScope.Record("The data label number format '" + format + "' is drawn as General.");
        return null;
    }

    private static bool IsUnpaintedLabelShape(OfficeXmlNode shape) {
        if (shape.HasAttributes) return false;
        foreach (OfficeXmlNode child in shape.Children()) {
            if (child.Is(NsA, "noFill")) continue;
            if (child.Is(NsA, "effectLst") && !child.HasChildElements()) continue;
            if (child.Is(NsA, "ln")) {
                bool unpainted = true;
                foreach (OfficeXmlNode line in child.Children()) {
                    if (!line.Is(NsA, "noFill")) {
                        unpainted = false;
                        break;
                    }
                }
                if (unpainted) continue;
            }
            return false;
        }
        return true;
    }

    // The c15 showDataLabelsRange flag in the label extension list; null when the labels do not set it.
    private static bool? ReadRangeTextFlag(OfficeXmlNode labels) {
        OfficeXmlNode flag = default;
        foreach (OfficeXmlNode list in labels.Children(NsC, "extLst")) {
            flag = list.DescendantsWhere(element => element.LocalName == "showDataLabelsRange").FirstOrDefault();
            if (!flag.IsNone) break;
        }
        if (flag.IsNone) return null;
        string? value = flag.Attributes().FirstOrDefault(attribute => attribute.LocalName == "val").Value;
        return value is null or "1" or "true";
    }

    private static bool Flag(OfficeXmlNode labels, string localName, bool? inherited) =>
        labels.Child(NsC, localName) is { IsNone: false } flag ? flag.BoolVal() != false : inherited ?? false;

    // SDK callers
    internal static LabelLayout ReadLabels(C.Chart chart, A.ColorScheme? scheme = null) =>
        ReadLabels(OfficeXmlNode.FromSdk(chart), OfficeXmlNode.FromSdk(scheme));
}
