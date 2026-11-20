using static OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartTextReader;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.OpenXml.Internal;

internal static partial class OfficeOpenXmlChartSeriesReader {
    internal sealed class LabelLayout {
        internal bool Values, Categories, SeriesNames, Percentages, LeaderLines;
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
    internal static LabelLayout ReadLabels(C.Chart chart, A.ColorScheme? scheme = null) {
        LabelLayout? selected = null;
        var labeledSeries = new HashSet<uint>();
        var hiddenPoints = new Dictionary<uint, HashSet<int>>();
        var labeledPoints = new Dictionary<uint, HashSet<int>>();
        bool everySeriesLabeled = true;
        void Select(LabelLayout layout) {
            if (selected == null) selected = layout;
            else if (!selected.SameAs(layout))
                OfficeOpenXmlChartFidelityScope.Approximate("Different label layouts across chart layers cannot be projected together.",
                    "Data labels with different settings per series are drawn with the first labelled series' settings.");
        }
        foreach (var layer in chart.PlotArea?.ChildElements.OfType<OpenXmlCompositeElement>()
            .Where(item => item.LocalName.EndsWith("Chart", StringComparison.Ordinal)) ?? Enumerable.Empty<OpenXmlCompositeElement>()) {
            bool radial = layer is C.PieChart or C.DoughnutChart;
            C.DataLabels? groupLabels = layer.GetFirstChild<C.DataLabels>();
            LabelLayout group = ReadLabelSet(chart, groupLabels, null, radial, scheme);
            if (groupLabels?.Elements<C.DataLabel>().Any() == true)
                OfficeOpenXmlChartFidelityScope.Approximate("The native data label overrides cannot be projected.",
                    "Point labels set on a whole chart group are drawn like the group's other labels.");
            int position = 0;
            foreach (OpenXmlCompositeElement series in layer.ChildElements.OfType<OpenXmlCompositeElement>().Where(item => item.LocalName == "ser")) {
                uint key = series.GetFirstChild<C.Index>()?.Val?.Value ?? (uint)position;
                position++;
                C.DataLabels? seriesLabels = series.GetFirstChild<C.DataLabels>();
                LabelLayout effective = seriesLabels == null ? group : ReadLabelSet(chart, seriesLabels, group, radial, scheme);
                if (effective.Visible) {
                    Select(effective);
                    labeledSeries.Add(key);
                } else {
                    everySeriesLabeled = false;
                }
                foreach (C.DataLabel point in seriesLabels?.Elements<C.DataLabel>() ?? Enumerable.Empty<C.DataLabel>()) {
                    if (point.Index?.Val?.Value is not uint pointIndex || pointIndex > int.MaxValue) continue;
                    if (point.GetFirstChild<C.Delete>() is C.Delete deleted && deleted.Val?.Value != false) {
                        if (effective.Visible) Add(hiddenPoints, key, (int)pointIndex);
                        continue;
                    }
                    if (point.GetFirstChild<C.ChartText>() != null ||
                        point.GetFirstChild<C.Layout>()?.GetFirstChild<C.ManualLayout>() != null)
                        OfficeOpenXmlChartFidelityScope.Approximate("Custom data label text or placement cannot be projected.",
                            "A data label with its own text or position is drawn as a standard label.");
                    LabelLayout own = ReadLabelSet(chart, point, effective, radial, scheme);
                    if (own.Visible == effective.Visible) {
                        if (own.Visible && !own.SameAs(effective))
                            OfficeOpenXmlChartFidelityScope.Approximate("The native data label overrides cannot be projected.",
                                "A data label formatted on its own is drawn like the other labels of its series.");
                        continue;
                    }
                    if (!own.Visible) {
                        Add(hiddenPoints, key, (int)pointIndex);
                    } else {
                        Select(own);
                        Add(labeledPoints, key, (int)pointIndex);
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

    /// <summary>Reads one <c>c:dLbls</c> or <c>c:dLbl</c>; elements it omits come from <paramref name="inherited"/>.</summary>
    private static LabelLayout ReadLabelSet(C.Chart chart, OpenXmlCompositeElement? labels, LabelLayout? inherited, bool radial,
        A.ColorScheme? scheme) {
        if (labels == null) return inherited ?? new LabelLayout();
        if (labels.GetFirstChild<C.Delete>() is C.Delete deleted && deleted.Val?.Value != false) return new LabelLayout();
        var separator = labels.GetFirstChild<C.Separator>();
        string? separatorText = separator == null ? inherited?.Separator : separator.Text;
        if (separator != null && string.IsNullOrEmpty(separator.Text)) {
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
            Values = Flag<C.ShowValue>(labels, inherited?.Values),
            Categories = Flag<C.ShowCategoryName>(labels, inherited?.Categories),
            SeriesNames = Flag<C.ShowSeriesName>(labels, inherited?.SeriesNames),
            Percentages = radial && Flag<C.ShowPercent>(labels, inherited?.Percentages),
            // Leader lines connect radial labels to their slices; Office ignores them elsewhere.
            LeaderLines = radial && Flag<C.ShowLeaderLines>(labels, inherited?.LeaderLines),
            Separator = separatorText,
            NumberFormat = inherited?.NumberFormat,
            Position = inherited?.Position ?? OfficeChartDataLabelPosition.BestFit,
            FontSize = inherited?.FontSize, FontFamily = inherited?.FontFamily, FontStyle = inherited?.FontStyle, TextColor = inherited?.TextColor
        };
        if (labels.GetFirstChild<C.DataLabelPosition>() is C.DataLabelPosition position)
            current.Position = ReadLabelPosition(position.Val?.InnerText);
        if (Flag<C.ShowLegendKey>(labels, null))
            OfficeOpenXmlChartFidelityScope.Approximate("The native data label overrides cannot be projected.",
                "Legend keys inside data labels are left out.");
        if (Flag<C.ShowBubbleSize>(labels, null))
            OfficeOpenXmlChartFidelityScope.Approximate("The native data label overrides cannot be projected.",
                "Bubble sizes in data labels are left out.");
        if (!current.Visible) {
            current.Separator = null;
            current.NumberFormat = null;
            current.Position = OfficeChartDataLabelPosition.BestFit;
            current.LeaderLines = false;
            return current;
        }
        if (radial && current.Position is not OfficeChartDataLabelPosition.BestFit and not OfficeChartDataLabelPosition.OutsideEnd) {
            OfficeOpenXmlChartFidelityScope.Approximate("The native radial data label position cannot be projected.",
                "Radial data labels in that position are drawn at the best-fit position inside their slices.");
            current.Position = OfficeChartDataLabelPosition.BestFit;
        }
        // Best-fit pie labels move outside, with leader lines, when they do not fit their slice; doughnut
        // labels stay on the ring.
        if (current.LeaderLines && current.Position == OfficeChartDataLabelPosition.BestFit && labels.Ancestors<C.PieChart>().Any())
            OfficeOpenXmlChartFidelityScope.Approximate("The native data label overrides cannot be projected.",
                "Best-fit radial data labels with leader lines are drawn inside their slices.");
        ReadLabelText(chart, labels, scheme, current);
        if (labels.GetFirstChild<C.NumberingFormat>() is C.NumberingFormat numbering)
            current.NumberFormat = numbering.SourceLinked?.Value == true
                ? ReadSourceLinkedLabelFormat(labels)
                : OfficeOpenXmlChartSecondaryAxis.ReadDrawableNumberFormat(labels, "data label");
        foreach (var child in labels.ChildElements) {
            if (child is C.Index or C.Delete or C.Layout or C.ChartText or C.TextProperties or C.ShowValue or C.ShowCategoryName or
                C.ShowSeriesName or C.ShowPercent or C.ShowLegendKey or C.ShowBubbleSize or C.ShowLeaderLines or C.Separator or
                C.NumberingFormat or C.DataLabelPosition or C.DataLabel or C.LeaderLines) continue;
            if (child is C.ChartShapeProperties shape) {
                if (!IsUnpaintedLabelShape(shape))
                    OfficeOpenXmlChartFidelityScope.Approximate("The native data label appearance cannot be projected.",
                        "Data label fills and borders are left out.");
                continue;
            }
            if (child is C.DLblsExtensionList or C.DLblExtensionList) {
                if (child.Descendants().Any(element => element.LocalName == "showDataLabelsRange" &&
                        element.GetAttributes().Any(attribute => attribute.LocalName == "val" && attribute.Value is "1" or "true")))
                    OfficeOpenXmlChartFidelityScope.Approximate("Data labels taken from cell ranges cannot be projected.",
                        "Data labels taken from cell ranges are left out of the label text.");
                continue;
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

    private static void ReadLabelText(C.Chart chart, OpenXmlCompositeElement labels, A.ColorScheme? scheme, LabelLayout current) {
        // Labels without their own text properties keep the text their group or series set.
        if (labels.GetFirstChild<C.TextProperties>() == null && (current.FontSize.HasValue || current.FontFamily != null ||
                current.FontStyle.HasValue || current.TextColor.HasValue)) return;
        ChartPart? chartPart = (chart.Parent as C.ChartSpace)?.OpenXmlPart as ChartPart;
        NativeText text = ReadNativeText(chart, labels, scheme);
        NativeText styled = chartPart == null ? default : ReadChartStyleLabelText(chartPart, scheme);
        if (labels.GetFirstChild<C.TextProperties>() is C.TextProperties properties &&
            chartPart?.GetPartsOfType<ChartStylePart>().Any() == true) {
            OpenXmlElement? colorMap = OfficeOpenXmlThemeColorResolver.ResolveChartColorMap(chartPart);
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
        current.FontSize = text.Size ?? styled.Size ?? current.FontSize;
        current.FontFamily = text.Family ?? styled.Family ?? current.FontFamily;
        current.FontStyle = text.Style ?? styled.Style ?? current.FontStyle;
        current.TextColor = text.Color ?? styled.Color ?? current.TextColor;
    }

    private static string? ReadSourceLinkedLabelFormat(OpenXmlCompositeElement labels) {
        // Linked labels show the plotted cells' format, which the series value cache records.
        OpenXmlElement? series = labels.Ancestors().FirstOrDefault(element => element.LocalName == "ser");
        IEnumerable<OpenXmlElement> scopes = series != null ? new[] { series } :
            labels.Parent?.ChildElements.Where(element => element.LocalName == "ser") ?? Enumerable.Empty<OpenXmlElement>();
        string[] codes = scopes.Select(item => (item.GetFirstChild<C.Values>() ?? (OpenXmlElement?)item.GetFirstChild<C.YValues>())?
                .Descendants<C.FormatCode>().FirstOrDefault()?.Text)
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

    private static bool IsUnpaintedLabelShape(C.ChartShapeProperties shape) =>
        !shape.HasAttributes && shape.ChildElements.All(child =>
            child is A.NoFill ||
            child is A.EffectList && !child.HasChildren ||
            child is A.Outline outline && outline.ChildElements.All(line => line is A.NoFill));

    private static bool Flag<T>(OpenXmlCompositeElement labels, bool? inherited) where T : C.BooleanType =>
        labels.GetFirstChild<T>() is T flag ? flag.Val?.Value != false : inherited ?? false;
}
