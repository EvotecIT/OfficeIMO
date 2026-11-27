using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
        //     Represents a chart on a slide.
        public partial class PowerPointChart : PowerPointShape {
        /// <summary>The series type a <c>c:ser</c> is, from the chart group that holds it, as the SDK types it.</summary>
        private enum ChartSeriesKind {
            None,
            Bar,
            Line,
            Area,
            Pie,
            Scatter,
            Radar,
            Bubble
        }

        // The chart groups searched for a series by name, in the order they are searched.
        private static readonly string[] SeriesNameGroups = {
            "barChart", "lineChart", "areaChart", "pieChart", "doughnutChart", "scatterChart", "radarChart", "bubbleChart"
        };

        // The chart groups searched for a series marker by name.
        private static readonly string[] MarkerSeriesGroups = { "lineChart", "scatterChart", "radarChart" };

        private bool ApplySeriesByIndex(int seriesIndex, Action<OfficeXmlNode> apply) {
            OfficeXmlNode plotArea = PlotAreaNode;
            if (plotArea.IsNone) {
                return false;
            }

            // Public indices span the complete chart, including mixed chart groups.
            OfficeXmlNode series = SeriesAt(plotArea, seriesIndex);
            if (series.IsNone) return false;
            apply(series);
            return true;
        }

        private static OfficeXmlNode SeriesAt(OfficeXmlNode plotArea, int seriesIndex) {
            if (seriesIndex < 0) return default;
            var series = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode group in plotArea.Children()) {
                foreach (OfficeXmlNode child in group.Children()) {
                    if (IsSeriesElement(child)) series.Add(child);
                }
            }
            return series.OrderBy(GetSeriesIndex).Skip(seriesIndex).FirstOrDefault();
        }

        private bool ApplySeriesByName(string seriesName, bool ignoreCase, Action<OfficeXmlNode> apply) {
            OfficeXmlNode plotArea = PlotAreaNode;
            if (plotArea.IsNone) {
                return false;
            }

            OfficeXmlNode series = FindSeriesByName(plotArea, SeriesNameGroups, seriesName, ignoreCase);
            if (series.IsNone) return false;
            apply(series);
            return true;
        }

        private bool ApplySeriesMarkerByIndex(int seriesIndex, Action<OfficeXmlNode> apply) {
            OfficeXmlNode plotArea = PlotAreaNode;
            if (plotArea.IsNone) {
                return false;
            }

            OfficeXmlNode series = SeriesAt(plotArea, seriesIndex);
            if (SeriesKind(series) is not ChartSeriesKind.Line and not ChartSeriesKind.Scatter and not ChartSeriesKind.Radar) return false;
            ApplySeriesMarker(series, apply);
            return true;
        }

        private bool ApplySeriesMarkerByName(string seriesName, bool ignoreCase, Action<OfficeXmlNode> apply) {
            OfficeXmlNode plotArea = PlotAreaNode;
            if (plotArea.IsNone) {
                return false;
            }

            OfficeXmlNode series = FindSeriesByName(plotArea, MarkerSeriesGroups, seriesName, ignoreCase);
            if (series.IsNone) return false;
            ApplySeriesMarker(series, apply);
            return true;
        }

        private static void ApplySeriesMarker(OfficeXmlNode series, Action<OfficeXmlNode> apply) {
            OfficeXmlNode marker = series.Child(CNs, "marker");
            if (marker.IsNone) marker = InsertSeriesMarker(series, series.CreateElement(CNs, "marker"));
            apply(marker);
        }

        private static OfficeXmlNode FindSeriesByName(OfficeXmlNode plotArea, string[] groups, string seriesName, bool ignoreCase) {
            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            foreach (string groupName in groups) {
                foreach (OfficeXmlNode chart in plotArea.Children(CNs, groupName)) {
                    foreach (OfficeXmlNode series in chart.Children()) {
                        if (!IsSeriesElement(series)) continue;
                        string? name = GetSeriesName(series);
                        if (name != null && string.Equals(name, seriesName, comparison)) {
                            return series;
                        }
                    }
                }
            }

            return default;
        }

        private static ChartSeriesKind SeriesKind(OfficeXmlNode element) {
            if (!element.Is(CNs, "ser")) return ChartSeriesKind.None;
            OfficeXmlNode group = element.Parent;
            if (group.NamespaceId != CNs) return ChartSeriesKind.None;
            return group.LocalName switch {
                "barChart" or "bar3DChart" => ChartSeriesKind.Bar,
                "lineChart" or "line3DChart" or "stockChart" => ChartSeriesKind.Line,
                "areaChart" or "area3DChart" => ChartSeriesKind.Area,
                "pieChart" or "pie3DChart" or "doughnutChart" or "ofPieChart" => ChartSeriesKind.Pie,
                "scatterChart" => ChartSeriesKind.Scatter,
                "radarChart" => ChartSeriesKind.Radar,
                "bubbleChart" => ChartSeriesKind.Bubble,
                _ => ChartSeriesKind.None
            };
        }

        private static bool IsSeriesElement(OfficeXmlNode element) => SeriesKind(element) != ChartSeriesKind.None;

        private static int GetSeriesIndex(OfficeXmlNode series) {
            // A missing index reads as 0; one that can't be read throws rather than standing in for another series.
            return (int)(OfficeChartXml.UIntVal(series.Child(CNs, "idx")) ?? 0U);
        }

        private static string? GetSeriesName(OfficeXmlNode series) {
            OfficeXmlNode seriesText = series.Child(CNs, "tx");
            if (seriesText.IsNone) {
                return null;
            }

            string? cachedText = seriesText.Child(CNs, "strRef").Child(CNs, "strCache").Child(CNs, "pt").ChildText(CNs, "v");
            if (!string.IsNullOrWhiteSpace(cachedText)) {
                return cachedText;
            }

            string? literalText = seriesText.Child(CNs, "strLit").Child(CNs, "pt").ChildText(CNs, "v");
            if (!string.IsNullOrWhiteSpace(literalText)) {
                return literalText;
            }

            string text = seriesText.InnerText;
            return string.IsNullOrWhiteSpace(text) ? null : text;
        }

        /// <summary>
        /// A detached <c>c:tx</c> holding <paramref name="title"/> as rich text, one run per line, made for a child of
        /// <paramref name="context"/>.
        /// </summary>
        private static OfficeXmlNode CreateChartText(OfficeXmlNode context, string title) {
            OfficeXmlNode chartText = context.CreateElement(CNs, "tx");
            OfficeXmlNode richText = AppendNew(chartText, CNs, "rich");
            AppendNew(richText, ANs, "bodyPr");
            AppendNew(richText, ANs, "lstStyle");
            OfficeXmlNode paragraph = AppendNew(richText, ANs, "p");
            OfficeXmlNode properties = paragraph.CreateElement(ANs, "rPr");
            properties.SetAttribute("lang", "en-US");
            PowerPointTextContent.InsertInlineContent(paragraph, title, properties, default);
            return chartText;
        }

        private static OfficeXmlNode CreateAxisTitle(OfficeXmlNode axis, string title) {
            OfficeXmlNode axisTitle = axis.CreateElement(CNs, "title");
            axisTitle.Append(CreateChartText(axisTitle, title));
            AppendNew(axisTitle, CNs, "layout");
            AppendNew(axisTitle, CNs, "overlay").SetBooleanAttribute("val", false);
            return axisTitle;
        }

        /// <summary>Replaces the first child with this name by a new one appended last, and returns the new one.</summary>
        private static OfficeXmlNode ReplaceChild(OfficeXmlNode parent, string localName) {
            OfficeXmlNode existing = parent.Child(CNs, localName);
            if (existing.Exists) existing.Remove();
            return AppendNew(parent, CNs, localName);
        }

        /// <summary>Replaces the first child with this name by a new one carrying <paramref name="value"/> as its <c>val</c>.</summary>
        private static OfficeXmlNode ReplaceValChild(OfficeXmlNode parent, string localName, string value) {
            OfficeXmlNode child = ReplaceChild(parent, localName);
            child.SetAttribute("val", value);
            return child;
        }

        private static OfficeXmlNode ReplaceValChild(OfficeXmlNode parent, string localName, bool value) =>
            ReplaceValChild(parent, localName, value ? "1" : "0");

        private static OfficeXmlNode ReplaceValChild(OfficeXmlNode parent, string localName, double value) =>
            ReplaceValChild(parent, localName, OfficeXmlViewValues.Format(value)!);

        /// <summary>The schema child order of an axis element.</summary>
        private static OfficeXmlChildOrder AxisOrder(OfficeXmlNode axis) => axis.LocalName switch {
            "catAx" => CCatAxView.Order,
            "dateAx" => CDateAxView.Order,
            "serAx" => CSerAxView.Order,
            _ => CValAxView.Order
        };

        /// <summary>Replaces an axis child with a new one carrying <paramref name="value"/>, placed in schema order.</summary>
        private static void ReplaceAxisChild(OfficeXmlNode axis, string localName, string value) =>
            OfficeChartXml.SetOrdered(axis, axis.CreateVal(CNs, localName, value), AxisOrder(axis));

        // The SDK looked the axis c:spPr up as the chart space's shape properties type, which an axis never holds, so it
        // isn't among the elements new gridlines go before.
        private static OfficeXmlNode InsertAxisGridlines(OfficeXmlNode axis, string gridlinesName) {
            OfficeXmlNode insertBefore = gridlinesName == "majorGridlines" ? axis.Child(CNs, "minorGridlines") : default;
            if (insertBefore.IsNone) {
                insertBefore = FirstChildOf(axis, "title", "numFmt", "majorTickMark", "minorTickMark", "tickLblPos", "txPr", "crossAx", "crosses",
                    "crossesAt", "auto", "lblAlgn", "lblOffset", "noMultiLvlLbl", "crossBetween", "majorUnit", "minorUnit", "dispUnits");
            }
            return InsertBeforeOrAppend(axis, axis.CreateElement(CNs, gridlinesName), insertBefore);
        }

        private static void ReplaceValueAxisCrossBetween(OfficeXmlNode axis, string value) {
            OfficeXmlNode existing = axis.Child(CNs, "crossBetween");
            if (existing.Exists) existing.Remove();
            OfficeXmlNode insertBefore = FirstChildOf(axis, "majorUnit", "minorUnit", "dispUnits");
            InsertBeforeOrAppend(axis, axis.CreateElement(CNs, "crossBetween"), insertBefore).SetAttribute("val", value);
        }
    }
}
