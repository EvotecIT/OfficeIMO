using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    // Represents a chart on a worksheet.
    public sealed partial class ExcelChart {
        // Chart groups whose series are searched, in lookup order.
        private static readonly string[] SeriesChartNames = {
            "barChart", "bar3DChart", "lineChart", "line3DChart", "areaChart", "area3DChart", "pieChart", "pie3DChart",
            "ofPieChart", "doughnutChart", "scatterChart", "bubbleChart", "radarChart", "stockChart", "surface3DChart",
            "surfaceChart"
        };

        // Chart groups whose series take markers, in lookup order.
        private static readonly string[] MarkerSeriesChartNames = {
            "lineChart", "line3DChart", "scatterChart", "radarChart", "stockChart"
        };

        // Elements that follow c:marker in a series, in the order the insertion point is looked up.
        private static readonly string[] SeriesChildrenAfterMarker = {
            "dPt", "dLbls", "trendline", "errBars", "cat", "val", "xVal", "yVal", "bubbleSize", "smooth", "extLst"
        };

        // Elements that follow c:dPt in a series, in the order the insertion point is looked up.
        private static readonly string[] SeriesChildrenAfterDataPoint = {
            "dLbls", "trendline", "errBars", "cat", "val", "xVal", "yVal", "bubbleSize", "smooth", "extLst"
        };

        private bool ApplySeriesByIndex(int seriesIndex, Action<OfficeXmlNode> apply) {
            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return false;
            }

            // Combo layers may contain nonconsecutive native series indexes. Search all
            // layers before using a positional fallback from any individual layer.
            if (ApplySeriesByChartIndex(seriesIndex, apply)) return true;

            foreach (string chartName in SeriesChartNames) {
                if (ApplySeriesByIndex(plotArea.Children(C, chartName), seriesIndex, apply)) return true;
            }

            return false;
        }

        private bool ApplySeriesByChartIndex(int seriesIndex, Action<OfficeXmlNode> apply) {
            OfficeXmlNode series = FindSeriesByChartIndex(seriesIndex);
            if (series.IsNone) return false;
            apply(series);
            return true;
        }

        private OfficeXmlNode FindSeriesByChartIndex(int seriesIndex) {
            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            return plotArea.IsNone ? OfficeXmlNode.None : FindSeriesDescendant(plotArea, seriesIndex);
        }

        // The first series below the element, in document order and outside extension lists, with this c:idx.
        private static OfficeXmlNode FindSeriesDescendant(OfficeXmlNode element, int seriesIndex) =>
            element.FirstDescendantOutsideExtensions(child => IsSeriesElement(child) && GetSeriesIndex(child) == seriesIndex);

        private bool ApplySeriesMarkerByChartIndex(int seriesIndex, Action<OfficeXmlNode> apply) {
            OfficeXmlNode series = FindSeriesByChartIndex(seriesIndex);
            if (series.IsNone || !IsMarkerCapableSeriesElement(series)) return false;

            OfficeXmlNode marker = series.Child(C, "marker");
            if (marker.IsNone) marker = series.New(C, "marker");
            apply(marker);
            OfficeXmlNode insertBefore = FirstExistingChild(series, SeriesChildrenAfterMarker);
            EnsureSeriesChildPosition(series, marker, insertBefore);
            return true;
        }

        private bool ApplySeriesByName(string seriesName, bool ignoreCase, Action<OfficeXmlNode> apply) {
            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return false;
            }

            foreach (string chartName in SeriesChartNames) {
                if (ApplySeriesByName(plotArea.Children(C, chartName), seriesName, ignoreCase, apply)) return true;
            }

            return false;
        }

        private bool ApplySeriesMarkerByIndex(int seriesIndex, Action<OfficeXmlNode> apply) {
            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return false;
            }

            foreach (string chartName in MarkerSeriesChartNames) {
                if (ApplySeriesMarkerByIndex(plotArea.Children(C, chartName), seriesIndex, apply)) return true;
            }

            return false;
        }

        private bool ApplySeriesMarkerByName(string seriesName, bool ignoreCase, Action<OfficeXmlNode> apply) {
            OfficeXmlNode plotArea = ChartNode().Child(C, "plotArea");
            if (plotArea.IsNone) {
                return false;
            }

            foreach (string chartName in MarkerSeriesChartNames) {
                if (ApplySeriesMarkerByName(plotArea.Children(C, chartName), seriesName, ignoreCase, apply)) return true;
            }

            return false;
        }

        private static List<OfficeXmlNode> GetChartSeries(OfficeXmlNode chart) {
            var series = new List<OfficeXmlNode>();
            foreach (OfficeXmlNode child in chart.Children()) {
                if (IsSeriesElement(child)) series.Add(child);
            }
            return series;
        }

        private static bool ApplySeriesByIndex(OfficeXmlNode.ChildList charts, int seriesIndex, Action<OfficeXmlNode> apply) {
            foreach (OfficeXmlNode chart in charts) {
                List<OfficeXmlNode> series = GetChartSeries(chart);

                OfficeXmlNode match = series.Find(s => GetSeriesIndex(s) == seriesIndex);
                if (match.Exists) {
                    apply(match);
                    return true;
                }

                if (seriesIndex < 0 || seriesIndex >= series.Count) {
                    continue;
                }

                apply(series[seriesIndex]);
                return true;
            }

            return false;
        }

        private static bool ApplySeriesByName(OfficeXmlNode.ChildList charts, string seriesName, bool ignoreCase,
            Action<OfficeXmlNode> apply) {
            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            foreach (OfficeXmlNode chart in charts) {
                foreach (OfficeXmlNode series in chart.Children()) {
                    if (!IsSeriesElement(series)) continue;
                    string? name = GetSeriesName(series);
                    if (name != null && string.Equals(name, seriesName, comparison)) {
                        apply(series);
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool ApplySeriesMarkerByIndex(OfficeXmlNode.ChildList charts, int seriesIndex, Action<OfficeXmlNode> apply) {
            foreach (OfficeXmlNode chart in charts) {
                List<OfficeXmlNode> series = GetChartSeries(chart);

                OfficeXmlNode seriesElement = series.Find(s => GetSeriesIndex(s) == seriesIndex);
                if (seriesElement.IsNone) {
                    if (seriesIndex < 0 || seriesIndex >= series.Count) {
                        continue;
                    }
                    seriesElement = series[seriesIndex];
                }

                ApplySeriesMarker(seriesElement, apply);
                return true;
            }

            return false;
        }

        private static bool ApplySeriesMarkerByName(OfficeXmlNode.ChildList charts, string seriesName, bool ignoreCase, Action<OfficeXmlNode> apply) {
            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            foreach (OfficeXmlNode chart in charts) {
                foreach (OfficeXmlNode series in chart.Children()) {
                    if (!IsSeriesElement(series)) continue;
                    string? name = GetSeriesName(series);
                    if (name != null && string.Equals(name, seriesName, comparison)) {
                        ApplySeriesMarker(series, apply);
                        return true;
                    }
                }
            }

            return false;
        }

        // Applies the marker settings and puts c:marker before the series children that follow it, data points aside.
        private static void ApplySeriesMarker(OfficeXmlNode series, Action<OfficeXmlNode> apply) {
            OfficeXmlNode marker = series.Child(C, "marker");
            if (marker.IsNone) marker = series.New(C, "marker");
            apply(marker);
            OfficeXmlNode insertBefore = FirstExistingChild(series, SeriesChildrenAfterDataPoint);
            EnsureSeriesChildPosition(series, marker, insertBefore);
        }

        /// <summary>Whether the element is a <c>c:ser</c> of one of the chart groups.</summary>
        private static bool IsSeriesElement(OfficeXmlNode element) {
            if (!element.Is(C, "ser")) return false;
            OfficeXmlNode parent = element.Parent;
            if (parent.NamespaceId != C) return false;
            return Array.IndexOf(SeriesChartNames, parent.LocalName) >= 0;
        }

        /// <summary>Whether the element is a <c>c:ser</c> directly under a chart group with one of these names.</summary>
        private static bool IsSeriesUnder(OfficeXmlNode element, string chartName, string? otherChartName = null) {
            if (!element.Is(C, "ser")) return false;
            OfficeXmlNode parent = element.Parent;
            return parent.Is(C, chartName) || otherChartName != null && parent.Is(C, otherChartName);
        }

        // Line, 3-D line and stock chart groups hold line series.
        private static bool IsLineSeries(OfficeXmlNode element) =>
            IsSeriesUnder(element, "lineChart", "line3DChart") || IsSeriesUnder(element, "stockChart");

        private static bool IsMarkerCapableSeriesElement(OfficeXmlNode element) =>
            IsLineSeries(element) ||
            IsSeriesUnder(element, "scatterChart") ||
            IsSeriesUnder(element, "radarChart");

        private static int GetSeriesIndex(OfficeXmlNode series) {
            return (int)(series.Child(C, "idx").UInt32Val() ?? 0U);
        }

        // The c:v text of the first c:pt under the element, or null.
        private static string? FirstPointText(OfficeXmlNode points) {
            OfficeXmlNode value = points.Child(C, "pt").Child(C, "v");
            return value.IsNone ? null : value.InnerText;
        }

        private static string? GetSeriesName(OfficeXmlNode series) {
            OfficeXmlNode seriesText = series.Child(C, "tx");
            if (seriesText.IsNone) {
                return null;
            }

            string? cachedText = FirstPointText(seriesText.Child(C, "strRef").Child(C, "strCache"));
            if (!string.IsNullOrWhiteSpace(cachedText)) {
                return cachedText;
            }

            string? literalText = FirstPointText(seriesText.Child(C, "strLit"));
            if (!string.IsNullOrWhiteSpace(literalText)) {
                return literalText;
            }

            string innerText = seriesText.InnerText;
            return string.IsNullOrWhiteSpace(innerText) ? null : innerText;
        }

        /// <summary>A new <c>c:tx</c> holding <paramref name="title"/> as rich text, made for <paramref name="context"/>'s part.</summary>
        private static OfficeXmlNode CreateChartText(OfficeXmlNode context, string title) {
            OfficeXmlNode chartText = context.New(C, "tx");
            FillChartText(chartText, title);
            return chartText;
        }

        // c:tx/c:rich with a:bodyPr, a:lstStyle and one paragraph holding one en-US run of the text.
        private static void FillChartText(OfficeXmlNode chartText, string title) {
            OfficeXmlNode richText = chartText.New(C, "rich");
            OfficeXmlNode paragraph = richText.New(A, "p");
            OfficeXmlNode run = paragraph.New(A, "r");
            run.With(
                run.New(A, "rPr").Attr("lang", "en-US"),
                run.NewText(A, "t", title));
            paragraph.Append(run);
            richText.With(
                richText.New(A, "bodyPr"),
                richText.New(A, "lstStyle"),
                paragraph);
            chartText.Append(richText);
        }

        private static OfficeXmlNode CreateAxisTitle(OfficeXmlNode context, string title) {
            OfficeXmlNode axisTitle = context.New(C, "title");
            FillAxisTitle(axisTitle, title);
            return axisTitle;
        }

        // c:title with the text, a c:layout and c:overlay off.
        private static void FillAxisTitle(OfficeXmlNode axisTitle, string title) {
            axisTitle.With(
                CreateChartText(axisTitle, title),
                axisTitle.New(C, "layout"),
                axisTitle.NewVal(C, "overlay", false));
        }

        /// <summary>Replaces the first child with the same name as <paramref name="child"/>, or appends it; returns it as held.</summary>
        private static OfficeXmlNode ReplaceChild(OfficeXmlNode parent, OfficeXmlNode child) {
            OfficeXmlNode existing = parent.Child(child.NamespaceId, child.LocalName);
            return existing.Exists ? existing.ReplaceWith(child) : parent.Append(child);
        }

        private static string NormalizeHexColor(string hex) {
            hex = hex.Trim();
            if (hex.StartsWith("#", StringComparison.Ordinal)) {
                hex = hex.Substring(1);
            }
            if (hex.Length == 6) return hex.ToUpperInvariant();
            if (hex.Length == 8) return hex.Substring(2).ToUpperInvariant();
            return hex.ToUpperInvariant();
        }
    }
}
