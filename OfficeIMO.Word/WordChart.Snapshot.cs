using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using OfficeIMO.OpenXml.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word {
    public partial class WordChart {
        private const uint MaxCachedChartPoints = 10000U;

        /// <summary>
        /// Tries to create a dependency-free chart snapshot from cached Word chart data.
        /// </summary>
        public bool TryGetSnapshot(out WordChartSnapshot snapshot) {
            snapshot = null!;
            try {
                C.Chart? chart = _chartPart?.ChartSpace?.GetFirstChild<C.Chart>() ?? _chart;
                C.PlotArea? plotArea = chart?.GetFirstChild<C.PlotArea>();
                if (chart == null || plotArea == null || !HasSingleSupportedChartElement(plotArea)) return false;
                OpenXmlCompositeElement group = plotArea.ChildElements.OfType<OpenXmlCompositeElement>().Single(IsSupportedChartPlotElement);
                if (!OfficeOpenXmlChartSeriesReader.TryReadKind(group, out OfficeIMO.Drawing.OfficeChartKind kind)) return false;
                A.ColorScheme? scheme = _document.MainDocumentPartRoot.ThemePart?.Theme?.ThemeElements?.ColorScheme;
                OfficeOpenXmlChartSeriesReader.Result? result = group is C.ScatterChart scatter
                    ? OfficeOpenXmlChartSeriesReader.ReadScatter(scatter.Elements<C.ScatterChartSeries>(), scheme, (int)MaxCachedChartPoints)
                    : OfficeOpenXmlChartSeriesReader.ReadCategories(group.ChildElements.OfType<OpenXmlCompositeElement>()
                        .Where(element => element.LocalName == "ser"), kind, scheme, OfficeIMO.Drawing.OfficeChartAxisGroup.Primary, (int)MaxCachedChartPoints);
                if (result == null || result.Series.Any(series => series.HasUnsupportedAppearance)) return false;
                var data = new WordChartData(result.Categories, result.Series.Select(series => new WordChartSeries(series.Data)).ToArray());
                snapshot = CreateSnapshot(chart, ToWordSnapshotKind(kind), data);
                return true;
            } catch {
                snapshot = null!;
                return false;
            }
        }

        private static WordChartSnapshotKind ToWordSnapshotKind(OfficeIMO.Drawing.OfficeChartKind kind) => kind switch {
            OfficeIMO.Drawing.OfficeChartKind.ColumnClustered => WordChartSnapshotKind.ClusteredColumn,
            OfficeIMO.Drawing.OfficeChartKind.ColumnStacked => WordChartSnapshotKind.StackedColumn,
            OfficeIMO.Drawing.OfficeChartKind.ColumnStacked100 => WordChartSnapshotKind.StackedColumn100,
            OfficeIMO.Drawing.OfficeChartKind.BarClustered => WordChartSnapshotKind.ClusteredBar,
            OfficeIMO.Drawing.OfficeChartKind.BarStacked => WordChartSnapshotKind.StackedBar,
            OfficeIMO.Drawing.OfficeChartKind.BarStacked100 => WordChartSnapshotKind.StackedBar100,
            OfficeIMO.Drawing.OfficeChartKind.Line => WordChartSnapshotKind.Line,
            OfficeIMO.Drawing.OfficeChartKind.LineStacked => WordChartSnapshotKind.StackedLine,
            OfficeIMO.Drawing.OfficeChartKind.LineStacked100 => WordChartSnapshotKind.StackedLine100,
            OfficeIMO.Drawing.OfficeChartKind.Area => WordChartSnapshotKind.Area,
            OfficeIMO.Drawing.OfficeChartKind.AreaStacked => WordChartSnapshotKind.StackedArea,
            OfficeIMO.Drawing.OfficeChartKind.AreaStacked100 => WordChartSnapshotKind.StackedArea100,
            OfficeIMO.Drawing.OfficeChartKind.Radar => WordChartSnapshotKind.Radar,
            OfficeIMO.Drawing.OfficeChartKind.Scatter => WordChartSnapshotKind.Scatter,
            OfficeIMO.Drawing.OfficeChartKind.Pie => WordChartSnapshotKind.Pie,
            OfficeIMO.Drawing.OfficeChartKind.Doughnut => WordChartSnapshotKind.Doughnut,
            _ => throw new NotSupportedException("This chart family has no legacy Word snapshot projection.")
        };
        private static bool HasSingleSupportedChartElement(C.PlotArea plotArea) {
            int chartElementCount = 0;
            int supportedChartElementCount = 0;

            foreach (OpenXmlElement child in plotArea.ChildElements) {
                if (!IsChartPlotElement(child)) {
                    continue;
                }

                chartElementCount++;
                if (IsSupportedChartPlotElement(child)) {
                    supportedChartElementCount++;
                }
            }

            return chartElementCount == 1 && supportedChartElementCount == 1;
        }

        private static bool IsChartPlotElement(OpenXmlElement element) =>
            element != null &&
            element.LocalName != null &&
            element.LocalName.EndsWith("Chart", StringComparison.OrdinalIgnoreCase);

        private static bool IsSupportedChartPlotElement(OpenXmlElement element) {
            return element is C.BarChart
                || element is C.Bar3DChart
                || element is C.LineChart
                || element is C.Line3DChart
                || element is C.AreaChart
                || element is C.Area3DChart
                || element is C.RadarChart
                || element is C.ScatterChart
                || element is C.PieChart
                || element is C.Pie3DChart
                || element is C.DoughnutChart;
        }

        private WordChartSnapshot CreateSnapshot(C.Chart chart, WordChartSnapshotKind kind, WordChartData data) {
            return new WordChartSnapshot(
                ReadDrawingName(),
                ReadTitle(chart),
                kind,
                data,
                GetWidthPoints(),
                GetHeightPoints(),
                OfficeOpenXmlChartRadialLayout.Read(chart));
        }

        private static string? ReadTitle(C.Chart chart) {
            C.ChartText? chartText = chart.GetFirstChild<C.Title>()?.GetFirstChild<C.ChartText>();
            if (chartText == null) {
                return null;
            }

            string text = string.Concat(chartText.Descendants<A.Text>().Select(item => item.Text));
            if (!string.IsNullOrWhiteSpace(text)) {
                return text.Trim();
            }

            IReadOnlyList<string> cached = ReadCachedStrings(chartText);
            return cached.Count > 0 && !string.IsNullOrWhiteSpace(cached[0]) ? cached[0].Trim() : null;
        }

        private static IReadOnlyList<string> ReadCachedStrings(OpenXmlElement? container) =>
            OfficeOpenXmlChartCacheReader.ReadCachedStrings(container, (int)MaxCachedChartPoints);
        private string ReadDrawingName() {
            return WordDrawingLayoutReader.TryRead(_drawing, out WordDrawingLayoutSnapshot layout)
                ? layout.Name
                : string.Empty;
        }

        private double GetWidthPoints() {
            return WordDrawingLayoutReader.TryRead(_drawing, out WordDrawingLayoutSnapshot layout)
                ? layout.WidthPoints
                : 450D;
        }

        private double GetHeightPoints() {
            return WordDrawingLayoutReader.TryRead(_drawing, out WordDrawingLayoutSnapshot layout)
                ? layout.HeightPoints
                : 300D;
        }

        private static double EmuToPoints(long emu) {
            return emu * 72D / EnglishMetricUnitsPerInch;
        }
    }
}
