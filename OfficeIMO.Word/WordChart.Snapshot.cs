using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word {
    public partial class WordChart {
        private const uint MaxCachedChartPoints = 10000U;

        /// <summary>
        /// Tries to create a dependency-free chart snapshot from cached Word chart data.
        /// </summary>
        internal bool TryGetSnapshot(out WordChartInfo snapshot) {
            snapshot = null!;
            try {
                OfficeXmlNode chart = CurrentChart;
                OfficeXmlNode plotArea = chart.IsNone ? default : chart.Child(NsC, "plotArea");
                if (chart.IsNone || plotArea.IsNone || !HasSingleSupportedChartElement(plotArea)) return false;
                OfficeXmlNode group = plotArea.Children().ToList().Single(IsSupportedChartPlotElement);
                if (!OfficeOpenXmlChartSeriesReader.TryReadKind(group, out OfficeIMO.OfficeChartType kind)) return false;
                OfficeXmlNode scheme = ThemeColorScheme;
                OfficeOpenXmlChartSeriesReader.SourceVisibility? visibility = null;
                IOfficeChartPart? part = SharedChartPart;
                if (part != null && !OfficeOpenXmlChartSeriesReader.TryReadSourceVisibility(part, chart, out visibility, out _)) return false;
                List<OfficeXmlNode> series = group.Children().ToList().Where(element => element.LocalName == "ser").ToList();
                OfficeOpenXmlChartSeriesReader.Result? result = group.Is(NsC, "scatterChart")
                    ? OfficeOpenXmlChartSeriesReader.ReadScatter(group.Children(NsC, "ser").ToList(), scheme, (int)MaxCachedChartPoints,
                        visibility: visibility)
                    : OfficeOpenXmlChartSeriesReader.ReadCategories(series, kind, scheme, OfficeIMO.Drawing.OfficeChartAxisGroup.Primary, (int)MaxCachedChartPoints,
                        visibility: visibility);
                if (result == null || result.Series.Any(item => item.HasUnsupportedAppearance)) return false;
                var data = new WordChartData(result.Categories, result.Series.Select(item => new WordChartSeries(item.Data)).ToArray());
                snapshot = CreateSnapshot(chart, ToWordSnapshotKind(kind), data);
                return true;
            } catch {
                snapshot = null!;
                return false;
            }
        }

        private static WordChartSnapshotKind ToWordSnapshotKind(OfficeIMO.OfficeChartType kind) => kind switch {
            OfficeIMO.OfficeChartType.ColumnClustered => WordChartSnapshotKind.ClusteredColumn,
            OfficeIMO.OfficeChartType.ColumnStacked => WordChartSnapshotKind.StackedColumn,
            OfficeIMO.OfficeChartType.ColumnStacked100 => WordChartSnapshotKind.StackedColumn100,
            OfficeIMO.OfficeChartType.BarClustered => WordChartSnapshotKind.ClusteredBar,
            OfficeIMO.OfficeChartType.BarStacked => WordChartSnapshotKind.StackedBar,
            OfficeIMO.OfficeChartType.BarStacked100 => WordChartSnapshotKind.StackedBar100,
            OfficeIMO.OfficeChartType.Line => WordChartSnapshotKind.Line,
            OfficeIMO.OfficeChartType.LineStacked => WordChartSnapshotKind.StackedLine,
            OfficeIMO.OfficeChartType.LineStacked100 => WordChartSnapshotKind.StackedLine100,
            OfficeIMO.OfficeChartType.Area => WordChartSnapshotKind.Area,
            OfficeIMO.OfficeChartType.AreaStacked => WordChartSnapshotKind.StackedArea,
            OfficeIMO.OfficeChartType.AreaStacked100 => WordChartSnapshotKind.StackedArea100,
            OfficeIMO.OfficeChartType.Radar => WordChartSnapshotKind.Radar,
            OfficeIMO.OfficeChartType.Scatter => WordChartSnapshotKind.Scatter,
            OfficeIMO.OfficeChartType.Pie => WordChartSnapshotKind.Pie,
            OfficeIMO.OfficeChartType.Doughnut => WordChartSnapshotKind.Doughnut,
            _ => throw new NotSupportedException("This chart family has no legacy Word snapshot projection.")
        };
        private static bool HasSingleSupportedChartElement(OfficeXmlNode plotArea) {
            int chartElementCount = 0;
            int supportedChartElementCount = 0;

            foreach (OfficeXmlNode child in plotArea.Children()) {
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

        private static bool IsChartPlotElement(OfficeXmlNode element) =>
            element.LocalName.EndsWith("Chart", StringComparison.OrdinalIgnoreCase);

        private static bool IsSupportedChartPlotElement(OfficeXmlNode element) =>
            element.NamespaceId == NsC && element.LocalName is "barChart" or "bar3DChart" or "lineChart" or "line3DChart" or
                "areaChart" or "area3DChart" or "radarChart" or "scatterChart" or "pieChart" or "pie3DChart" or "doughnutChart";

        private WordChartInfo CreateSnapshot(OfficeXmlNode chart, WordChartSnapshotKind kind, WordChartData data) {
            return new WordChartInfo(
                ReadDrawingName(),
                ReadTitle(chart),
                kind,
                data,
                GetWidthPoints(),
                GetHeightPoints(),
                OfficeOpenXmlChartRadialLayout.Read(chart));
        }

        private static string? ReadTitle(OfficeXmlNode chart) {
            OfficeXmlNode title = chart.Child(NsC, "title");
            OfficeXmlNode chartText = title.IsNone ? default : title.Child(NsC, "tx");
            if (chartText.IsNone) {
                return null;
            }

            string text = string.Concat(chartText.ContentDescendants(NsA, "t").Select(item => item.InnerText));
            if (!string.IsNullOrWhiteSpace(text)) {
                return text.Trim();
            }

            IReadOnlyList<string> cached = ReadCachedStrings(chartText);
            return cached.Count > 0 && !string.IsNullOrWhiteSpace(cached[0]) ? cached[0].Trim() : null;
        }

        private static IReadOnlyList<string> ReadCachedStrings(OfficeXmlNode container) =>
            OfficeOpenXmlChartCacheReader.ReadCachedStrings(container, (int)MaxCachedChartPoints);
        private string ReadDrawingName() {
            return TryReadDrawingLayout(out WordDrawingLayoutInfo layout)
                ? layout.Name
                : string.Empty;
        }

        private double GetWidthPoints() {
            return TryReadDrawingLayout(out WordDrawingLayoutInfo layout)
                ? layout.WidthPoints
                : 450D;
        }

        private double GetHeightPoints() {
            return TryReadDrawingLayout(out WordDrawingLayoutInfo layout)
                ? layout.HeightPoints
                : 300D;
        }

        private static double EmuToPoints(long emu) {
            return emu * 72D / EnglishMetricUnitsPerInch;
        }
    }
}
