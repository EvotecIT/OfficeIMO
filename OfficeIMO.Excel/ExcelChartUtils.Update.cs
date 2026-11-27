using OfficeIMO.Excel.OpenXml;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        // The plot area children that hold series, counted to tell a combination chart from a single chart.
        private static readonly string[] SeriesChartElementNames = {
            "barChart", "bar3DChart", "lineChart", "line3DChart", "areaChart", "area3DChart", "pieChart", "pie3DChart",
            "ofPieChart", "doughnutChart", "scatterChart", "bubbleChart", "radarChart", "stockChart", "surface3DChart", "surfaceChart"
        };

        internal static void UpdateChartData(ExcelPartRef chartPart, ExcelChartData data, ExcelChartDataRange range) {
            if (chartPart.IsNone) throw new ArgumentNullException(nameof(chartPart));
            ArgumentNullException.ThrowIfNull(data);
            ExcelChartAxisIdGenerator.EnsureInitialized(chartPart);

            OfficeXmlNode plotArea = chartPart.Root.Child(C, "chart").Child(C, "plotArea");
            if (plotArea.IsNone) {
                throw new InvalidOperationException("Chart plot area not found.");
            }

            int chartElementCount = 0;
            foreach (OfficeXmlNode child in plotArea.Children()) {
                if (child.NamespaceId == C && Array.IndexOf(SeriesChartElementNames, child.LocalName) >= 0) chartElementCount++;
            }

            OfficeChartType defaultType = InferChartType(plotArea);
            List<SeriesDescriptor> descriptors = BuildSeriesDescriptors(range, data, defaultType, useSeriesOverrides: chartElementCount > 1);
            ValidateSingleSeriesPieVariants(descriptors);

            if (plotArea.Child(C, "scatterChart") is { Exists: true } scatterChart) {
                if (chartElementCount > 1) {
                    UpdateComboChartData(plotArea, data, range, descriptors);
                } else {
                    UpdateScatterChartSeries(scatterChart, data, range, descriptors);
                }
                return;
            }

            if (plotArea.Child(C, "bubbleChart").Exists) {
                throw new NotSupportedException("Updating bubble charts is not supported. Use range-based charts for bubble data.");
            }
            if (plotArea.Child(C, "stockChart") is { Exists: true } stockChart) {
                if (chartElementCount > 1) {
                    throw new NotSupportedException("Stock charts cannot be updated as combination charts.");
                }
                UpdateStockChartSeries(stockChart, data, range, descriptors);
                return;
            }
            if (plotArea.Child(C, "surface3DChart") is { Exists: true } surface3DChart) {
                if (chartElementCount > 1) {
                    throw new NotSupportedException("Surface charts cannot be updated as combination charts.");
                }
                UpdateCategoryChartSeries(surface3DChart, data, range, descriptors);
                return;
            }
            if (plotArea.Child(C, "surfaceChart") is { Exists: true } surfaceChart) {
                if (chartElementCount > 1) {
                    throw new NotSupportedException("Surface charts cannot be updated as combination charts.");
                }
                UpdateCategoryChartSeries(surfaceChart, data, range, descriptors);
                return;
            }

            if (chartElementCount > 1) {
                UpdateComboChartData(plotArea, data, range, descriptors);
                return;
            }

            if (plotArea.Child(C, "barChart") is { Exists: true } barChart) {
                UpdateBarChartSeries(barChart, data, range, descriptors);
                return;
            }
            if (plotArea.Child(C, "bar3DChart") is { Exists: true } bar3DChart) {
                UpdateCategoryChartSeries(bar3DChart, data, range, descriptors);
                return;
            }
            if (plotArea.Child(C, "lineChart") is { Exists: true } lineChart) {
                UpdateLineChartSeries(lineChart, data, range, descriptors);
                return;
            }
            if (plotArea.Child(C, "line3DChart") is { Exists: true } line3DChart) {
                UpdateCategoryChartSeries(line3DChart, data, range, descriptors);
                return;
            }
            if (plotArea.Child(C, "areaChart") is { Exists: true } areaChart) {
                UpdateAreaChartSeries(areaChart, data, range, descriptors);
                return;
            }
            if (plotArea.Child(C, "area3DChart") is { Exists: true } area3DChart) {
                UpdateCategoryChartSeries(area3DChart, data, range, descriptors);
                return;
            }
            if (plotArea.Child(C, "radarChart") is { Exists: true } radarChart) {
                UpdateCategoryChartSeries(radarChart, data, range, descriptors);
                return;
            }
            if (plotArea.Child(C, "pieChart") is { Exists: true } pieChart) {
                UpdateCategoryChartSeries(pieChart, data, range, descriptors);
                return;
            }
            if (plotArea.Child(C, "pie3DChart") is { Exists: true } pie3DChart) {
                UpdateCategoryChartSeries(pie3DChart, data, range, descriptors);
                return;
            }
            if (plotArea.Child(C, "ofPieChart") is { Exists: true } ofPieChart) {
                UpdateCategoryChartSeries(ofPieChart, data, range, descriptors);
                return;
            }
            if (plotArea.Child(C, "doughnutChart") is { Exists: true } doughnutChart) {
                UpdateCategoryChartSeries(doughnutChart, data, range, descriptors);
                return;
            }

            throw new NotSupportedException("Chart type is not supported for data updates.");
        }
    }
}
