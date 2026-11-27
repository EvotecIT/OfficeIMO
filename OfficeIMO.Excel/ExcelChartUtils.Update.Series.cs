using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        // Bar, line, area, radar, surface, pie and doughnut series all take a name, categories and values; a series
        // the data adds is a copy of the last existing one, or a new empty series.
        private static void UpdateCategoryChartSeries(OfficeXmlNode chart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<OfficeXmlNode> existingSeries = chart.Children(C, "ser").ToList();
            OfficeXmlNode template = existingSeries.Count > 0 ? existingSeries[existingSeries.Count - 1] : OfficeXmlNode.None;
            var indexSet = CreateDescriptorIndexSet(descriptors);
            var existingByIndex = CreateSeriesIndexMap(existingSeries);

            foreach (var descriptor in descriptors) {
                if (!existingByIndex.TryGetValue(descriptor.Index, out OfficeXmlNode seriesElement)) {
                    seriesElement = InsertSeries(chart, template.Exists ? template.Clone() : chart.New(C, "ser"));
                }

                UpdateSeriesIndexOrder(seriesElement, descriptor.Index);
                string name = descriptor.Series?.Name ?? $"Series {descriptor.Index + 1}";
                UpdateSeriesText(seriesElement, range, descriptor.Index, name);
                UpdateCategoryAxisData(seriesElement, range, data.Categories);
                UpdateValues(seriesElement, range, descriptor.Index, data.Series[descriptor.Index].Values);
            }

            foreach (var seriesElement in existingSeries) {
                int idx = GetSeriesIndex(seriesElement);
                if (!indexSet.Contains(idx)) {
                    seriesElement.Remove();
                }
            }
        }

        private static void UpdateBarChartSeries(OfficeXmlNode barChart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) =>
            UpdateCategoryChartSeries(barChart, data, range, descriptors);

        private static void UpdateLineChartSeries(OfficeXmlNode lineChart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) =>
            UpdateCategoryChartSeries(lineChart, data, range, descriptors);

        private static void UpdateAreaChartSeries(OfficeXmlNode areaChart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) =>
            UpdateCategoryChartSeries(areaChart, data, range, descriptors);

        private static void UpdateStockChartSeries(OfficeXmlNode stockChart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            if (descriptors.Count < 3 || descriptors.Count > 4) {
                throw new ArgumentException("Stock charts require three series (high-low-close) or four series (open-high-low-close).", nameof(descriptors));
            }

            UpdateCategoryChartSeries(stockChart, data, range, descriptors);
            EnsureStockChartLines(stockChart, descriptors.Count);
        }

        private static void UpdateScatterChartSeries(OfficeXmlNode chart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<OfficeXmlNode> existingSeries = chart.Children(C, "ser").ToList();
            OfficeXmlNode template = existingSeries.Count > 0 ? existingSeries[existingSeries.Count - 1] : OfficeXmlNode.None;
            var indexSet = CreateDescriptorIndexSet(descriptors);
            var existingByIndex = CreateSeriesIndexMap(existingSeries);
            IReadOnlyList<double> categoryXValues = ParseNumericCategories(data.Categories);

            foreach (var descriptor in descriptors) {
                if (!existingByIndex.TryGetValue(descriptor.Index, out OfficeXmlNode seriesElement)) {
                    seriesElement = InsertSeries(chart, template.Exists ? template.Clone() : chart.New(C, "ser"));
                }

                UpdateSeriesIndexOrder(seriesElement, descriptor.Index);
                ExcelChartSeries seriesData = data.Series[descriptor.Index];
                string name = descriptor.Series?.Name ?? $"Series {descriptor.Index + 1}";
                IReadOnlyList<double>? seriesXValues = seriesData.XValues;
                UpdateSeriesText(seriesElement, range, descriptor.Index, name);
                UpdateXValues(seriesElement, range, seriesXValues ?? categoryXValues, ShouldUseLiteralXValues(seriesXValues, categoryXValues));
                UpdateYValues(seriesElement, range, descriptor.Index, seriesData.Values);
            }

            foreach (var seriesElement in existingSeries) {
                int idx = GetSeriesIndex(seriesElement);
                if (!indexSet.Contains(idx)) {
                    seriesElement.Remove();
                }
            }
        }
    }
}
