using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using ChartIndex = DocumentFormat.OpenXml.Drawing.Charts.Index;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        private static void UpdateBarChartSeries(BarChart barChart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<BarChartSeries> existingSeries = barChart.Elements<BarChartSeries>().ToList();
            BarChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = CreateDescriptorIndexSet(descriptors);
            var existingByIndex = CreateSeriesIndexMap(existingSeries);

            foreach (var descriptor in descriptors) {
                BarChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (BarChartSeries)template.CloneNode(true) : new BarChartSeries();
                    InsertSeries(barChart, seriesElement);
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

        private static void UpdateBar3DChartSeries(Bar3DChart barChart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<BarChartSeries> existingSeries = barChart.Elements<BarChartSeries>().ToList();
            BarChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = CreateDescriptorIndexSet(descriptors);
            var existingByIndex = CreateSeriesIndexMap(existingSeries);

            foreach (var descriptor in descriptors) {
                BarChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (BarChartSeries)template.CloneNode(true) : new BarChartSeries();
                    InsertSeries(barChart, seriesElement);
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

        private static void UpdateLineChartSeries(LineChart lineChart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<LineChartSeries> existingSeries = lineChart.Elements<LineChartSeries>().ToList();
            LineChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = CreateDescriptorIndexSet(descriptors);
            var existingByIndex = CreateSeriesIndexMap(existingSeries);

            foreach (var descriptor in descriptors) {
                LineChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (LineChartSeries)template.CloneNode(true) : new LineChartSeries();
                    InsertSeries(lineChart, seriesElement);
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

        private static void UpdateLine3DChartSeries(Line3DChart lineChart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<LineChartSeries> existingSeries = lineChart.Elements<LineChartSeries>().ToList();
            LineChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = CreateDescriptorIndexSet(descriptors);
            var existingByIndex = CreateSeriesIndexMap(existingSeries);

            foreach (var descriptor in descriptors) {
                LineChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (LineChartSeries)template.CloneNode(true) : new LineChartSeries();
                    InsertSeries(lineChart, seriesElement);
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

        private static void UpdateAreaChartSeries(AreaChart areaChart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<AreaChartSeries> existingSeries = areaChart.Elements<AreaChartSeries>().ToList();
            AreaChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = CreateDescriptorIndexSet(descriptors);
            var existingByIndex = CreateSeriesIndexMap(existingSeries);

            foreach (var descriptor in descriptors) {
                AreaChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (AreaChartSeries)template.CloneNode(true) : new AreaChartSeries();
                    InsertSeries(areaChart, seriesElement);
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

        private static void UpdateArea3DChartSeries(Area3DChart areaChart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<AreaChartSeries> existingSeries = areaChart.Elements<AreaChartSeries>().ToList();
            AreaChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = CreateDescriptorIndexSet(descriptors);
            var existingByIndex = CreateSeriesIndexMap(existingSeries);

            foreach (var descriptor in descriptors) {
                AreaChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (AreaChartSeries)template.CloneNode(true) : new AreaChartSeries();
                    InsertSeries(areaChart, seriesElement);
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

        private static void UpdateRadarChartSeries(RadarChart radarChart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<RadarChartSeries> existingSeries = radarChart.Elements<RadarChartSeries>().ToList();
            RadarChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = CreateDescriptorIndexSet(descriptors);
            var existingByIndex = CreateSeriesIndexMap(existingSeries);

            foreach (var descriptor in descriptors) {
                RadarChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (RadarChartSeries)template.CloneNode(true) : new RadarChartSeries();
                    InsertSeries(radarChart, seriesElement);
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

        private static void UpdateStockChartSeries(StockChart stockChart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            if (descriptors.Count < 3 || descriptors.Count > 4) {
                throw new ArgumentException("Stock charts require three series (high-low-close) or four series (open-high-low-close).", nameof(descriptors));
            }

            List<LineChartSeries> existingSeries = stockChart.Elements<LineChartSeries>().ToList();
            LineChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = CreateDescriptorIndexSet(descriptors);
            var existingByIndex = CreateSeriesIndexMap(existingSeries);

            foreach (var descriptor in descriptors) {
                LineChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (LineChartSeries)template.CloneNode(true) : new LineChartSeries();
                    InsertSeries(stockChart, seriesElement);
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

            EnsureStockChartLines(stockChart, descriptors.Count);
        }

        private static void UpdateSurfaceChartSeries(OpenXmlCompositeElement surfaceChart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<SurfaceChartSeries> existingSeries = surfaceChart.Elements<SurfaceChartSeries>().ToList();
            SurfaceChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = CreateDescriptorIndexSet(descriptors);
            var existingByIndex = CreateSeriesIndexMap(existingSeries);

            foreach (var descriptor in descriptors) {
                SurfaceChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (SurfaceChartSeries)template.CloneNode(true) : new SurfaceChartSeries();
                    InsertSeries(surfaceChart, seriesElement);
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

        private static void UpdatePieChartSeries(PieChart chart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<PieChartSeries> existingSeries = chart.Elements<PieChartSeries>().ToList();
            PieChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = CreateDescriptorIndexSet(descriptors);
            var existingByIndex = CreateSeriesIndexMap(existingSeries);

            foreach (var descriptor in descriptors) {
                PieChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (PieChartSeries)template.CloneNode(true) : new PieChartSeries();
                    InsertSeries(chart, seriesElement);
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

        private static void UpdatePie3DChartSeries(Pie3DChart chart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<PieChartSeries> existingSeries = chart.Elements<PieChartSeries>().ToList();
            PieChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = CreateDescriptorIndexSet(descriptors);
            var existingByIndex = CreateSeriesIndexMap(existingSeries);

            foreach (var descriptor in descriptors) {
                PieChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (PieChartSeries)template.CloneNode(true) : new PieChartSeries();
                    InsertSeries(chart, seriesElement);
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

        private static void UpdateOfPieChartSeries(OfPieChart chart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<PieChartSeries> existingSeries = chart.Elements<PieChartSeries>().ToList();
            PieChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = CreateDescriptorIndexSet(descriptors);
            var existingByIndex = CreateSeriesIndexMap(existingSeries);

            foreach (var descriptor in descriptors) {
                PieChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (PieChartSeries)template.CloneNode(true) : new PieChartSeries();
                    InsertSeries(chart, seriesElement);
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

        private static void UpdateDoughnutChartSeries(DoughnutChart chart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<PieChartSeries> existingSeries = chart.Elements<PieChartSeries>().ToList();
            PieChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = CreateDescriptorIndexSet(descriptors);
            var existingByIndex = CreateSeriesIndexMap(existingSeries);

            foreach (var descriptor in descriptors) {
                PieChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (PieChartSeries)template.CloneNode(true) : new PieChartSeries();
                    InsertSeries(chart, seriesElement);
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

        private static void UpdateScatterChartSeries(ScatterChart chart, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            List<ScatterChartSeries> existingSeries = chart.Elements<ScatterChartSeries>().ToList();
            ScatterChartSeries? template = existingSeries.LastOrDefault();
            var indexSet = CreateDescriptorIndexSet(descriptors);
            var existingByIndex = CreateSeriesIndexMap(existingSeries);
            IReadOnlyList<double> categoryXValues = ParseNumericCategories(data.Categories);

            foreach (var descriptor in descriptors) {
                ScatterChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (ScatterChartSeries)template.CloneNode(true) : new ScatterChartSeries();
                    InsertSeries(chart, seriesElement);
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
