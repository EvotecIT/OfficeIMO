using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;
using ChartIndex = DocumentFormat.OpenXml.Drawing.Charts.Index;

namespace OfficeIMO.Excel {
    internal static partial class ExcelChartUtils {
        internal static void UpdateChartData(ChartPart chartPart, ExcelChartData data, ExcelChartDataRange range) {
            if (chartPart == null) {
                throw new ArgumentNullException(nameof(chartPart));
            }
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            ChartSpace? chartSpace = chartPart.ChartSpace;
            Chart? chart = chartSpace?.GetFirstChild<Chart>();
            PlotArea? plotArea = chart?.GetFirstChild<PlotArea>();
            if (plotArea == null) {
                throw new InvalidOperationException("Chart plot area not found.");
            }

            int chartElementCount =
                plotArea.Elements<BarChart>().Count()
                + plotArea.Elements<Bar3DChart>().Count()
                + plotArea.Elements<LineChart>().Count()
                + plotArea.Elements<Line3DChart>().Count()
                + plotArea.Elements<AreaChart>().Count()
                + plotArea.Elements<Area3DChart>().Count()
                + plotArea.Elements<PieChart>().Count()
                + plotArea.Elements<Pie3DChart>().Count()
                + plotArea.Elements<OfPieChart>().Count()
                + plotArea.Elements<DoughnutChart>().Count()
                + plotArea.Elements<ScatterChart>().Count()
                + plotArea.Elements<BubbleChart>().Count()
                + plotArea.Elements<RadarChart>().Count()
                + plotArea.Elements<StockChart>().Count()
                + plotArea.Elements<Surface3DChart>().Count()
                + plotArea.Elements<SurfaceChart>().Count();

            ExcelChartType defaultType = InferChartType(plotArea);
            List<SeriesDescriptor> descriptors = BuildSeriesDescriptors(range, data, defaultType, useSeriesOverrides: chartElementCount > 1);
            ValidateSingleSeriesPieVariants(descriptors);

            if (plotArea.GetFirstChild<ScatterChart>() is ScatterChart scatterChart) {
                if (chartElementCount > 1) {
                    UpdateComboChartData(plotArea, data, range, descriptors);
                } else {
                    UpdateScatterChartSeries(scatterChart, data, range, descriptors);
                }
                return;
            }

            if (plotArea.GetFirstChild<BubbleChart>() != null) {
                throw new NotSupportedException("Updating bubble charts is not supported. Use range-based charts for bubble data.");
            }
            if (plotArea.GetFirstChild<StockChart>() is StockChart stockChart) {
                if (chartElementCount > 1) {
                    throw new NotSupportedException("Stock charts cannot be updated as combination charts.");
                }
                UpdateStockChartSeries(stockChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<Surface3DChart>() is Surface3DChart surface3DChart) {
                if (chartElementCount > 1) {
                    throw new NotSupportedException("Surface charts cannot be updated as combination charts.");
                }
                UpdateSurfaceChartSeries(surface3DChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<SurfaceChart>() is SurfaceChart surfaceChart) {
                if (chartElementCount > 1) {
                    throw new NotSupportedException("Surface charts cannot be updated as combination charts.");
                }
                UpdateSurfaceChartSeries(surfaceChart, data, range, descriptors);
                return;
            }

            if (chartElementCount > 1) {
                UpdateComboChartData(plotArea, data, range, descriptors);
                return;
            }

            if (plotArea.GetFirstChild<BarChart>() is BarChart barChart) {
                UpdateBarChartSeries(barChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<Bar3DChart>() is Bar3DChart bar3DChart) {
                UpdateBar3DChartSeries(bar3DChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<LineChart>() is LineChart lineChart) {
                UpdateLineChartSeries(lineChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<Line3DChart>() is Line3DChart line3DChart) {
                UpdateLine3DChartSeries(line3DChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<AreaChart>() is AreaChart areaChart) {
                UpdateAreaChartSeries(areaChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<Area3DChart>() is Area3DChart area3DChart) {
                UpdateArea3DChartSeries(area3DChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<RadarChart>() is RadarChart radarChart) {
                UpdateRadarChartSeries(radarChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<PieChart>() is PieChart pieChart) {
                UpdatePieChartSeries(pieChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<Pie3DChart>() is Pie3DChart pie3DChart) {
                UpdatePie3DChartSeries(pie3DChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<OfPieChart>() is OfPieChart ofPieChart) {
                UpdateOfPieChartSeries(ofPieChart, data, range, descriptors);
                return;
            }
            if (plotArea.GetFirstChild<DoughnutChart>() is DoughnutChart doughnutChart) {
                UpdateDoughnutChartSeries(doughnutChart, data, range, descriptors);
                return;
            }

            throw new NotSupportedException("Chart type is not supported for data updates.");
        }

        internal static ExcelChartDataRange? TryExtractDataRange(ChartPart chartPart) {
            var chart = chartPart.ChartSpace?.GetFirstChild<Chart>();
            var plotArea = chart?.GetFirstChild<PlotArea>();
            if (plotArea == null) return null;

            IReadOnlyList<OpenXmlCompositeElement> seriesList;
            if (plotArea.GetFirstChild<BarChart>() is BarChart bar) {
                seriesList = bar.Elements<BarChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<Bar3DChart>() is Bar3DChart bar3D) {
                seriesList = bar3D.Elements<BarChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<LineChart>() is LineChart line) {
                seriesList = line.Elements<LineChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<Line3DChart>() is Line3DChart line3D) {
                seriesList = line3D.Elements<LineChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<AreaChart>() is AreaChart area) {
                seriesList = area.Elements<AreaChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<Area3DChart>() is Area3DChart area3D) {
                seriesList = area3D.Elements<AreaChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<PieChart>() is PieChart pie) {
                seriesList = pie.Elements<PieChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<Pie3DChart>() is Pie3DChart pie3D) {
                seriesList = pie3D.Elements<PieChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<OfPieChart>() is OfPieChart ofPie) {
                seriesList = ofPie.Elements<PieChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<DoughnutChart>() is DoughnutChart doughnut) {
                seriesList = doughnut.Elements<PieChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<RadarChart>() is RadarChart radar) {
                seriesList = radar.Elements<RadarChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<StockChart>() is StockChart stock) {
                seriesList = stock.Elements<LineChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<Surface3DChart>() is Surface3DChart surface3D) {
                seriesList = surface3D.Elements<SurfaceChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<SurfaceChart>() is SurfaceChart surface) {
                seriesList = surface.Elements<SurfaceChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else if (plotArea.GetFirstChild<ScatterChart>() is ScatterChart scatter) {
                seriesList = scatter.Elements<ScatterChartSeries>().Cast<OpenXmlCompositeElement>().ToList();
            } else {
                return null;
            }

            if (seriesList.Count == 0) return null;

            var series = seriesList[0];
            string? catFormula;
            string? valFormula;
            if (series is ScatterChartSeries scatterSeries) {
                catFormula = scatterSeries.GetFirstChild<XValues>()?
                    .GetFirstChild<NumberReference>()?
                    .Formula?.Text;
                valFormula = scatterSeries.GetFirstChild<YValues>()?
                    .GetFirstChild<NumberReference>()?
                    .Formula?.Text;
            } else {
                catFormula = series.GetFirstChild<CategoryAxisData>()?
                    .GetFirstChild<StringReference>()?
                    .Formula?.Text;
                valFormula = series.GetFirstChild<Values>()?
                    .GetFirstChild<NumberReference>()?
                    .Formula?.Text;
            }

            if (!TryParseSheetQualifiedRange(catFormula, out var sheetName, out var catRange)) return null;
            if (!TryParseSheetQualifiedRange(valFormula, out var sheetNameValues, out var valRange)) return null;

            if (!string.Equals(sheetName, sheetNameValues, StringComparison.OrdinalIgnoreCase)) return null;
            if (!TryParseRange(catRange, out int r1, out int c1, out int r2, out _)) return null;
            if (!TryParseRange(valRange, out _, out _, out _, out _)) return null;

            int categoryCount = r2 - r1 + 1;
            if (categoryCount <= 0) return null;

            bool hasHeaderRow = false;
            int headerRow = r1 - 1;
            if (headerRow > 0) {
                foreach (var seriesElement in seriesList) {
                    string? nameFormula = seriesElement.GetFirstChild<SeriesText>()?
                        .GetFirstChild<StringReference>()?
                        .Formula?.Text;
                    if (!TryParseSheetQualifiedRange(nameFormula, out var nameSheet, out var nameRange)) {
                        continue;
                    }
                    if (!string.Equals(nameSheet, sheetName, StringComparison.OrdinalIgnoreCase)) {
                        continue;
                    }
                    if (!TryParseRange(nameRange, out int nameR1, out _, out int nameR2, out _)) {
                        continue;
                    }
                    if (nameR1 == nameR2 && nameR1 == headerRow) {
                        hasHeaderRow = true;
                        break;
                    }
                }
            }

            int startRow = hasHeaderRow ? headerRow : r1;
            int startColumn = c1;

            return new ExcelChartDataRange(sheetName, startRow, startColumn, categoryCount, seriesList.Count, hasHeaderRow);
        }

        internal static ExcelChartData? TryReadChartData(ExcelSheet sheet, ExcelChartDataRange range) {
            try {
                var categories = new List<string>(range.CategoryCount);
                for (int i = 0; i < range.CategoryCount; i++) {
                    int row = range.CategoryStartRow + i;
                    if (sheet.TryGetCellText(row, range.StartColumn, out var text)) {
                        categories.Add(text ?? string.Empty);
                    } else {
                        categories.Add(string.Empty);
                    }
                }

                var series = new List<ExcelChartSeries>(range.SeriesCount);
                for (int s = 0; s < range.SeriesCount; s++) {
                    int col = range.SeriesStartColumn + s;
                    string name = $"Series {s + 1}";
                    if (range.HasHeaderRow) {
                        if (sheet.TryGetCellText(range.StartRow, col, out var header) && !string.IsNullOrWhiteSpace(header)) {
                            name = header;
                        }
                    }

                    var values = new List<double>(range.CategoryCount);
                    for (int i = 0; i < range.CategoryCount; i++) {
                        int row = range.CategoryStartRow + i;
                        if (sheet.TryGetCellText(row, col, out var raw)
                            && double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out var val)) {
                            values.Add(val);
                        } else {
                            values.Add(0d);
                        }
                    }
                    series.Add(new ExcelChartSeries(name, values));
                }

                return new ExcelChartData(categories, series);
            } catch {
                return null;
            }
        }

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
            IReadOnlyList<double> xValues = ParseNumericCategories(data.Categories);

            foreach (var descriptor in descriptors) {
                ScatterChartSeries seriesElement;
                if (!existingByIndex.TryGetValue(descriptor.Index, out seriesElement!)) {
                    seriesElement = template != null ? (ScatterChartSeries)template.CloneNode(true) : new ScatterChartSeries();
                    InsertSeries(chart, seriesElement);
                }

                UpdateSeriesIndexOrder(seriesElement, descriptor.Index);
                string name = descriptor.Series?.Name ?? $"Series {descriptor.Index + 1}";
                UpdateSeriesText(seriesElement, range, descriptor.Index, name);
                UpdateXValues(seriesElement, range, xValues);
                UpdateYValues(seriesElement, range, descriptor.Index, data.Series[descriptor.Index].Values);
            }

            foreach (var seriesElement in existingSeries) {
                int idx = GetSeriesIndex(seriesElement);
                if (!indexSet.Contains(idx)) {
                    seriesElement.Remove();
                }
            }
        }

        private static void UpdateComboChartData(PlotArea plotArea, ExcelChartData data, ExcelChartDataRange range, IReadOnlyList<SeriesDescriptor> descriptors) {
            if (plotArea == null) {
                throw new ArgumentNullException(nameof(plotArea));
            }

            SeriesDescriptorSummary summary = SummarizeSeriesDescriptors(descriptors);
            if (summary.HasStock) {
                throw new NotSupportedException("Stock charts cannot be used in combination chart updates.");
            }
            if (summary.HasSurface) {
                throw new NotSupportedException("Surface charts cannot be used in combination chart updates.");
            }
            if (summary.HasLine3D) {
                throw new NotSupportedException("3-D line charts cannot be used in combination chart updates.");
            }
            if (summary.HasBar3D) {
                throw new NotSupportedException("3-D bar and column charts cannot be used in combination chart updates.");
            }
            if (summary.HasArea3D) {
                throw new NotSupportedException("3-D area charts cannot be used in combination chart updates.");
            }
            if (summary.HasBar && summary.HasNonBar) {
                throw new NotSupportedException("Cannot combine horizontal bar charts with other chart types.");
            }

            if (summary.HasPieOrDoughnut) {
                throw new NotSupportedException("Pie and doughnut charts cannot be combined with other chart types.");
            }

            bool isBarOrientation = summary.HasBar;
            AxisPositionValues primaryCategoryPosition = isBarOrientation ? AxisPositionValues.Left : AxisPositionValues.Bottom;
            AxisPositionValues primaryValuePosition = isBarOrientation ? AxisPositionValues.Bottom : AxisPositionValues.Left;
            AxisPositionValues secondaryCategoryPosition = isBarOrientation ? AxisPositionValues.Right : AxisPositionValues.Top;
            AxisPositionValues secondaryValuePosition = isBarOrientation ? AxisPositionValues.Top : AxisPositionValues.Right;

            var axisIds = EnsureAxisPairs(plotArea, summary.HasSecondary, primaryCategoryPosition, primaryValuePosition, secondaryCategoryPosition, secondaryValuePosition);

            var usedCharts = new List<OpenXmlCompositeElement>();
            foreach (SeriesDescriptorGroup group in GroupSeriesDescriptors(descriptors)) {
                uint categoryAxisId = group.AxisGroup == ExcelChartAxisGroup.Secondary ? axisIds.SecondaryCategoryId : axisIds.PrimaryCategoryId;
                uint valueAxisId = group.AxisGroup == ExcelChartAxisGroup.Secondary ? axisIds.SecondaryValueId : axisIds.PrimaryValueId;

                switch (group.ChartType) {
                    case ExcelChartType.ColumnClustered:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionValues.Column, BarGroupingValues.Clustered, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.ColumnStacked:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionValues.Column, BarGroupingValues.Stacked, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.ColumnStacked100:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionValues.Column, BarGroupingValues.PercentStacked, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.BarClustered:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionValues.Bar, BarGroupingValues.Clustered, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.BarStacked:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionValues.Bar, BarGroupingValues.Stacked, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.BarStacked100:
                        usedCharts.Add(UpdateOrCreateBarChart(plotArea, group.Descriptors, data, range, BarDirectionValues.Bar, BarGroupingValues.PercentStacked, categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Line:
                    case ExcelChartType.LineStacked:
                    case ExcelChartType.LineStacked100:
                        usedCharts.Add(UpdateOrCreateLineChart(plotArea, group.Descriptors, data, range, GetLineGrouping(group.ChartType), categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Area:
                    case ExcelChartType.AreaStacked:
                    case ExcelChartType.AreaStacked100:
                        usedCharts.Add(UpdateOrCreateAreaChart(plotArea, group.Descriptors, data, range, GetAreaGrouping(group.ChartType), categoryAxisId, valueAxisId));
                        break;
                    case ExcelChartType.Scatter:
                        usedCharts.Add(UpdateOrCreateScatterChart(plotArea, group.Descriptors, data, range, categoryAxisId, valueAxisId));
                        break;
                    default:
                        throw new NotSupportedException($"Chart type {group.ChartType} is not supported in combination charts.");
                }
            }

            foreach (var chart in plotArea.Elements<BarChart>().Cast<OpenXmlCompositeElement>().ToList()) {
                if (!usedCharts.Contains(chart)) chart.Remove();
            }
            foreach (var chart in plotArea.Elements<LineChart>().Cast<OpenXmlCompositeElement>().ToList()) {
                if (!usedCharts.Contains(chart)) chart.Remove();
            }
            foreach (var chart in plotArea.Elements<AreaChart>().Cast<OpenXmlCompositeElement>().ToList()) {
                if (!usedCharts.Contains(chart)) chart.Remove();
            }
            foreach (var chart in plotArea.Elements<ScatterChart>().Cast<OpenXmlCompositeElement>().ToList()) {
                if (!usedCharts.Contains(chart)) chart.Remove();
            }
        }

        private static AxisIdSet EnsureAxisPairs(PlotArea plotArea, bool includeSecondary, AxisPositionValues primaryCategoryPosition,
            AxisPositionValues primaryValuePosition, AxisPositionValues secondaryCategoryPosition, AxisPositionValues secondaryValuePosition) {
            var categoryAxes = plotArea.Elements<CategoryAxis>().ToList();
            var valueAxes = plotArea.Elements<ValueAxis>().ToList();

            CategoryAxis? primaryCategory = categoryAxes.FirstOrDefault(ax => ax.AxisPosition?.Val?.Value == primaryCategoryPosition)
                ?? categoryAxes.FirstOrDefault();
            ValueAxis? primaryValue = valueAxes.FirstOrDefault(ax => ax.AxisPosition?.Val?.Value == primaryValuePosition)
                ?? valueAxes.FirstOrDefault();

            uint primaryCategoryId = primaryCategory?.AxisId?.Val?.Value ?? ExcelChartAxisIdGenerator.GetNextId();
            uint primaryValueId = primaryValue?.AxisId?.Val?.Value ?? ExcelChartAxisIdGenerator.GetNextId();

            if (primaryCategory == null) {
                primaryCategory = CreateCategoryAxis(primaryCategoryId, primaryValueId, primaryCategoryPosition);
                plotArea.Append(primaryCategory);
            } else {
                EnsureAxisId(primaryCategory, primaryCategoryId);
                EnsureAxisPosition(primaryCategory, primaryCategoryPosition);
                EnsureCrossingAxis(primaryCategory, primaryValueId);
            }

            if (primaryValue == null) {
                primaryValue = CreateValueAxis(primaryValueId, primaryCategoryId, primaryValuePosition);
                plotArea.Append(primaryValue);
            } else {
                EnsureAxisId(primaryValue, primaryValueId);
                EnsureAxisPosition(primaryValue, primaryValuePosition);
                EnsureCrossingAxis(primaryValue, primaryCategoryId);
            }

            uint secondaryCategoryId = 0;
            uint secondaryValueId = 0;
            if (includeSecondary) {
                CategoryAxis? secondaryCategory = categoryAxes.FirstOrDefault(ax => ax.AxisPosition?.Val?.Value == secondaryCategoryPosition);
                ValueAxis? secondaryValue = valueAxes.FirstOrDefault(ax => ax.AxisPosition?.Val?.Value == secondaryValuePosition);

                secondaryCategoryId = secondaryCategory?.AxisId?.Val?.Value ?? ExcelChartAxisIdGenerator.GetNextId();
                secondaryValueId = secondaryValue?.AxisId?.Val?.Value ?? ExcelChartAxisIdGenerator.GetNextId();

                if (secondaryCategory == null) {
                    secondaryCategory = CreateCategoryAxis(secondaryCategoryId, secondaryValueId, secondaryCategoryPosition);
                    plotArea.Append(secondaryCategory);
                } else {
                    EnsureAxisId(secondaryCategory, secondaryCategoryId);
                    EnsureAxisPosition(secondaryCategory, secondaryCategoryPosition);
                    EnsureCrossingAxis(secondaryCategory, secondaryValueId);
                }

                if (secondaryValue == null) {
                    secondaryValue = CreateValueAxis(secondaryValueId, secondaryCategoryId, secondaryValuePosition);
                    plotArea.Append(secondaryValue);
                } else {
                    EnsureAxisId(secondaryValue, secondaryValueId);
                    EnsureAxisPosition(secondaryValue, secondaryValuePosition);
                    EnsureCrossingAxis(secondaryValue, secondaryCategoryId);
                }
            }

            return new AxisIdSet(primaryCategoryId, primaryValueId, secondaryCategoryId, secondaryValueId);
        }

        private static void EnsureAxisId(OpenXmlCompositeElement axis, uint axisId) {
            AxisId id = axis.GetFirstChild<AxisId>() ?? new AxisId();
            id.Val = axisId;
            if (id.Parent == null) {
                axis.PrependChild(id);
            }
        }

        private static void EnsureAxisPosition(OpenXmlCompositeElement axis, AxisPositionValues position) {
            AxisPosition axisPosition = axis.GetFirstChild<AxisPosition>() ?? new AxisPosition();
            axisPosition.Val = position;
            if (axisPosition.Parent == null) {
                axis.Append(axisPosition);
            }
        }

        private static void EnsureCrossingAxis(OpenXmlCompositeElement axis, uint crossAxisId) {
            CrossingAxis crossingAxis = axis.GetFirstChild<CrossingAxis>() ?? new CrossingAxis();
            crossingAxis.Val = crossAxisId;
            if (crossingAxis.Parent == null) {
                axis.Append(crossingAxis);
            }
        }

        private static BarChart UpdateOrCreateBarChart(PlotArea plotArea, IReadOnlyList<SeriesDescriptor> descriptors, ExcelChartData data, ExcelChartDataRange range,
            BarDirectionValues direction, BarGroupingValues grouping, uint categoryAxisId, uint valueAxisId) {
            BarChart? chart = plotArea.Elements<BarChart>()
                .FirstOrDefault(c => (c.GetFirstChild<BarDirection>()?.Val ?? BarDirectionValues.Column) == direction
                                     && (c.GetFirstChild<BarGrouping>()?.Val ?? BarGroupingValues.Clustered) == grouping
                                     && ChartHasAxisIds(c, categoryAxisId, valueAxisId));

            chart ??= plotArea.Elements<BarChart>()
                .FirstOrDefault(c => (c.GetFirstChild<BarDirection>()?.Val ?? BarDirectionValues.Column) == direction
                                     && (c.GetFirstChild<BarGrouping>()?.Val ?? BarGroupingValues.Clustered) == grouping);

            if (chart == null) {
                chart = CreateBarChart(range, descriptors, direction, grouping, categoryAxisId, valueAxisId);
                plotArea.Append(chart);
            } else {
                ResetAxisIds(chart, categoryAxisId, valueAxisId);
                UpdateBarChartSeries(chart, data, range, descriptors);
            }

            return chart;
        }

        private static LineChart UpdateOrCreateLineChart(PlotArea plotArea, IReadOnlyList<SeriesDescriptor> descriptors, ExcelChartData data, ExcelChartDataRange range,
            GroupingValues grouping, uint categoryAxisId, uint valueAxisId) {
            LineChart? chart = plotArea.Elements<LineChart>()
                .FirstOrDefault(c => (c.GetFirstChild<Grouping>()?.Val ?? GroupingValues.Standard) == grouping
                                     && ChartHasAxisIds(c, categoryAxisId, valueAxisId));

            chart ??= plotArea.Elements<LineChart>()
                .FirstOrDefault(c => (c.GetFirstChild<Grouping>()?.Val ?? GroupingValues.Standard) == grouping);

            if (chart == null) {
                chart = CreateLineChart(range, descriptors, grouping, categoryAxisId, valueAxisId);
                plotArea.Append(chart);
            } else {
                EnsureGrouping(chart, grouping);
                ResetAxisIds(chart, categoryAxisId, valueAxisId);
                UpdateLineChartSeries(chart, data, range, descriptors);
            }

            return chart;
        }

        private static AreaChart UpdateOrCreateAreaChart(PlotArea plotArea, IReadOnlyList<SeriesDescriptor> descriptors, ExcelChartData data, ExcelChartDataRange range,
            GroupingValues grouping, uint categoryAxisId, uint valueAxisId) {
            AreaChart? chart = plotArea.Elements<AreaChart>()
                .FirstOrDefault(c => (c.GetFirstChild<Grouping>()?.Val ?? GroupingValues.Standard) == grouping
                                     && ChartHasAxisIds(c, categoryAxisId, valueAxisId));

            chart ??= plotArea.Elements<AreaChart>()
                .FirstOrDefault(c => (c.GetFirstChild<Grouping>()?.Val ?? GroupingValues.Standard) == grouping);

            if (chart == null) {
                chart = CreateAreaChart(range, descriptors, grouping, categoryAxisId, valueAxisId);
                plotArea.Append(chart);
            } else {
                EnsureGrouping(chart, grouping);
                ResetAxisIds(chart, categoryAxisId, valueAxisId);
                UpdateAreaChartSeries(chart, data, range, descriptors);
            }

            return chart;
        }

        private static void EnsureGrouping(OpenXmlCompositeElement chart, GroupingValues groupingValue) {
            Grouping grouping = chart.GetFirstChild<Grouping>() ?? new Grouping();
            grouping.Val = groupingValue;
            if (grouping.Parent == null) {
                chart.PrependChild(grouping);
            }
        }

        private static ScatterChart UpdateOrCreateScatterChart(PlotArea plotArea, IReadOnlyList<SeriesDescriptor> descriptors, ExcelChartData data, ExcelChartDataRange range,
            uint xAxisId, uint yAxisId) {
            ScatterChart? chart = plotArea.Elements<ScatterChart>()
                .FirstOrDefault(c => ChartHasAxisIds(c, xAxisId, yAxisId))
                ?? plotArea.Elements<ScatterChart>().FirstOrDefault();

            if (chart == null) {
                chart = CreateScatterChart(range, descriptors, xAxisId, yAxisId, data);
                plotArea.Append(chart);
            } else {
                ResetAxisIds(chart, xAxisId, yAxisId);
                UpdateScatterChartSeries(chart, data, range, descriptors);
            }

            return chart;
        }

        private static bool ChartHasAxisIds(OpenXmlCompositeElement chart, uint categoryAxisId, uint valueAxisId) {
            bool hasCategoryAxis = false;
            bool hasValueAxis = false;
            foreach (AxisId id in chart.Elements<AxisId>()) {
                uint? value = id.Val?.Value;
                if (value == categoryAxisId) {
                    hasCategoryAxis = true;
                    if (hasValueAxis) {
                        return true;
                    }
                } else if (value == valueAxisId) {
                    hasValueAxis = true;
                    if (hasCategoryAxis) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static void ResetAxisIds(OpenXmlCompositeElement chart, uint categoryAxisId, uint valueAxisId) {
            chart.RemoveAllChildren<AxisId>();
            chart.Append(new AxisId { Val = categoryAxisId });
            chart.Append(new AxisId { Val = valueAxisId });
        }

        private static void UpdateSeriesIndexOrder(OpenXmlCompositeElement series, int index) {
            ChartIndex idx = series.GetFirstChild<ChartIndex>() ?? new ChartIndex();
            idx.Val = (uint)index;
            if (idx.Parent == null) {
                series.PrependChild(idx);
            }

            Order order = series.GetFirstChild<Order>() ?? new Order();
            order.Val = (uint)index;
            if (order.Parent == null) {
                series.InsertAfter(order, idx);
            }
        }

        private static int GetSeriesIndex(OpenXmlCompositeElement series) {
            return (int)(series.GetFirstChild<ChartIndex>()?.Val?.Value ?? 0U);
        }

        private static HashSet<int> CreateDescriptorIndexSet(IReadOnlyList<SeriesDescriptor> descriptors) {
            var indexes = new HashSet<int>();
            for (int i = 0; i < descriptors.Count; i++) {
                indexes.Add(descriptors[i].Index);
            }

            return indexes;
        }

        private static Dictionary<int, TSeries> CreateSeriesIndexMap<TSeries>(IReadOnlyList<TSeries> series)
            where TSeries : OpenXmlCompositeElement {
            var map = new Dictionary<int, TSeries>(series.Count);
            for (int i = 0; i < series.Count; i++) {
                map.Add(GetSeriesIndex(series[i]), series[i]);
            }

            return map;
        }

        private static void UpdateSeriesText(OpenXmlCompositeElement series, ExcelChartDataRange range, int seriesIndex, string seriesName) {
            SeriesText seriesText = series.GetFirstChild<SeriesText>() ?? new SeriesText();
            seriesText.RemoveAllChildren<StringReference>();
            seriesText.RemoveAllChildren<StringLiteral>();
            seriesText.RemoveAllChildren<NumericValue>();

            if (range.HasHeaderRow) {
                string seriesCell = range.SeriesNameCellA1(seriesIndex);
                string formula = BuildSheetQualifiedRange(range.SheetName, seriesCell);
                seriesText.Append(CreateSingleStringReference(formula, seriesName));
            } else {
                seriesText.Append(new NumericValue { Text = seriesName });
            }

            if (seriesText.Parent == null) {
                OpenXmlElement? insertAfter = series.GetFirstChild<Order>();
                insertAfter ??= series.GetFirstChild<ChartIndex>();
                if (insertAfter != null) {
                    series.InsertAfter(seriesText, insertAfter);
                } else {
                    series.PrependChild(seriesText);
                }
            }
        }

        private static void UpdateCategoryAxisData(OpenXmlCompositeElement series, ExcelChartDataRange range, IReadOnlyList<string> categories) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.CategoriesRangeA1);
            CategoryAxisData categoryAxisData = series.GetFirstChild<CategoryAxisData>() ?? new CategoryAxisData();
            categoryAxisData.RemoveAllChildren<StringReference>();
            categoryAxisData.RemoveAllChildren<StringLiteral>();
            categoryAxisData.Append(CreateStringReference(formula, categories));

            if (categoryAxisData.Parent == null) {
                series.Append(categoryAxisData);
            }
        }

        private static void UpdateValues(OpenXmlCompositeElement series, ExcelChartDataRange range, int seriesIndex, IReadOnlyList<double> values) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.SeriesValuesRangeA1(seriesIndex));
            Values valueElement = series.GetFirstChild<Values>() ?? new Values();
            valueElement.RemoveAllChildren<NumberReference>();
            valueElement.RemoveAllChildren<NumberLiteral>();
            valueElement.Append(CreateNumberReference(formula, values));

            if (valueElement.Parent == null) {
                series.Append(valueElement);
            }
        }

        private static void UpdateXValues(ScatterChartSeries series, ExcelChartDataRange range, IReadOnlyList<double> xValues) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.CategoriesRangeA1);
            XValues xValueElement = series.GetFirstChild<XValues>() ?? new XValues();
            xValueElement.RemoveAllChildren<NumberReference>();
            xValueElement.RemoveAllChildren<NumberLiteral>();
            xValueElement.Append(CreateNumberReference(formula, xValues));

            if (xValueElement.Parent == null) {
                series.Append(xValueElement);
            }
        }

        private static void UpdateYValues(ScatterChartSeries series, ExcelChartDataRange range, int seriesIndex, IReadOnlyList<double> values) {
            string formula = BuildSheetQualifiedRange(range.SheetName, range.SeriesValuesRangeA1(seriesIndex));
            YValues yValueElement = series.GetFirstChild<YValues>() ?? new YValues();
            yValueElement.RemoveAllChildren<NumberReference>();
            yValueElement.RemoveAllChildren<NumberLiteral>();
            yValueElement.Append(CreateNumberReference(formula, values));

            if (yValueElement.Parent == null) {
                series.Append(yValueElement);
            }
        }

        private static void InsertSeries(OpenXmlCompositeElement chart, OpenXmlElement series) {
            OpenXmlElement? insertBefore = chart.ChildElements.FirstOrDefault(child =>
                child is DataLabels ||
                child is GapWidth ||
                child is GapDepth ||
                child is Overlap ||
                child is HighLowLines ||
                child is UpDownBars ||
                child is BandFormats ||
                child is Shape ||
                child is AxisId ||
                child is Marker ||
                child is Smooth ||
                child is SeriesLines);

            if (insertBefore != null) {
                chart.InsertBefore(series, insertBefore);
            } else {
                chart.Append(series);
            }
        }

        private static void EnsureStockChartLines(StockChart stockChart, int seriesCount) {
            if (stockChart.GetFirstChild<HighLowLines>() == null) {
                OpenXmlElement? insertBefore = stockChart.GetFirstChild<UpDownBars>();
                insertBefore ??= stockChart.GetFirstChild<AxisId>();
                if (insertBefore != null) {
                    stockChart.InsertBefore(new HighLowLines(), insertBefore);
                } else {
                    stockChart.Append(new HighLowLines());
                }
            }

            UpDownBars? bars = stockChart.GetFirstChild<UpDownBars>();
            if (seriesCount == 4) {
                if (bars == null) {
                    bars = new UpDownBars(
                        new GapWidth { Val = (UInt16Value)150U },
                        new UpBars(),
                        new DownBars());
                    OpenXmlElement? insertBefore = stockChart.GetFirstChild<AxisId>();
                    if (insertBefore != null) {
                        stockChart.InsertBefore(bars, insertBefore);
                    } else {
                        stockChart.Append(bars);
                    }
                }
            } else {
                bars?.Remove();
            }
        }
    }
}
