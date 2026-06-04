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
        internal static ExcelChartDataRange? TryExtractDataRange(ChartPart chartPart) {
            var chart = chartPart.ChartSpace?.GetFirstChild<Chart>();
            var plotArea = chart?.GetFirstChild<PlotArea>();
            if (plotArea == null) return null;

            IReadOnlyList<OpenXmlCompositeElement> seriesList = GetChartSeries(plotArea);

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

        private static IReadOnlyList<OpenXmlCompositeElement> GetChartSeries(PlotArea plotArea) {
            var series = new List<OpenXmlCompositeElement>();
            foreach (OpenXmlElement chartElement in plotArea.ChildElements) {
                switch (chartElement) {
                    case BarChart bar:
                        series.AddRange(bar.Elements<BarChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case Bar3DChart bar3D:
                        series.AddRange(bar3D.Elements<BarChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case LineChart line:
                        series.AddRange(line.Elements<LineChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case Line3DChart line3D:
                        series.AddRange(line3D.Elements<LineChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case AreaChart area:
                        series.AddRange(area.Elements<AreaChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case Area3DChart area3D:
                        series.AddRange(area3D.Elements<AreaChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case PieChart pie:
                        series.AddRange(pie.Elements<PieChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case Pie3DChart pie3D:
                        series.AddRange(pie3D.Elements<PieChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case OfPieChart ofPie:
                        series.AddRange(ofPie.Elements<PieChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case DoughnutChart doughnut:
                        series.AddRange(doughnut.Elements<PieChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case RadarChart radar:
                        series.AddRange(radar.Elements<RadarChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case StockChart stock:
                        series.AddRange(stock.Elements<LineChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case Surface3DChart surface3D:
                        series.AddRange(surface3D.Elements<SurfaceChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case SurfaceChart surface:
                        series.AddRange(surface.Elements<SurfaceChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                    case ScatterChart scatter:
                        series.AddRange(scatter.Elements<ScatterChartSeries>().Cast<OpenXmlCompositeElement>());
                        break;
                }
            }

            return series
                .OrderBy(item => item.GetFirstChild<ChartIndex>()?.Val?.Value ?? uint.MaxValue)
                .ToList();
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

        internal static ExcelChartData ApplyChartSeriesTypes(ChartPart chartPart, ExcelChartData data, ExcelChartType defaultType) {
            var chart = chartPart.ChartSpace?.GetFirstChild<Chart>();
            var plotArea = chart?.GetFirstChild<PlotArea>();
            if (plotArea == null || data.Series.Count == 0) {
                return data;
            }

            var seriesTypes = new Dictionary<int, ExcelChartType>();
            int chartElementCount = 0;
            foreach (OpenXmlElement chartElement in plotArea.ChildElements) {
                if (!TryGetChartElementType(chartElement, out ExcelChartType chartType)) {
                    continue;
                }

                chartElementCount++;
                foreach (OpenXmlElement child in chartElement.ChildElements) {
                    ChartIndex? index = (child as OpenXmlCompositeElement)?.GetFirstChild<ChartIndex>();
                    if (index?.Val?.Value != null) {
                        seriesTypes[(int)index.Val.Value] = chartType;
                    }
                }
            }

            if (chartElementCount <= 1 || seriesTypes.Count == 0) {
                return data;
            }

            var series = new List<ExcelChartSeries>(data.Series.Count);
            for (int i = 0; i < data.Series.Count; i++) {
                ExcelChartSeries current = data.Series[i];
                ExcelChartType? seriesType = current.ChartType;
                if (seriesTypes.TryGetValue(i, out ExcelChartType chartType) && chartType != defaultType) {
                    seriesType = chartType;
                }

                series.Add(new ExcelChartSeries(current.Name, current.Values, seriesType, current.AxisGroup));
            }

            return new ExcelChartData(data.Categories, series);
        }

        private static bool TryGetChartElementType(OpenXmlElement element, out ExcelChartType chartType) {
            switch (element) {
                case BarChart bar:
                    BarDirectionValues direction = bar.GetFirstChild<BarDirection>()?.Val?.Value ?? BarDirectionValues.Column;
                    BarGroupingValues barGrouping = bar.GetFirstChild<BarGrouping>()?.Val?.Value ?? BarGroupingValues.Clustered;
                    if (direction == BarDirectionValues.Bar) {
                        chartType = barGrouping == BarGroupingValues.PercentStacked
                            ? ExcelChartType.BarStacked100
                            : barGrouping == BarGroupingValues.Stacked ? ExcelChartType.BarStacked : ExcelChartType.BarClustered;
                    } else {
                        chartType = barGrouping == BarGroupingValues.PercentStacked
                            ? ExcelChartType.ColumnStacked100
                            : barGrouping == BarGroupingValues.Stacked ? ExcelChartType.ColumnStacked : ExcelChartType.ColumnClustered;
                    }

                    return true;
                case LineChart line:
                    GroupingValues lineGrouping = line.GetFirstChild<Grouping>()?.Val?.Value ?? GroupingValues.Standard;
                    chartType = lineGrouping == GroupingValues.PercentStacked
                        ? ExcelChartType.LineStacked100
                        : lineGrouping == GroupingValues.Stacked ? ExcelChartType.LineStacked : ExcelChartType.Line;
                    return true;
                case AreaChart area:
                    GroupingValues areaGrouping = area.GetFirstChild<Grouping>()?.Val?.Value ?? GroupingValues.Standard;
                    chartType = areaGrouping == GroupingValues.PercentStacked
                        ? ExcelChartType.AreaStacked100
                        : areaGrouping == GroupingValues.Stacked ? ExcelChartType.AreaStacked : ExcelChartType.Area;
                    return true;
                case ScatterChart:
                    chartType = ExcelChartType.Scatter;
                    return true;
                case RadarChart:
                    chartType = ExcelChartType.Radar;
                    return true;
                case PieChart:
                    chartType = ExcelChartType.Pie;
                    return true;
                case Pie3DChart:
                    chartType = ExcelChartType.Pie3D;
                    return true;
                case DoughnutChart:
                    chartType = ExcelChartType.Doughnut;
                    return true;
                default:
                    chartType = default;
                    return false;
            }
        }
    }
}
