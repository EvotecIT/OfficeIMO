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
    }
}
