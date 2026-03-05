using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.PowerPoint {
    internal enum PowerPointChartKind {
        ClusteredColumn,
        Line,
        Scatter,
        Pie,
        Doughnut
    }

    internal static partial class PowerPointUtils {
        private static readonly Lazy<byte[]> ChartStyle251Bytes =
            new(() => LoadEmbeddedResource("OfficeIMO.PowerPoint.Resources.chart-style-251.xml"));

        private static readonly Lazy<byte[]> ChartColorStyle10Bytes =
            new(() => LoadEmbeddedResource("OfficeIMO.PowerPoint.Resources.chart-colors-10.xml"));

        private const string ChartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        private const string DrawingNamespace = "http://schemas.openxmlformats.org/drawingml/2006/main";
        private const string RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

        internal static void PopulateChartStyle(ChartStylePart stylePart) {
            if (stylePart == null) {
                throw new ArgumentNullException(nameof(stylePart));
            }

            using var stream = new MemoryStream(ChartStyle251Bytes.Value);
            stylePart.FeedData(stream);
        }

        internal static void PopulateChartColorStyle(ChartColorStylePart colorStylePart) {
            if (colorStylePart == null) {
                throw new ArgumentNullException(nameof(colorStylePart));
            }

            using var stream = new MemoryStream(ChartColorStyle10Bytes.Value);
            colorStylePart.FeedData(stream);
        }

        internal static void PopulateChart(ChartPart chartPart, string embeddedRelId, PowerPointChartData data,
            PowerPointChartKind chartKind = PowerPointChartKind.ClusteredColumn) {
            if (chartPart == null) {
                throw new ArgumentNullException(nameof(chartPart));
            }
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            C.ChartSpace chartSpace = new();
            chartSpace.AddNamespaceDeclaration("c", ChartNamespace);
            chartSpace.AddNamespaceDeclaration("a", DrawingNamespace);
            chartSpace.AddNamespaceDeclaration("r", RelationshipNamespace);

            chartSpace.Append(new C.Date1904 { Val = false });
            chartSpace.Append(new C.EditingLanguage { Val = "en-US" });
            chartSpace.Append(new C.RoundedCorners { Val = false });

            C.Chart chart = new();
            chart.Append(new C.AutoTitleDeleted { Val = false });

            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());

            AppendChartContent(plotArea, data, chartKind);

            chart.Append(plotArea);
            chart.Append(new C.Legend(
                new C.LegendPosition { Val = C.LegendPositionValues.Bottom },
                new C.Layout(),
                new C.Overlay { Val = false }));
            chart.Append(new C.PlotVisibleOnly { Val = true });
            chart.Append(new C.DisplayBlanksAs { Val = C.DisplayBlanksAsValues.Gap });
            chart.Append(new C.ShowDataLabelsOverMaximum { Val = false });

            chartSpace.Append(chart);

            if (!string.IsNullOrWhiteSpace(embeddedRelId)) {
                chartSpace.Append(new C.ExternalData {
                    Id = embeddedRelId,
                    AutoUpdate = new C.AutoUpdate { Val = false }
                });
            }

            chartPart.ChartSpace = chartSpace;
        }

        internal static void PopulateChart(ChartPart chartPart, string embeddedRelId, PowerPointScatterChartData data,
            PowerPointChartKind chartKind = PowerPointChartKind.Scatter) {
            if (chartPart == null) {
                throw new ArgumentNullException(nameof(chartPart));
            }
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            C.ChartSpace chartSpace = new();
            chartSpace.AddNamespaceDeclaration("c", ChartNamespace);
            chartSpace.AddNamespaceDeclaration("a", DrawingNamespace);
            chartSpace.AddNamespaceDeclaration("r", RelationshipNamespace);

            chartSpace.Append(new C.Date1904 { Val = false });
            chartSpace.Append(new C.EditingLanguage { Val = "en-US" });
            chartSpace.Append(new C.RoundedCorners { Val = false });

            C.Chart chart = new();
            chart.Append(new C.AutoTitleDeleted { Val = false });

            C.PlotArea plotArea = new();
            plotArea.Append(new C.Layout());

            AppendChartContent(plotArea, data, chartKind);

            chart.Append(plotArea);
            chart.Append(new C.Legend(
                new C.LegendPosition { Val = C.LegendPositionValues.Bottom },
                new C.Layout(),
                new C.Overlay { Val = false }));
            chart.Append(new C.PlotVisibleOnly { Val = true });
            chart.Append(new C.DisplayBlanksAs { Val = C.DisplayBlanksAsValues.Gap });
            chart.Append(new C.ShowDataLabelsOverMaximum { Val = false });

            chartSpace.Append(chart);

            if (!string.IsNullOrWhiteSpace(embeddedRelId)) {
                chartSpace.Append(new C.ExternalData {
                    Id = embeddedRelId,
                    AutoUpdate = new C.AutoUpdate { Val = false }
                });
            }

            chartPart.ChartSpace = chartSpace;
        }

        private static void AppendChartContent(C.PlotArea plotArea, PowerPointChartData data, PowerPointChartKind chartKind) {
            switch (chartKind) {
                case PowerPointChartKind.ClusteredColumn:
                    uint categoryAxisId;
                    uint valueAxisId;
                    C.BarChart barChart = CreateBarChart(data, out categoryAxisId, out valueAxisId);
                    plotArea.Append(barChart);
                    plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
                    plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));
                    return;
                case PowerPointChartKind.Line:
                    uint lineCategoryAxisId;
                    uint lineValueAxisId;
                    C.LineChart lineChart = CreateLineChart(data, out lineCategoryAxisId, out lineValueAxisId);
                    plotArea.Append(lineChart);
                    plotArea.Append(CreateCategoryAxis(lineCategoryAxisId, lineValueAxisId));
                    plotArea.Append(CreateValueAxis(lineValueAxisId, lineCategoryAxisId));
                    return;
                case PowerPointChartKind.Pie:
                    plotArea.Append(CreatePieChart(data));
                    return;
                case PowerPointChartKind.Doughnut:
                    plotArea.Append(CreateDoughnutChart(data));
                    return;
                default:
                    throw new NotSupportedException($"Chart kind {chartKind} is not supported.");
            }
        }

        private static void AppendChartContent(C.PlotArea plotArea, PowerPointScatterChartData data, PowerPointChartKind chartKind) {
            switch (chartKind) {
                case PowerPointChartKind.Scatter:
                    uint xAxisId;
                    uint yAxisId;
                    C.ScatterChart scatterChart = CreateScatterChart(data, out xAxisId, out yAxisId);
                    plotArea.Append(scatterChart);
                    plotArea.Append(CreateValueAxis(xAxisId, yAxisId, C.AxisPositionValues.Bottom));
                    plotArea.Append(CreateValueAxis(yAxisId, xAxisId, C.AxisPositionValues.Left));
                    return;
                default:
                    throw new NotSupportedException($"Chart kind {chartKind} is not supported for scatter data.");
            }
        }

        internal static void UpdateChartData(ChartPart chartPart, PowerPointChartData data) {
            if (chartPart == null) {
                throw new ArgumentNullException(nameof(chartPart));
            }
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            C.ChartSpace? chartSpace = chartPart.ChartSpace;
            C.Chart? chart = chartSpace?.GetFirstChild<C.Chart>();
            C.PlotArea? plotArea = chart?.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                throw new InvalidOperationException("Chart plot area not found.");
            }

            if (plotArea.GetFirstChild<C.BarChart>() is C.BarChart barChart) {
                UpdateBarChartSeries(barChart, data);
                return;
            }

            if (plotArea.GetFirstChild<C.LineChart>() is C.LineChart lineChart) {
                UpdateLineChartSeries(lineChart, data);
                return;
            }

            if (plotArea.GetFirstChild<C.AreaChart>() is C.AreaChart areaChart) {
                UpdateAreaChartSeries(areaChart, data);
                return;
            }

            throw new NotSupportedException("Chart type is not supported for data updates.");
        }

        internal static void UpdateChartData(ChartPart chartPart, PowerPointScatterChartData data) {
            if (chartPart == null) {
                throw new ArgumentNullException(nameof(chartPart));
            }
            if (data == null) {
                throw new ArgumentNullException(nameof(data));
            }

            C.ChartSpace? chartSpace = chartPart.ChartSpace;
            C.Chart? chart = chartSpace?.GetFirstChild<C.Chart>();
            C.PlotArea? plotArea = chart?.GetFirstChild<C.PlotArea>();
            if (plotArea == null) {
                throw new InvalidOperationException("Chart plot area not found.");
            }

            if (plotArea.GetFirstChild<C.ScatterChart>() is C.ScatterChart scatterChart) {
                UpdateScatterChartSeries(scatterChart, data);
                return;
            }

            throw new NotSupportedException("Chart type is not supported for scatter data updates.");
        }

        internal static byte[] BuildChartWorkbook(PowerPointChartData data) {
            using MemoryStream ms = new();
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook)) {
                WorkbookPart wbPart = doc.AddWorkbookPart();
                wbPart.Workbook = new S.Workbook();

                WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
                var sheetData = new S.SheetData();

                int seriesCount = data.Series.Count;
                int categoryCount = data.Categories.Count;
                int totalColumns = seriesCount + 1;
                int totalRows = categoryCount + 1;
                string lastColumn = ColumnLetter(totalColumns);
                string dimensionRef = $"A1:{lastColumn}{totalRows}";

                wsPart.Worksheet = new S.Worksheet(
                    new S.SheetDimension { Reference = dimensionRef },
                    sheetData);

                var sharedStringsPart = wbPart.AddNewPart<SharedStringTablePart>();
                sharedStringsPart.SharedStringTable = new S.SharedStringTable();

                var stringIndex = new Dictionary<string, int>(StringComparer.Ordinal);
                int GetStringIndex(string value) {
                    if (!stringIndex.TryGetValue(value, out int idx)) {
                        idx = stringIndex.Count;
                        stringIndex[value] = idx;
                        sharedStringsPart.SharedStringTable.AppendChild(new S.SharedStringItem(new S.Text(value)));
                    }
                    return idx;
                }

                var headerRow = new S.Row { RowIndex = 1U, Spans = new ListValue<StringValue> { InnerText = $"1:{totalColumns}" } };
                headerRow.Append(CreateSharedStringCell("A1", GetStringIndex(" ")));
                for (int i = 0; i < seriesCount; i++) {
                    string cellRef = $"{ColumnLetter(i + 2)}1";
                    headerRow.Append(CreateSharedStringCell(cellRef, GetStringIndex(data.Series[i].Name)));
                }
                sheetData.Append(headerRow);

                for (int rowIndex = 0; rowIndex < categoryCount; rowIndex++) {
                    uint excelRow = (uint)(rowIndex + 2);
                    var row = new S.Row { RowIndex = excelRow, Spans = new ListValue<StringValue> { InnerText = $"1:{totalColumns}" } };
                    string category = data.Categories[rowIndex] ?? string.Empty;
                    row.Append(CreateSharedStringCell($"A{excelRow}", GetStringIndex(category)));

                    for (int seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
                        string cellRef = $"{ColumnLetter(seriesIndex + 2)}{excelRow}";
                        double value = data.Series[seriesIndex].Values[rowIndex];
                        row.Append(CreateNumberCell(cellRef, value));
                    }
                    sheetData.Append(row);
                }

                var sheets = wbPart.Workbook.AppendChild(new S.Sheets());
                sheets.Append(new S.Sheet {
                    Id = wbPart.GetIdOfPart(wsPart),
                    SheetId = 1U,
                    Name = "Sheet1"
                });

                sharedStringsPart.SharedStringTable.Save();
                wsPart.Worksheet.Save();
                wbPart.Workbook.Save();
            }

            return ms.ToArray();
        }

        internal static byte[] BuildChartWorkbook(PowerPointScatterChartData data) {
            using MemoryStream ms = new();
            using (SpreadsheetDocument doc = SpreadsheetDocument.Create(ms, SpreadsheetDocumentType.Workbook)) {
                WorkbookPart wbPart = doc.AddWorkbookPart();
                wbPart.Workbook = new S.Workbook();

                WorksheetPart wsPart = wbPart.AddNewPart<WorksheetPart>();
                var sheetData = new S.SheetData();

                int seriesCount = data.Series.Count;
                int totalColumns = seriesCount * 2;
                int maxPoints = data.Series.Max(series => series.XValues.Count);
                int totalRows = maxPoints + 1;
                string lastColumn = ColumnLetter(totalColumns);
                string dimensionRef = $"A1:{lastColumn}{totalRows}";

                wsPart.Worksheet = new S.Worksheet(
                    new S.SheetDimension { Reference = dimensionRef },
                    sheetData);

                var sharedStringsPart = wbPart.AddNewPart<SharedStringTablePart>();
                sharedStringsPart.SharedStringTable = new S.SharedStringTable();

                var stringIndex = new Dictionary<string, int>(StringComparer.Ordinal);
                int GetStringIndex(string value) {
                    if (!stringIndex.TryGetValue(value, out int idx)) {
                        idx = stringIndex.Count;
                        stringIndex[value] = idx;
                        sharedStringsPart.SharedStringTable.AppendChild(new S.SharedStringItem(new S.Text(value)));
                    }
                    return idx;
                }

                var headerRow = new S.Row { RowIndex = 1U, Spans = new ListValue<StringValue> { InnerText = $"1:{totalColumns}" } };
                for (int i = 0; i < seriesCount; i++) {
                    int xColumnIndex = (i * 2) + 1;
                    int yColumnIndex = xColumnIndex + 1;
                    string xHeaderRef = $"{ColumnLetter(xColumnIndex)}1";
                    string yHeaderRef = $"{ColumnLetter(yColumnIndex)}1";

                    headerRow.Append(CreateSharedStringCell(xHeaderRef, GetStringIndex($"{data.Series[i].Name} X")));
                    headerRow.Append(CreateSharedStringCell(yHeaderRef, GetStringIndex(data.Series[i].Name)));
                }
                sheetData.Append(headerRow);

                for (int pointIndex = 0; pointIndex < maxPoints; pointIndex++) {
                    uint excelRow = (uint)(pointIndex + 2);
                    var row = new S.Row { RowIndex = excelRow, Spans = new ListValue<StringValue> { InnerText = $"1:{totalColumns}" } };

                    for (int seriesIndex = 0; seriesIndex < seriesCount; seriesIndex++) {
                        PowerPointScatterChartSeries series = data.Series[seriesIndex];
                        int xColumnIndex = (seriesIndex * 2) + 1;
                        int yColumnIndex = xColumnIndex + 1;

                        if (pointIndex < series.XValues.Count) {
                            row.Append(CreateNumberCell($"{ColumnLetter(xColumnIndex)}{excelRow}", series.XValues[pointIndex]));
                        }
                        if (pointIndex < series.YValues.Count) {
                            row.Append(CreateNumberCell($"{ColumnLetter(yColumnIndex)}{excelRow}", series.YValues[pointIndex]));
                        }
                    }

                    sheetData.Append(row);
                }

                var sheets = wbPart.Workbook.AppendChild(new S.Sheets());
                sheets.Append(new S.Sheet {
                    Id = wbPart.GetIdOfPart(wsPart),
                    SheetId = 1U,
                    Name = "Sheet1"
                });

                sharedStringsPart.SharedStringTable.Save();
                wsPart.Worksheet.Save();
                wbPart.Workbook.Save();
            }

            return ms.ToArray();
        }

        private static C.BarChart CreateBarChart(PowerPointChartData data, out uint categoryAxisId, out uint valueAxisId) {
            categoryAxisId = PowerPointChartAxisIdGenerator.GetNextId();
            valueAxisId = PowerPointChartAxisIdGenerator.GetNextId();

            C.BarChart barChart = new();
            barChart.Append(new C.BarDirection { Val = C.BarDirectionValues.Column });
            barChart.Append(new C.BarGrouping { Val = C.BarGroupingValues.Clustered });
            barChart.Append(new C.VaryColors { Val = false });

            for (int i = 0; i < data.Series.Count; i++) {
                barChart.Append(CreateBarChartSeries(i, data.Series[i], data.Categories));
            }

            barChart.Append(CreateDefaultDataLabels());
            barChart.Append(new C.GapWidth { Val = (UInt16Value)219U });
            barChart.Append(new C.Overlap { Val = (SByteValue)(sbyte)-27 });
            barChart.Append(new C.AxisId { Val = categoryAxisId });
            barChart.Append(new C.AxisId { Val = valueAxisId });
            return barChart;
        }

        private static C.BarChartSeries CreateBarChartSeries(int seriesIndex, PowerPointChartSeries series, IReadOnlyList<string> categories) {
            int lastRow = categories.Count + 1;
            string seriesColumn = ColumnLetter(seriesIndex + 2);
            string seriesNameRef = $"Sheet1!${seriesColumn}$1";
            string categoriesRef = $"Sheet1!$A$2:$A${lastRow}";
            string valuesRef = $"Sheet1!${seriesColumn}$2:${seriesColumn}${lastRow}";

            C.BarChartSeries seriesElement = new(
                new C.Index { Val = (uint)seriesIndex },
                new C.Order { Val = (uint)seriesIndex },
                new C.SeriesText(CreateStringReference(seriesNameRef, new[] { series.Name })),
                new C.InvertIfNegative { Val = false },
                new C.CategoryAxisData(CreateStringReference(categoriesRef, categories)),
                new C.Values(CreateNumberReference(valuesRef, series.Values))
            );

            return seriesElement;
        }

        private static C.LineChart CreateLineChart(PowerPointChartData data, out uint categoryAxisId, out uint valueAxisId) {
            categoryAxisId = PowerPointChartAxisIdGenerator.GetNextId();
            valueAxisId = PowerPointChartAxisIdGenerator.GetNextId();

            C.LineChart lineChart = new(
                new C.Grouping { Val = C.GroupingValues.Standard },
                new C.VaryColors { Val = false });

            for (int i = 0; i < data.Series.Count; i++) {
                lineChart.Append(CreateLineChartSeries(i, data.Series[i], data.Categories));
            }

            lineChart.Append(CreateDefaultDataLabels());
            lineChart.Append(new C.AxisId { Val = categoryAxisId });
            lineChart.Append(new C.AxisId { Val = valueAxisId });
            return lineChart;
        }

        private static C.LineChartSeries CreateLineChartSeries(int seriesIndex, PowerPointChartSeries series, IReadOnlyList<string> categories) {
            int lastRow = categories.Count + 1;
            string seriesColumn = ColumnLetter(seriesIndex + 2);
            string seriesNameRef = $"Sheet1!${seriesColumn}$1";
            string categoriesRef = $"Sheet1!$A$2:$A${lastRow}";
            string valuesRef = $"Sheet1!${seriesColumn}$2:${seriesColumn}${lastRow}";

            C.LineChartSeries seriesElement = new(
                new C.Index { Val = (uint)seriesIndex },
                new C.Order { Val = (uint)seriesIndex },
                new C.SeriesText(CreateStringReference(seriesNameRef, new[] { series.Name })),
                new C.CategoryAxisData(CreateStringReference(categoriesRef, categories)),
                new C.Values(CreateNumberReference(valuesRef, series.Values))
            );

            return seriesElement;
        }

        private static C.ScatterChart CreateScatterChart(PowerPointScatterChartData data, out uint xAxisId, out uint yAxisId) {
            xAxisId = PowerPointChartAxisIdGenerator.GetNextId();
            yAxisId = PowerPointChartAxisIdGenerator.GetNextId();

            C.ScatterChart scatterChart = new(
                new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker },
                new C.VaryColors { Val = false });

            for (int i = 0; i < data.Series.Count; i++) {
                scatterChart.Append(CreateScatterChartSeries(i, data.Series[i]));
            }

            scatterChart.Append(CreateDefaultDataLabels());
            scatterChart.Append(new C.AxisId { Val = xAxisId });
            scatterChart.Append(new C.AxisId { Val = yAxisId });
            return scatterChart;
        }

        private static C.ScatterChartSeries CreateScatterChartSeries(int seriesIndex, PowerPointScatterChartSeries series) {
            string seriesNameRef = GetScatterSeriesNameReference(seriesIndex);
            string xValuesRef = GetScatterXValuesReference(seriesIndex, series.XValues.Count);
            string yValuesRef = GetScatterYValuesReference(seriesIndex, series.YValues.Count);

            C.ScatterChartSeries seriesElement = new(
                new C.Index { Val = (uint)seriesIndex },
                new C.Order { Val = (uint)seriesIndex },
                new C.SeriesText(CreateStringReference(seriesNameRef, new[] { series.Name })),
                new C.XValues(CreateNumberReference(xValuesRef, series.XValues)),
                new C.YValues(CreateNumberReference(yValuesRef, series.YValues))
            );

            return seriesElement;
        }

        private static C.PieChart CreatePieChart(PowerPointChartData data) {
            C.PieChart pieChart = new(
                new C.VaryColors { Val = true });

            for (int i = 0; i < data.Series.Count; i++) {
                pieChart.Append(CreatePieChartSeries(i, data.Series[i], data.Categories));
            }

            pieChart.Append(CreateDefaultDataLabels());
            pieChart.Append(new C.FirstSliceAngle { Val = (UInt16Value)0U });
            return pieChart;
        }

        private static C.DoughnutChart CreateDoughnutChart(PowerPointChartData data) {
            C.DoughnutChart doughnutChart = new(
                new C.VaryColors { Val = true });

            for (int i = 0; i < data.Series.Count; i++) {
                doughnutChart.Append(CreatePieChartSeries(i, data.Series[i], data.Categories));
            }

            doughnutChart.Append(CreateDefaultDataLabels());
            doughnutChart.Append(new C.FirstSliceAngle { Val = (UInt16Value)0U });
            doughnutChart.Append(new C.HoleSize { Val = (ByteValue)50 });
            return doughnutChart;
        }

        private static C.PieChartSeries CreatePieChartSeries(int seriesIndex, PowerPointChartSeries series, IReadOnlyList<string> categories) {
            int lastRow = categories.Count + 1;
            string seriesColumn = ColumnLetter(seriesIndex + 2);
            string seriesNameRef = $"Sheet1!${seriesColumn}$1";
            string categoriesRef = $"Sheet1!$A$2:$A${lastRow}";
            string valuesRef = $"Sheet1!${seriesColumn}$2:${seriesColumn}${lastRow}";

            C.PieChartSeries seriesElement = new(
                new C.Index { Val = (uint)seriesIndex },
                new C.Order { Val = (uint)seriesIndex },
                new C.SeriesText(CreateStringReference(seriesNameRef, new[] { series.Name })),
                new C.CategoryAxisData(CreateStringReference(categoriesRef, categories)),
                new C.Values(CreateNumberReference(valuesRef, series.Values))
            );

            return seriesElement;
        }

        private static C.CategoryAxis CreateCategoryAxis(uint axisId, uint crossingAxisId) {
            C.CategoryAxis axis = new(
                new C.AxisId { Val = axisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
                new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
                new C.MajorTickMark { Val = C.TickMarkValues.None },
                new C.MinorTickMark { Val = C.TickMarkValues.None },
                new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                new C.CrossingAxis { Val = crossingAxisId },
                new C.Crosses { Val = C.CrossesValues.AutoZero },
                new C.AutoLabeled { Val = true },
                new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
                new C.LabelOffset { Val = (UInt16Value)100U },
                new C.NoMultiLevelLabels { Val = false }
            );

            return axis;
        }

        private static void UpdateBarChartSeries(C.BarChart barChart, PowerPointChartData data) {
            List<C.BarChartSeries> existingSeries = barChart.Elements<C.BarChartSeries>().ToList();
            C.BarChartSeries? template = existingSeries.LastOrDefault();

            int seriesCount = data.Series.Count;
            for (int i = 0; i < seriesCount; i++) {
                C.BarChartSeries seriesElement;
                if (i < existingSeries.Count) {
                    seriesElement = existingSeries[i];
                } else {
                    seriesElement = template != null ? (C.BarChartSeries)template.CloneNode(true) : new C.BarChartSeries();
                    InsertSeries(barChart, seriesElement);
                    existingSeries.Add(seriesElement);
                }

                UpdateSeriesIndexOrder(seriesElement, i);
                UpdateSeriesText(seriesElement, i, data.Series[i].Name);
                UpdateCategoryAxisData(seriesElement, data.Categories);
                UpdateValues(seriesElement, i, data.Series[i].Values);
            }

            for (int i = existingSeries.Count - 1; i >= seriesCount; i--) {
                existingSeries[i].Remove();
            }
        }

        private static void UpdateLineChartSeries(C.LineChart lineChart, PowerPointChartData data) {
            List<C.LineChartSeries> existingSeries = lineChart.Elements<C.LineChartSeries>().ToList();
            C.LineChartSeries? template = existingSeries.LastOrDefault();

            int seriesCount = data.Series.Count;
            for (int i = 0; i < seriesCount; i++) {
                C.LineChartSeries seriesElement;
                if (i < existingSeries.Count) {
                    seriesElement = existingSeries[i];
                } else {
                    seriesElement = template != null ? (C.LineChartSeries)template.CloneNode(true) : new C.LineChartSeries();
                    InsertSeries(lineChart, seriesElement);
                    existingSeries.Add(seriesElement);
                }

                UpdateSeriesIndexOrder(seriesElement, i);
                UpdateSeriesText(seriesElement, i, data.Series[i].Name);
                UpdateCategoryAxisData(seriesElement, data.Categories);
                UpdateValues(seriesElement, i, data.Series[i].Values);
            }

            for (int i = existingSeries.Count - 1; i >= seriesCount; i--) {
                existingSeries[i].Remove();
            }
        }

        private static void UpdateAreaChartSeries(C.AreaChart areaChart, PowerPointChartData data) {
            List<C.AreaChartSeries> existingSeries = areaChart.Elements<C.AreaChartSeries>().ToList();
            C.AreaChartSeries? template = existingSeries.LastOrDefault();

            int seriesCount = data.Series.Count;
            for (int i = 0; i < seriesCount; i++) {
                C.AreaChartSeries seriesElement;
                if (i < existingSeries.Count) {
                    seriesElement = existingSeries[i];
                } else {
                    seriesElement = template != null ? (C.AreaChartSeries)template.CloneNode(true) : new C.AreaChartSeries();
                    InsertSeries(areaChart, seriesElement);
                    existingSeries.Add(seriesElement);
                }

                UpdateSeriesIndexOrder(seriesElement, i);
                UpdateSeriesText(seriesElement, i, data.Series[i].Name);
                UpdateCategoryAxisData(seriesElement, data.Categories);
                UpdateValues(seriesElement, i, data.Series[i].Values);
            }

            for (int i = existingSeries.Count - 1; i >= seriesCount; i--) {
                existingSeries[i].Remove();
            }
        }

        private static void UpdateScatterChartSeries(C.ScatterChart scatterChart, PowerPointScatterChartData data) {
            List<C.ScatterChartSeries> existingSeries = scatterChart.Elements<C.ScatterChartSeries>().ToList();
            C.ScatterChartSeries? template = existingSeries.LastOrDefault();

            int seriesCount = data.Series.Count;
            for (int i = 0; i < seriesCount; i++) {
                C.ScatterChartSeries seriesElement;
                if (i < existingSeries.Count) {
                    seriesElement = existingSeries[i];
                } else {
                    seriesElement = template != null ? (C.ScatterChartSeries)template.CloneNode(true) : new C.ScatterChartSeries();
                    InsertSeries(scatterChart, seriesElement);
                    existingSeries.Add(seriesElement);
                }

                UpdateSeriesIndexOrder(seriesElement, i);
                UpdateScatterSeriesText(seriesElement, i, data.Series[i].Name);
                UpdateXValues(seriesElement, i, data.Series[i].XValues);
                UpdateYValues(seriesElement, i, data.Series[i].YValues);
            }

            for (int i = existingSeries.Count - 1; i >= seriesCount; i--) {
                existingSeries[i].Remove();
            }
        }

        private static void UpdateSeriesIndexOrder(OpenXmlCompositeElement series, int index) {
            C.Index idx = series.GetFirstChild<C.Index>() ?? new C.Index();
            idx.Val = (uint)index;
            if (idx.Parent == null) {
                series.PrependChild(idx);
            }

            C.Order order = series.GetFirstChild<C.Order>() ?? new C.Order();
            order.Val = (uint)index;
            if (order.Parent == null) {
                series.InsertAfter(order, idx);
            }
        }

        private static void UpdateSeriesText(OpenXmlCompositeElement series, int seriesIndex, string seriesName) {
            string seriesColumn = ColumnLetter(seriesIndex + 2);
            string seriesNameRef = $"Sheet1!${seriesColumn}$1";
            C.SeriesText seriesText = series.GetFirstChild<C.SeriesText>() ?? new C.SeriesText();
            seriesText.RemoveAllChildren<C.StringReference>();
            seriesText.RemoveAllChildren<C.StringLiteral>();
            seriesText.Append(CreateStringReference(seriesNameRef, new[] { seriesName }));

            if (seriesText.Parent == null) {
                OpenXmlElement? insertAfter = series.GetFirstChild<C.Order>();
                insertAfter ??= series.GetFirstChild<C.Index>();
                if (insertAfter != null) {
                    series.InsertAfter(seriesText, insertAfter);
                } else {
                    series.PrependChild(seriesText);
                }
            }
        }

        private static void UpdateCategoryAxisData(OpenXmlCompositeElement series, IReadOnlyList<string> categories) {
            int lastRow = categories.Count + 1;
            string categoriesRef = $"Sheet1!$A$2:$A${lastRow}";
            C.CategoryAxisData categoryAxisData = series.GetFirstChild<C.CategoryAxisData>() ?? new C.CategoryAxisData();
            categoryAxisData.RemoveAllChildren<C.StringReference>();
            categoryAxisData.RemoveAllChildren<C.StringLiteral>();
            categoryAxisData.Append(CreateStringReference(categoriesRef, categories));

            if (categoryAxisData.Parent == null) {
                series.Append(categoryAxisData);
            }
        }

        private static void UpdateScatterSeriesText(C.ScatterChartSeries series, int seriesIndex, string seriesName) {
            string seriesNameRef = GetScatterSeriesNameReference(seriesIndex);
            C.SeriesText seriesText = series.GetFirstChild<C.SeriesText>() ?? new C.SeriesText();
            seriesText.RemoveAllChildren<C.StringReference>();
            seriesText.RemoveAllChildren<C.StringLiteral>();
            seriesText.Append(CreateStringReference(seriesNameRef, new[] { seriesName }));

            if (seriesText.Parent == null) {
                OpenXmlElement? insertAfter = series.GetFirstChild<C.Order>();
                insertAfter ??= series.GetFirstChild<C.Index>();
                if (insertAfter != null) {
                    series.InsertAfter(seriesText, insertAfter);
                } else {
                    series.PrependChild(seriesText);
                }
            }
        }

        private static void UpdateValues(OpenXmlCompositeElement series, int seriesIndex, IReadOnlyList<double> values) {
            int lastRow = values.Count + 1;
            string seriesColumn = ColumnLetter(seriesIndex + 2);
            string valuesRef = $"Sheet1!${seriesColumn}$2:${seriesColumn}${lastRow}";
            C.Values valueElement = series.GetFirstChild<C.Values>() ?? new C.Values();
            valueElement.RemoveAllChildren<C.NumberReference>();
            valueElement.RemoveAllChildren<C.NumberLiteral>();
            valueElement.Append(CreateNumberReference(valuesRef, values));

            if (valueElement.Parent == null) {
                series.Append(valueElement);
            }
        }

        private static void UpdateXValues(C.ScatterChartSeries series, int seriesIndex, IReadOnlyList<double> values) {
            string valuesRef = GetScatterXValuesReference(seriesIndex, values.Count);
            C.XValues xValueElement = series.GetFirstChild<C.XValues>() ?? new C.XValues();
            xValueElement.RemoveAllChildren<C.NumberReference>();
            xValueElement.RemoveAllChildren<C.NumberLiteral>();
            xValueElement.Append(CreateNumberReference(valuesRef, values));

            if (xValueElement.Parent == null) {
                series.Append(xValueElement);
            }
        }

        private static void UpdateYValues(C.ScatterChartSeries series, int seriesIndex, IReadOnlyList<double> values) {
            string valuesRef = GetScatterYValuesReference(seriesIndex, values.Count);
            C.YValues yValueElement = series.GetFirstChild<C.YValues>() ?? new C.YValues();
            yValueElement.RemoveAllChildren<C.NumberReference>();
            yValueElement.RemoveAllChildren<C.NumberLiteral>();
            yValueElement.Append(CreateNumberReference(valuesRef, values));

            if (yValueElement.Parent == null) {
                series.Append(yValueElement);
            }
        }

        private static void InsertSeries(OpenXmlCompositeElement chart, OpenXmlElement series) {
            OpenXmlElement? insertBefore = chart.ChildElements.FirstOrDefault(child =>
                child is C.DataLabels ||
                child is C.GapWidth ||
                child is C.Overlap ||
                child is C.AxisId ||
                child is C.Marker ||
                child is C.Smooth ||
                child is C.SeriesLines);

            if (insertBefore != null) {
                chart.InsertBefore(series, insertBefore);
            } else {
                chart.Append(series);
            }
        }

        private static string GetScatterSeriesNameReference(int seriesIndex) {
            string yColumn = ColumnLetter((seriesIndex * 2) + 2);
            return $"Sheet1!${yColumn}$1";
        }

        private static string GetScatterXValuesReference(int seriesIndex, int pointCount) {
            string xColumn = ColumnLetter((seriesIndex * 2) + 1);
            int lastRow = pointCount + 1;
            return $"Sheet1!${xColumn}$2:${xColumn}${lastRow}";
        }

        private static string GetScatterYValuesReference(int seriesIndex, int pointCount) {
            string yColumn = ColumnLetter((seriesIndex * 2) + 2);
            int lastRow = pointCount + 1;
            return $"Sheet1!${yColumn}$2:${yColumn}${lastRow}";
        }

        private static C.ValueAxis CreateValueAxis(uint axisId, uint crossingAxisId) {
            return CreateValueAxis(axisId, crossingAxisId, C.AxisPositionValues.Left);
        }

        private static C.ValueAxis CreateValueAxis(uint axisId, uint crossingAxisId,
            C.AxisPositionValues axisPosition) {
            C.ValueAxis axis = new(
                new C.AxisId { Val = axisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = axisPosition },
                new C.MajorGridlines(),
                new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
                new C.MajorTickMark { Val = C.TickMarkValues.None },
                new C.MinorTickMark { Val = C.TickMarkValues.None },
                new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                new C.CrossingAxis { Val = crossingAxisId },
                new C.Crosses { Val = C.CrossesValues.AutoZero },
                new C.CrossBetween { Val = C.CrossBetweenValues.Between }
            );

            return axis;
        }

        private static C.DataLabels CreateDefaultDataLabels() {
            return new C.DataLabels(
                new C.ShowLegendKey { Val = false },
                new C.ShowValue { Val = false },
                new C.ShowCategoryName { Val = false },
                new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false }
            );
        }

        private static C.StringReference CreateStringReference(string formula, IReadOnlyList<string> values) {
            C.StringCache cache = new();
            cache.Append(new C.PointCount { Val = (uint)values.Count });
            for (int i = 0; i < values.Count; i++) {
                cache.Append(new C.StringPoint {
                    Index = (uint)i,
                    NumericValue = new C.NumericValue { Text = values[i] ?? string.Empty }
                });
            }

            return new C.StringReference(
                new C.Formula { Text = formula },
                cache);
        }

        private static C.NumberReference CreateNumberReference(string formula, IReadOnlyList<double> values) {
            C.NumberingCache cache = new();
            cache.Append(new C.FormatCode { Text = "General" });
            cache.Append(new C.PointCount { Val = (uint)values.Count });
            for (int i = 0; i < values.Count; i++) {
                cache.Append(new C.NumericPoint {
                    Index = (uint)i,
                    NumericValue = new C.NumericValue { Text = values[i].ToString(CultureInfo.InvariantCulture) }
                });
            }

            return new C.NumberReference(
                new C.Formula { Text = formula },
                cache);
        }

        private static string ColumnLetter(int column) {
            int dividend = column;
            string columnName = string.Empty;
            while (dividend > 0) {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar('A' + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }
            return columnName;
        }

        private static S.Cell CreateSharedStringCell(string cellReference, int sharedStringIndex) {
            return new S.Cell {
                CellReference = cellReference,
                DataType = S.CellValues.SharedString,
                CellValue = new S.CellValue(sharedStringIndex.ToString(CultureInfo.InvariantCulture))
            };
        }

        private static S.Cell CreateNumberCell(string cellReference, double value) {
            return new S.Cell {
                CellReference = cellReference,
                CellValue = new S.CellValue(value.ToString(CultureInfo.InvariantCulture))
            };
        }

    }
}
