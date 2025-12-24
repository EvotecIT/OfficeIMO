using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using S = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.PowerPoint {
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

        internal static void PopulateChart(ChartPart chartPart, string embeddedRelId, PowerPointChartData data) {
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

            uint categoryAxisId;
            uint valueAxisId;
            C.BarChart barChart = CreateBarChart(data, out categoryAxisId, out valueAxisId);
            plotArea.Append(barChart);
            plotArea.Append(CreateCategoryAxis(categoryAxisId, valueAxisId));
            plotArea.Append(CreateValueAxis(valueAxisId, categoryAxisId));

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

        private static C.ValueAxis CreateValueAxis(uint axisId, uint crossingAxisId) {
            C.ValueAxis axis = new(
                new C.AxisId { Val = axisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = C.AxisPositionValues.Left },
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
