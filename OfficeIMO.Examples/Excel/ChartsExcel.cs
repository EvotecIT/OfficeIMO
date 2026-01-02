using System;
using System.IO;
using System.Collections.Generic;
using OfficeIMO.Excel;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Examples.Excel {
    /// <summary>
    /// Demonstrates Excel chart creation.
    /// </summary>
    public class ChartsExcel {
        public static void Charts_Basic(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Creating chart demo");
            string filePath = Path.Combine(folderPath, "ExcelCharts.Basic.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                document.DefaultChartStylePreset = ExcelChartStylePreset.Default;
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3", "Q4" },
                    new[] {
                        new ExcelChartSeries("Sales", new[] { 10d, 20d, 25d, 30d }),
                        new ExcelChartSeries("Target", new[] { 12d, 22d, 24d, 32d })
                    });

                sheet.AddChart(data, row: 2, column: 6, widthPixels: 640, heightPixels: 360,
                    type: ExcelChartType.ColumnClustered, title: "Quarterly Sales");

                document.Save(openExcel);
            }
        }

        public static void Charts_ComboAndScatter(string folderPath, bool openExcel) {
            Console.WriteLine("[*] Excel - Creating combo + scatter chart demo");
            string filePath = Path.Combine(folderPath, "ExcelCharts.ComboScatter.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                document.DefaultChartStylePreset = ExcelChartStylePreset.Default;
                var sheet = document.AddWorkSheet("Summary");

                var comboData = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3", "Q4" },
                    new[] {
                        new ExcelChartSeries("Sales", new[] { 10d, 20d, 25d, 30d }, ExcelChartType.ColumnClustered, ExcelChartAxisGroup.Primary),
                        new ExcelChartSeries("Trend", new[] { 12d, 18d, 28d, 35d }, ExcelChartType.Line, ExcelChartAxisGroup.Secondary)
                    });

                var comboChart = sheet.AddChart(comboData, row: 2, column: 6, widthPixels: 640, heightPixels: 360,
                    type: ExcelChartType.ColumnClustered, title: "Sales vs Trend");
                comboChart.ApplyStylePreset()
                          .SetSeriesMarker(1, C.MarkerStyleValues.Circle, size: 6, lineColor: "4472C4");
                comboChart.SetValueAxisNumberFormat("0.00", sourceLinked: false, axisGroup: ExcelChartAxisGroup.Secondary)
                          .SetSeriesDataLabels(1, showValue: true, position: C.DataLabelPositionValues.Top, numberFormat: "0.0")
                          .SetSeriesDataLabelTextStyle(1, fontSizePoints: 9, color: "1F4E79")
                          .SetSeriesDataLabelShapeStyle(1, fillColor: "FFFFFF", lineColor: "1F4E79", lineWidthPoints: 0.5);
                comboChart.SetSeriesDataLabelLeaderLines(1, showLeaderLines: true, lineColor: "1F4E79", lineWidthPoints: 0.5);
                comboChart.SetLegend(C.LegendPositionValues.Right)
                          .SetTitleTextStyle(fontSizePoints: 14, bold: true, color: "1F4E79")
                          .SetLegendTextStyle(fontSizePoints: 9, color: "404040")
                          .SetCategoryAxisTitle("Quarter")
                          .SetValueAxisTitle("Revenue")
                          .SetCategoryAxisTitleTextStyle(fontSizePoints: 10, bold: true)
                          .SetValueAxisLabelTextStyle(fontSizePoints: 9, color: "404040")
                          .SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: "C0C0C0", lineWidthPoints: 0.75)
                          .SetCategoryAxisLabelRotation(45)
                          .SetValueAxisTickLabelPosition(C.TickLabelPositionValues.Low);
                comboChart.SetCategoryAxisReverseOrder()
                          .SetValueAxisScale(minimum: 0, maximum: 40, majorUnit: 10, minorUnit: 5);
                comboChart.SetValueAxisCrossing(C.CrossesValues.Maximum)
                          .SetCategoryAxisCrossing(C.CrossesValues.Minimum)
                          .SetValueAxisCrossBetween(C.CrossBetweenValues.Between)
                          .SetValueAxisDisplayUnits(C.BuiltInUnitValues.Thousands, "Thousands USD", showLabel: true);
                comboChart.SetChartAreaStyle(fillColor: "F2F2F2", lineColor: "404040", lineWidthPoints: 1)
                          .SetPlotAreaStyle(fillColor: "FFFFFF", lineColor: "BFBFBF", lineWidthPoints: 0.75);
                comboChart.SetSeriesTrendline(1, C.TrendlineValues.Linear, displayEquation: true, displayRSquared: true,
                    lineColor: "A5A5A5", lineWidthPoints: 1);
                var labelTemplate = new ExcelChartDataLabelTemplate {
                    ShowValue = true,
                    Position = C.DataLabelPositionValues.Top,
                    NumberFormat = "0.0",
                    FontSizePoints = 9,
                    TextColor = "404040",
                    Separator = " - "
                };
                comboChart.SetSeriesDataLabelTemplate(1, labelTemplate)
                          .SetSeriesDataLabelForPoint(1, 2, showValue: true, position: C.DataLabelPositionValues.OutsideEnd,
                            numberFormat: "0.00")
                          .SetSeriesDataLabelSeparatorForPoint(1, 2, " | ")
                          .SetSeriesDataLabelTextStyleForPoint(1, 2, fontSizePoints: 11, bold: true, color: "FF0000");

                var scatterData = new ExcelChartData(
                    new[] { "1", "2", "3", "4" },
                    new[] { new ExcelChartSeries("Points", new[] { 2d, 4d, 3d, 5d }, ExcelChartType.Scatter) });

                var scatterChart = sheet.AddChart(scatterData, row: 22, column: 6, widthPixels: 640, heightPixels: 360,
                    type: ExcelChartType.Scatter, title: "Scatter Sample");
                scatterChart.SetScatterXAxisScale(minimum: 1, maximum: 10, majorUnit: 1, logScale: true);
                scatterChart.SetScatterYAxisScale(minimum: 0, maximum: 6, majorUnit: 1);
                scatterChart.SetScatterYAxisCrossing(C.CrossesValues.Minimum, crossesAt: 2d);

                var rangeCells = new List<(int Row, int Column, object Value)> {
                    (30, 1, "X"), (30, 2, "Y1"), (30, 3, "Y2"), (30, 4, "Size"),
                    (31, 1, 1d), (31, 2, 2d), (31, 3, 3d), (31, 4, 4d),
                    (32, 1, 2d), (32, 2, 4d), (32, 3, 2d), (32, 4, 5d),
                    (33, 1, 3d), (33, 2, 3d), (33, 3, 5d), (33, 4, 6d)
                };
                sheet.CellValues(rangeCells, null);

                sheet.AddScatterChartFromRanges(new[] {
                    new ExcelChartSeriesRange("Series 1", "A31:A33", "B31:B33"),
                    new ExcelChartSeriesRange("Series 2", "A31:A33", "C31:C33")
                }, row: 38, column: 6, widthPixels: 640, heightPixels: 360, title: "Scatter (Ranges)");

                sheet.AddBubbleChartFromRanges(new[] {
                    new ExcelChartSeriesRange("Bubbles", "A31:A33", "B31:B33", "D31:D33")
                }, row: 54, column: 6, widthPixels: 640, heightPixels: 360, title: "Bubble");

                document.Save(openExcel);
            }
        }
    }
}
