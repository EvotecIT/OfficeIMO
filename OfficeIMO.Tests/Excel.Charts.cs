using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ExcelCharts_CanCreateChartFromData() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Basic.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3", "Q4" },
                    new[] {
                        new ExcelChartSeries("Sales", new[] { 10d, 20d, 25d, 30d }),
                        new ExcelChartSeries("Target", new[] { 12d, 22d, 24d, 32d })
                    });

                sheet.AddChart(data, row: 1, column: 6, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.ColumnClustered, title: "Quarterly");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                Assert.NotNull(wsPart.DrawingsPart);
                Assert.True(wsPart.DrawingsPart!.ChartParts.Any());

                var hiddenSheets = spreadsheet.WorkbookPart.Workbook.Sheets!
                    .OfType<Sheet>()
                    .Where(s => s.State?.Value == SheetStateValues.Hidden)
                    .ToList();
                Assert.Single(hiddenSheets);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }
        }

        [Fact]
        public void Test_ExcelCharts_UpdateData() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Update.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Jan", "Feb" },
                    new[] { new ExcelChartSeries("Sales", new[] { 1d, 2d }) });

                var chart = sheet.AddChart(data, row: 2, column: 6, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.Line, title: "Monthly");

                var updated = new ExcelChartData(
                    new[] { "Jan", "Feb", "Mar" },
                    new[] { new ExcelChartSeries("Sales", new[] { 1d, 2d, 3d }) });

                chart.UpdateData(updated);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
                var plotArea = chart.GetFirstChild<C.PlotArea>()!;
                var lineChart = plotArea.GetFirstChild<C.LineChart>()!;
                var series = lineChart.Elements<C.LineChartSeries>().First();
                var cache = series.GetFirstChild<C.Values>()!
                    .GetFirstChild<C.NumberReference>()!
                    .GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)3, cache.PointCount!.Val!.Value);
            }
        }

        [Fact]
        public void Test_ExcelCharts_Scatter_CanCreateChart() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Scatter.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "1", "2", "3" },
                    new[] { new ExcelChartSeries("Series 1", new[] { 2d, 4d, 6d }, ExcelChartType.Scatter) });

                sheet.AddChart(data, row: 1, column: 6, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.Scatter, title: "Scatter");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;

                Assert.NotNull(plotArea.GetFirstChild<C.ScatterChart>());
                Assert.Equal(2, plotArea.Elements<C.ValueAxis>().Count());
            }
        }

        [Fact]
        public void Test_ExcelCharts_Combo_WithSecondaryAxis() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Combo.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] {
                        new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }, ExcelChartType.ColumnClustered, ExcelChartAxisGroup.Primary),
                        new ExcelChartSeries("Trend", new[] { 12d, 18d, 28d }, ExcelChartType.Line, ExcelChartAxisGroup.Secondary)
                    });

                sheet.AddChart(data, row: 1, column: 6, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.ColumnClustered, title: "Combo");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;

                Assert.NotNull(plotArea.GetFirstChild<C.BarChart>());
                Assert.NotNull(plotArea.GetFirstChild<C.LineChart>());
                Assert.Equal(2, plotArea.Elements<C.CategoryAxis>().Count());
                Assert.Equal(2, plotArea.Elements<C.ValueAxis>().Count());

                Assert.Contains(plotArea.Elements<C.CategoryAxis>(), ax => ax.AxisPosition?.Val != null && ax.AxisPosition.Val.Value == C.AxisPositionValues.Bottom);
                Assert.Contains(plotArea.Elements<C.CategoryAxis>(), ax => ax.AxisPosition?.Val != null && ax.AxisPosition.Val.Value == C.AxisPositionValues.Top);
                Assert.Contains(plotArea.Elements<C.ValueAxis>(), ax => ax.AxisPosition?.Val != null && ax.AxisPosition.Val.Value == C.AxisPositionValues.Left);
                Assert.Contains(plotArea.Elements<C.ValueAxis>(), ax => ax.AxisPosition?.Val != null && ax.AxisPosition.Val.Value == C.AxisPositionValues.Right);
            }
        }

        [Fact]
        public void Test_ExcelCharts_SeriesDataLabels_AndSecondaryAxisFormat() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SeriesLabels.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] {
                        new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }, ExcelChartType.ColumnClustered, ExcelChartAxisGroup.Primary),
                        new ExcelChartSeries("Trend", new[] { 12d, 18d, 28d }, ExcelChartType.Line, ExcelChartAxisGroup.Secondary)
                    });

                var chart = sheet.AddChart(data, row: 1, column: 6, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.ColumnClustered, title: "Combo");
                chart.SetValueAxisNumberFormat("0.00", sourceLinked: false, axisGroup: ExcelChartAxisGroup.Secondary)
                     .SetSeriesDataLabels(1, showValue: true, position: C.DataLabelPositionValues.Top, numberFormat: "0.0");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;

                var secondaryAxis = plotArea.Elements<C.ValueAxis>()
                    .FirstOrDefault(ax => ax.AxisPosition?.Val != null && ax.AxisPosition.Val.Value == C.AxisPositionValues.Right);
                Assert.NotNull(secondaryAxis);
                Assert.Equal("0.00", secondaryAxis!.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value);

                var lineSeries = plotArea.GetFirstChild<C.LineChart>()?
                    .Elements<C.LineChartSeries>()
                    .FirstOrDefault(series => series.GetFirstChild<C.Index>()?.Val?.Value == 1U);
                Assert.NotNull(lineSeries);

                var labels = lineSeries!.GetFirstChild<C.DataLabels>();
                Assert.NotNull(labels);
                Assert.Equal(C.DataLabelPositionValues.Top, labels!.GetFirstChild<C.DataLabelPosition>()?.Val?.Value);
                Assert.Equal("0.0", labels.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value);
            }
        }

        [Fact]
        public void Test_ExcelCharts_Scatter_FromRanges() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.ScatterRanges.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"X"), (1, 2, "Y1"),
                    (2, 1, 1d), (2, 2, 2d),
                    (3, 1, 2d), (3, 2, 4d),
                    (4, 1, 3d), (4, 2, 6d)
                }, null);

                sheet.AddScatterChartFromRanges(new[] {
                    new ExcelChartSeriesRange("Series 1", "A2:A4", "B2:B4")
                }, row: 1, column: 4, widthPixels: 480, heightPixels: 320, title: "Scatter Ranges");

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var scatterChart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!
                    .GetFirstChild<C.PlotArea>()!
                    .GetFirstChild<C.ScatterChart>()!;

                var series = scatterChart.Elements<C.ScatterChartSeries>().First();
                string? xFormula = series.GetFirstChild<C.XValues>()?
                    .GetFirstChild<C.NumberReference>()?
                    .Formula?.Text;
                string? yFormula = series.GetFirstChild<C.YValues>()?
                    .GetFirstChild<C.NumberReference>()?
                    .Formula?.Text;

                Assert.Equal("'Data'!A2:A4", xFormula);
                Assert.Equal("'Data'!B2:B4", yFormula);
            }
        }

        [Fact]
        public void Test_ExcelCharts_Bubble_FromRanges() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.BubbleRanges.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"X"), (1, 2, "Y"), (1, 3, "Size"),
                    (2, 1, 1d), (2, 2, 2d), (2, 3, 4d),
                    (3, 1, 2d), (3, 2, 3d), (3, 3, 5d),
                    (4, 1, 3d), (4, 2, 4d), (4, 3, 6d)
                }, null);

                sheet.AddBubbleChartFromRanges(new[] {
                    new ExcelChartSeriesRange("Bubbles", "A2:A4", "B2:B4", "C2:C4")
                }, row: 8, column: 4, widthPixels: 480, heightPixels: 320, title: "Bubble Ranges");

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var bubbleChart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!
                    .GetFirstChild<C.PlotArea>()!
                    .GetFirstChild<C.BubbleChart>()!;

                var series = bubbleChart.Elements<C.BubbleChartSeries>().First();
                string? sizeFormula = series.GetFirstChild<C.BubbleSize>()?
                    .GetFirstChild<C.NumberReference>()?
                    .Formula?.Text;

                Assert.Equal("'Data'!C2:C4", sizeFormula);
            }
        }

        [Fact]
        public void Test_ExcelCharts_DefaultStylePreset_Applied() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.StylePreset.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                document.DefaultChartStylePreset = ExcelChartStylePreset.Default;
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2" },
                    new[] { new ExcelChartSeries("Sales", new[] { 1d, 2d }) });

                sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.ColumnClustered, title: "Styled");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                Assert.NotNull(chartPart.GetPartsOfType<ChartStylePart>().FirstOrDefault());
                Assert.NotNull(chartPart.GetPartsOfType<ChartColorStylePart>().FirstOrDefault());
            }
        }

        [Fact]
        public void Test_ExcelCharts_DataLabelTextStyle() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.LabelStyle.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2" },
                    new[] { new ExcelChartSeries("Sales", new[] { 1d, 2d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.ColumnClustered, title: "Labels");
                chart.SetSeriesDataLabels(0, showValue: true)
                     .SetSeriesDataLabelTextStyle(0, fontSizePoints: 12, bold: true, color: "FF0000", fontName: "Calibri");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
                var series = plotArea.GetFirstChild<C.BarChart>()!.Elements<C.BarChartSeries>().First();
                var labels = series.GetFirstChild<C.DataLabels>()!;
                var textProps = labels.GetFirstChild<C.TextProperties>()!;
                var paragraph = textProps.GetFirstChild<DocumentFormat.OpenXml.Drawing.Paragraph>()!;
                var runProps = paragraph.GetFirstChild<DocumentFormat.OpenXml.Drawing.ParagraphProperties>()!
                    .GetFirstChild<DocumentFormat.OpenXml.Drawing.DefaultRunProperties>()!;

                Assert.Equal(1200, runProps.FontSize!.Value);
                Assert.True(runProps.Bold!.Value);
                Assert.Equal("Calibri", runProps.GetFirstChild<DocumentFormat.OpenXml.Drawing.LatinFont>()?.Typeface);
                var fill = runProps.GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>();
                Assert.Equal("FF0000", fill?.RgbColorModelHex?.Val?.Value);
            }
        }

        [Fact]
        public void Test_ExcelCharts_DataLabelShapeStyle() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.LabelShape.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2" },
                    new[] { new ExcelChartSeries("Sales", new[] { 1d, 2d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.ColumnClustered, title: "Label Shapes");
                chart.SetSeriesDataLabels(0, showValue: true)
                     .SetSeriesDataLabelShapeStyle(0, fillColor: "FFFFCC", lineColor: "000000", lineWidthPoints: 1.5);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
                var series = plotArea.GetFirstChild<C.BarChart>()!.Elements<C.BarChartSeries>().First();
                var labels = series.GetFirstChild<C.DataLabels>()!;
                var shapeProps = labels.GetFirstChild<C.ChartShapeProperties>()!;
                var fill = shapeProps.GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>();
                var outline = shapeProps.GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>();

                Assert.Equal("FFFFCC", fill?.RgbColorModelHex?.Val?.Value);
                Assert.Equal("000000", outline?.GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()?.RgbColorModelHex?.Val?.Value);
                Assert.Equal(19050, outline?.Width?.Value);
            }
        }

        [Fact]
        public void Test_ExcelCharts_DataLabelLeaderLines() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.LeaderLines.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2" },
                    new[] { new ExcelChartSeries("Sales", new[] { 1d, 2d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.Pie, title: "Leader Lines");
                chart.SetSeriesDataLabels(0, showValue: true)
                     .SetSeriesDataLabelLeaderLines(0, showLeaderLines: true, lineColor: "000000", lineWidthPoints: 1);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
                var series = plotArea.GetFirstChild<C.PieChart>()!.Elements<C.PieChartSeries>().First();
                var labels = series.GetFirstChild<C.DataLabels>()!;
                var showLeaderLines = labels.GetFirstChild<C.ShowLeaderLines>();
                var leaderLines = labels.GetFirstChild<C.LeaderLines>();
                var outline = leaderLines?.GetFirstChild<C.ChartShapeProperties>()?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Outline>();

                Assert.True(showLeaderLines?.Val?.Value ?? false);
                Assert.Equal("000000", outline?.GetFirstChild<DocumentFormat.OpenXml.Drawing.SolidFill>()?.RgbColorModelHex?.Val?.Value);
                Assert.Equal(12700, outline?.Width?.Value);
            }
        }

        [Fact]
        public void Test_ExcelCharts_TitleAndLegendTextStyle() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.TitleLegendStyle.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2" },
                    new[] { new ExcelChartSeries("Sales", new[] { 1d, 2d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.ColumnClustered, title: "Styled");
                chart.SetLegend(C.LegendPositionValues.Right)
                     .SetTitleTextStyle(fontSizePoints: 14, bold: true, color: "1F4E79")
                     .SetLegendTextStyle(fontSizePoints: 9, italic: true, color: "404040", fontName: "Calibri");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;

                var titleRunProps = chart.GetFirstChild<C.Title>()?
                    .GetFirstChild<C.ChartText>()?
                    .GetFirstChild<C.RichText>()?
                    .GetFirstChild<A.Paragraph>()?
                    .GetFirstChild<A.Run>()?
                    .GetFirstChild<A.RunProperties>();

                Assert.NotNull(titleRunProps);
                Assert.Equal(1400, titleRunProps!.FontSize!.Value);
                Assert.True(titleRunProps.Bold!.Value);
                var titleFill = titleRunProps.GetFirstChild<A.SolidFill>();
                Assert.Equal("1F4E79", titleFill?.RgbColorModelHex?.Val?.Value);

                var legendRunProps = chart.GetFirstChild<C.Legend>()?
                    .GetFirstChild<C.TextProperties>()?
                    .GetFirstChild<A.Paragraph>()?
                    .GetFirstChild<A.ParagraphProperties>()?
                    .GetFirstChild<A.DefaultRunProperties>();

                Assert.NotNull(legendRunProps);
                Assert.Equal(900, legendRunProps!.FontSize!.Value);
                Assert.True(legendRunProps.Italic!.Value);
                Assert.Equal("Calibri", legendRunProps.GetFirstChild<A.LatinFont>()?.Typeface);
                var legendFill = legendRunProps.GetFirstChild<A.SolidFill>();
                Assert.Equal("404040", legendFill?.RgbColorModelHex?.Val?.Value);
            }
        }

        [Fact]
        public void Test_ExcelCharts_AxisTitleAndLabelTextStyle() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.AxisTextStyle.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2" },
                    new[] { new ExcelChartSeries("Sales", new[] { 1d, 2d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.ColumnClustered, title: "Axis Styles");
                chart.SetCategoryAxisTitle("Quarter")
                     .SetValueAxisTitle("Revenue")
                     .SetCategoryAxisTitleTextStyle(fontSizePoints: 11, bold: true, color: "006100")
                     .SetValueAxisTitleTextStyle(fontSizePoints: 11, bold: true, color: "006100")
                     .SetCategoryAxisLabelTextStyle(fontSizePoints: 9, color: "404040")
                     .SetValueAxisLabelTextStyle(fontSizePoints: 9, italic: true, fontName: "Calibri");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;

                var categoryAxis = plotArea.Elements<C.CategoryAxis>().First();
                var categoryTitleProps = categoryAxis.GetFirstChild<C.Title>()?
                    .GetFirstChild<C.ChartText>()?
                    .GetFirstChild<C.RichText>()?
                    .GetFirstChild<A.Paragraph>()?
                    .GetFirstChild<A.Run>()?
                    .GetFirstChild<A.RunProperties>();

                Assert.NotNull(categoryTitleProps);
                Assert.Equal(1100, categoryTitleProps!.FontSize!.Value);
                Assert.True(categoryTitleProps.Bold!.Value);

                var categoryLabelProps = categoryAxis.GetFirstChild<C.TextProperties>()?
                    .GetFirstChild<A.Paragraph>()?
                    .GetFirstChild<A.ParagraphProperties>()?
                    .GetFirstChild<A.DefaultRunProperties>();

                Assert.NotNull(categoryLabelProps);
                Assert.Equal(900, categoryLabelProps!.FontSize!.Value);

                var valueAxis = plotArea.Elements<C.ValueAxis>().First();
                var valueTitleProps = valueAxis.GetFirstChild<C.Title>()?
                    .GetFirstChild<C.ChartText>()?
                    .GetFirstChild<C.RichText>()?
                    .GetFirstChild<A.Paragraph>()?
                    .GetFirstChild<A.Run>()?
                    .GetFirstChild<A.RunProperties>();

                Assert.NotNull(valueTitleProps);
                Assert.Equal(1100, valueTitleProps!.FontSize!.Value);

                var valueLabelProps = valueAxis.GetFirstChild<C.TextProperties>()?
                    .GetFirstChild<A.Paragraph>()?
                    .GetFirstChild<A.ParagraphProperties>()?
                    .GetFirstChild<A.DefaultRunProperties>();

                Assert.NotNull(valueLabelProps);
                Assert.Equal(900, valueLabelProps!.FontSize!.Value);
                Assert.True(valueLabelProps.Italic!.Value);
                Assert.Equal("Calibri", valueLabelProps.GetFirstChild<A.LatinFont>()?.Typeface);
            }
        }

        [Fact]
        public void Test_ExcelCharts_AxisGridlinesRotationAndTicks() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.AxisGridlines.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.ColumnClustered, title: "Gridlines");
                chart.SetValueAxisGridlines(showMajor: true, showMinor: true, lineColor: "C0C0C0", lineWidthPoints: 0.75)
                     .SetCategoryAxisLabelRotation(45)
                     .SetValueAxisTickLabelPosition(C.TickLabelPositionValues.Low);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;

                var categoryAxis = plotArea.Elements<C.CategoryAxis>().First();
                var rotation = categoryAxis.GetFirstChild<C.TextProperties>()?
                    .GetFirstChild<A.BodyProperties>()?
                    .Rotation?.Value;
                Assert.Equal(2700000, rotation);

                var valueAxis = plotArea.Elements<C.ValueAxis>().First();
                var major = valueAxis.GetFirstChild<C.MajorGridlines>();
                var minor = valueAxis.GetFirstChild<C.MinorGridlines>();
                Assert.NotNull(major);
                Assert.NotNull(minor);

                var majorOutline = major!.GetFirstChild<C.ChartShapeProperties>()?
                    .GetFirstChild<A.Outline>();
                Assert.Equal("C0C0C0", majorOutline?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value);
                Assert.Equal(9525, majorOutline?.Width?.Value);

                var tickPos = valueAxis.GetFirstChild<C.TickLabelPosition>();
                Assert.Equal(C.TickLabelPositionValues.Low, tickPos?.Val?.Value);
            }
        }

        [Fact]
        public void Test_ExcelCharts_AxisCrossingAndDisplayUnits() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.AxisCrossing.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.ColumnClustered, title: "Axis Crossing");
                chart.SetValueAxisCrossing(C.CrossesValues.Maximum)
                     .SetCategoryAxisCrossing(C.CrossesValues.Minimum)
                     .SetValueAxisCrossBetween(C.CrossBetweenValues.Between)
                     .SetValueAxisDisplayUnits(C.BuiltInUnitValues.Thousands, "Thousands USD", showLabel: true);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;

                var valueAxis = plotArea.Elements<C.ValueAxis>().First();
                Assert.Equal(C.CrossesValues.Maximum, valueAxis.GetFirstChild<C.Crosses>()?.Val?.Value);
                Assert.Equal(C.CrossBetweenValues.Between, valueAxis.GetFirstChild<C.CrossBetween>()?.Val?.Value);
                var displayUnits = valueAxis.GetFirstChild<C.DisplayUnits>();
                Assert.Equal(C.BuiltInUnitValues.Thousands, displayUnits?.GetFirstChild<C.BuiltInUnit>()?.Val?.Value);
                var displayLabel = displayUnits?.GetFirstChild<C.DisplayUnitsLabel>();
                Assert.NotNull(displayLabel);
                Assert.Equal("Thousands USD", displayLabel?.ChartText?.InnerText);

                var categoryAxis = plotArea.Elements<C.CategoryAxis>().First();
                Assert.Equal(C.CrossesValues.Minimum, categoryAxis.GetFirstChild<C.Crosses>()?.Val?.Value);
            }
        }

        [Fact]
        public void Test_ExcelCharts_AxisCrossingAtValue() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.AxisCrossingAt.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.ColumnClustered, title: "Axis Crossing At");
                chart.SetValueAxisCrossing(C.CrossesValues.AutoZero, crossesAt: 2.5);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;

                var valueAxis = plotArea.Elements<C.ValueAxis>().First();
                Assert.Equal(2.5d, (double?)valueAxis.GetFirstChild<C.CrossesAt>()?.Val?.Value);
                Assert.Null(valueAxis.GetFirstChild<C.Crosses>());
            }
        }

        [Fact]
        public void Test_ExcelCharts_ValueAxisScale() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.AxisScale.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.ColumnClustered, title: "Axis Scale");
                chart.SetValueAxisScale(minimum: 0, maximum: 100, majorUnit: 25, minorUnit: 5, reverseOrder: true);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
                var valueAxis = plotArea.Elements<C.ValueAxis>().First();
                var scaling = valueAxis.GetFirstChild<C.Scaling>();

                Assert.Equal(0d, (double?)scaling?.GetFirstChild<C.MinAxisValue>()?.Val?.Value);
                Assert.Equal(100d, (double?)scaling?.GetFirstChild<C.MaxAxisValue>()?.Val?.Value);
                Assert.Equal(C.OrientationValues.MaxMin, scaling?.GetFirstChild<C.Orientation>()?.Val?.Value);
                Assert.Equal(25d, (double?)valueAxis.GetFirstChild<C.MajorUnit>()?.Val?.Value);
                Assert.Equal(5d, (double?)valueAxis.GetFirstChild<C.MinorUnit>()?.Val?.Value);
            }
        }

        [Fact]
        public void Test_ExcelCharts_ScatterXAxisScale() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.ScatterAxisScale.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "1", "2", "3", "4" },
                    new[] { new ExcelChartSeries("Points", new[] { 2d, 4d, 3d, 5d }, ExcelChartType.Scatter) });

                var chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.Scatter, title: "Scatter Axis");
                chart.SetScatterXAxisScale(minimum: 1, maximum: 10, majorUnit: 1, logScale: true);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
                var xAxis = plotArea.Elements<C.ValueAxis>()
                    .First(ax => ax.AxisPosition?.Val?.Value == C.AxisPositionValues.Bottom);
                var scaling = xAxis.GetFirstChild<C.Scaling>();

                Assert.Equal(1d, (double?)scaling?.GetFirstChild<C.MinAxisValue>()?.Val?.Value);
                Assert.Equal(10d, (double?)scaling?.GetFirstChild<C.MaxAxisValue>()?.Val?.Value);
                Assert.Equal(10d, (double?)scaling?.GetFirstChild<C.LogBase>()?.Val?.Value);
                Assert.Equal(1d, (double?)xAxis.GetFirstChild<C.MajorUnit>()?.Val?.Value);
            }
        }

        [Fact]
        public void Test_ExcelCharts_ScatterYAxisScale() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.ScatterYAxisScale.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "1", "2", "3", "4" },
                    new[] { new ExcelChartSeries("Points", new[] { 2d, 4d, 3d, 5d }, ExcelChartType.Scatter) });

                var chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.Scatter, title: "Scatter Axis");
                chart.SetScatterYAxisScale(minimum: 0, maximum: 6, majorUnit: 1);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
                var yAxis = plotArea.Elements<C.ValueAxis>()
                    .First(ax => ax.AxisPosition?.Val?.Value == C.AxisPositionValues.Left);
                var scaling = yAxis.GetFirstChild<C.Scaling>();

                Assert.Equal(0d, (double?)scaling?.GetFirstChild<C.MinAxisValue>()?.Val?.Value);
                Assert.Equal(6d, (double?)scaling?.GetFirstChild<C.MaxAxisValue>()?.Val?.Value);
                Assert.Equal(1d, (double?)yAxis.GetFirstChild<C.MajorUnit>()?.Val?.Value);
            }
        }

        [Fact]
        public void Test_ExcelCharts_ScatterXAxisCrossing_RejectsNonPositiveOnLogScale() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.ScatterAxisCrossingLog.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "1", "2", "3" },
                    new[] { new ExcelChartSeries("Points", new[] { 2d, 4d, 3d }, ExcelChartType.Scatter) });

                var chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.Scatter, title: "Scatter Crossing");
                chart.SetScatterXAxisScale(minimum: 1, maximum: 10, logScale: true);

                Assert.Throws<ArgumentOutOfRangeException>(() =>
                    chart.SetScatterXAxisCrossing(C.CrossesValues.AutoZero, crossesAt: 0));
            }
        }

        [Fact]
        public void Test_ExcelCharts_ScatterYAxisCrossing() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.ScatterYAxisCrossing.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "1", "2", "3" },
                    new[] { new ExcelChartSeries("Points", new[] { 2d, 4d, 3d }, ExcelChartType.Scatter) });

                var chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.Scatter, title: "Scatter Crossing");
                chart.SetScatterYAxisCrossing(C.CrossesValues.Minimum, crossesAt: 2d);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
                var yAxis = plotArea.Elements<C.ValueAxis>()
                    .First(ax => ax.AxisPosition?.Val?.Value == C.AxisPositionValues.Left);

                Assert.Equal(2d, (double?)yAxis.GetFirstChild<C.CrossesAt>()?.Val?.Value);
                Assert.Null(yAxis.GetFirstChild<C.Crosses>());
            }
        }

        [Fact]
        public void Test_ExcelCharts_DataLabelTemplateAndPointOverrides() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.LabelTemplate.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.ColumnClustered, title: "Label Template");
                var template = new ExcelChartDataLabelTemplate {
                    ShowValue = true,
                    Position = C.DataLabelPositionValues.Top,
                    NumberFormat = "0.0",
                    FontSizePoints = 9,
                    TextColor = "404040",
                    Separator = " - ",
                    FillColor = "FFFFFF",
                    LineColor = "000000",
                    LineWidthPoints = 0.5
                };
                chart.SetSeriesDataLabelTemplate(0, template)
                     .SetSeriesDataLabelForPoint(0, 1, showValue: true, position: C.DataLabelPositionValues.OutsideEnd,
                        numberFormat: "0.00")
                     .SetSeriesDataLabelSeparatorForPoint(0, 1, " | ")
                     .SetSeriesDataLabelTextStyleForPoint(0, 1, fontSizePoints: 11, bold: true, color: "FF0000")
                     .SetSeriesDataLabelShapeStyleForPoint(0, 1, fillColor: "FFFFCC", lineColor: "000000",
                        lineWidthPoints: 1);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
                var series = plotArea.GetFirstChild<C.BarChart>()!.Elements<C.BarChartSeries>().First();
                var labels = series.GetFirstChild<C.DataLabels>()!;
                Assert.True(labels.GetFirstChild<C.ShowValue>()?.Val?.Value ?? false);
                Assert.Equal(C.DataLabelPositionValues.Top, labels.GetFirstChild<C.DataLabelPosition>()?.Val?.Value);
                Assert.Equal(" - ", labels.GetFirstChild<C.Separator>()?.Text);

                var pointLabel = labels.Elements<C.DataLabel>()
                    .FirstOrDefault(l => l.GetFirstChild<C.Index>()?.Val?.Value == 1U);
                Assert.NotNull(pointLabel);
                Assert.Equal(C.DataLabelPositionValues.OutsideEnd, pointLabel!.GetFirstChild<C.DataLabelPosition>()?.Val?.Value);
                Assert.Equal("0.00", pointLabel.GetFirstChild<C.NumberingFormat>()?.FormatCode?.Value);
                Assert.Equal(" | ", pointLabel.GetFirstChild<C.Separator>()?.Text);

                var pointTextProps = pointLabel.GetFirstChild<C.TextProperties>()?
                    .GetFirstChild<A.Paragraph>()?
                    .GetFirstChild<A.ParagraphProperties>()?
                    .GetFirstChild<A.DefaultRunProperties>();
                Assert.Equal(1100, pointTextProps?.FontSize?.Value);
                Assert.True(pointTextProps?.Bold?.Value ?? false);
                Assert.Equal("FF0000", pointTextProps?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value);

                var pointShape = pointLabel.GetFirstChild<C.ChartShapeProperties>();
                Assert.Equal("FFFFCC", pointShape?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value);
                Assert.Equal("000000", pointShape?.GetFirstChild<A.Outline>()?
                    .GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value);
            }
        }

        [Fact]
        public void Test_ExcelCharts_SeriesTrendline() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Trendline.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.Line, title: "Trendline");
                chart.SetSeriesTrendline(0, C.TrendlineValues.Polynomial, order: 2,
                    displayEquation: true, displayRSquared: true, lineColor: "FF0000", lineWidthPoints: 1.5);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
                var series = plotArea.GetFirstChild<C.LineChart>()!.Elements<C.LineChartSeries>().First();
                var trendline = series.GetFirstChild<C.Trendline>();
                Assert.NotNull(trendline);

                var trendType = trendline!.GetFirstChild<C.TrendlineType>();
                Assert.Equal(C.TrendlineValues.Polynomial, trendType?.Val?.Value);
                Assert.Equal((int?)2, (int?)trendline.GetFirstChild<C.PolynomialOrder>()?.Val?.Value);
                Assert.True(trendline.GetFirstChild<C.DisplayEquation>()?.Val?.Value ?? false);
                Assert.True(trendline.GetFirstChild<C.DisplayRSquaredValue>()?.Val?.Value ?? false);

                var outline = trendline.GetFirstChild<C.ChartShapeProperties>()?
                    .GetFirstChild<A.Outline>();
                Assert.Equal("FF0000", outline?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value);
                Assert.Equal(19050, outline?.Width?.Value);
            }
        }

        [Fact]
        public void Test_ExcelCharts_ChartAndPlotAreaStyle() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.AreaStyle.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.ColumnClustered, title: "Area Style");
                chart.SetChartAreaStyle(fillColor: "F2F2F2", lineColor: "404040", lineWidthPoints: 1.25)
                     .SetPlotAreaStyle(fillColor: "FFFFFF", lineColor: "00B0F0", lineWidthPoints: 0.5);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var chartSpace = chartPart.ChartSpace;

                var chartProps = chartSpace.GetFirstChild<C.ShapeProperties>();
                Assert.NotNull(chartProps);
                var chartFill = chartProps!.GetFirstChild<A.SolidFill>();
                var chartOutline = chartProps.GetFirstChild<A.Outline>();
                Assert.Equal("F2F2F2", chartFill?.RgbColorModelHex?.Val?.Value);
                Assert.Equal("404040", chartOutline?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value);
                Assert.Equal((int)Math.Round(1.25d * 12700d), chartOutline?.Width?.Value);

                var plotArea = chartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
                var plotProps = plotArea.GetFirstChild<C.ShapeProperties>();
                Assert.NotNull(plotProps);
                var plotFill = plotProps!.GetFirstChild<A.SolidFill>();
                var plotOutline = plotProps.GetFirstChild<A.Outline>();
                Assert.Equal("FFFFFF", plotFill?.RgbColorModelHex?.Val?.Value);
                Assert.Equal("00B0F0", plotOutline?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val?.Value);
                Assert.Equal((int)Math.Round(0.5d * 12700d), plotOutline?.Width?.Value);
            }
        }
    }
}
