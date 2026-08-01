using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_ModernCharts_AuthorRoundTripMutatePreserveAndRemoveChartEx() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.ModernCharts.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet dashboard = document.AddWorksheet("Dashboard");
                var data = new ExcelChartData(
                    new[] { "East", "West", "North" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d, 15d }) });
                ExcelModernChart chart = dashboard.AddModernChart(
                    data,
                    row: 2,
                    column: 3,
                    chartType: ExcelModernChartType.Funnel,
                    title: "Sales Funnel");
                chart.Name = "NativeFunnel";

                Assert.Equal(ExcelModernChartType.Funnel, chart.ChartType);
                Assert.Equal("Sales Funnel", chart.Title);
                Assert.NotNull(chart.DataRange);
                Assert.Single(dashboard.ModernCharts);
                Assert.Empty(dashboard.Charts);
                ExcelFeatureFinding charts = Assert.Single(document.InspectFeatures().FindFeatures("Charts"));
                Assert.Equal(1, charts.Count);
                ExcelSheet secondary = document.AddWorksheet("Secondary");
                ExcelModernChart secondChart = secondary.AddModernChart(
                    data,
                    row: 1,
                    column: 1,
                    chartType: ExcelModernChartType.Treemap);
                ExcelChartDataRange secondRange = secondChart.DataRange!;
                ExtendedChartPart chartPart = dashboard.WorksheetPart.DrawingsPart!.ExtendedChartParts.Single();
                var importedExtension = new OpenXmlUnknownElement(
                    "cx",
                    "extLst",
                    "http://schemas.microsoft.com/office/drawing/2014/chartex");
                importedExtension.Append(new OpenXmlUnknownElement("x", "vendorState", "urn:vendor:chart"));
                chartPart.ChartSpace!.Append(importedExtension);

                chart.SetTitle("Updated Funnel")
                    .SetChartType(ExcelModernChartType.Waterfall)
                    .SetPlacement(4, 5, 700, 400)
                    .UpdateData(new ExcelChartData(
                        new[] { "East", "West", "North", "South" },
                        new[] {
                            new ExcelChartSeries("Sales", new[] { 10d, 20d, 15d, 12d }),
                            new ExcelChartSeries("Plan", new[] { 11d, 18d, 16d, 13d })
                        }));
                Assert.Equal(ExcelModernChartType.Waterfall, chart.ChartType);
                Assert.Equal("Updated Funnel", chart.Title);
                Assert.Equal(4, chart.DataRange!.CategoryCount);
                Assert.Equal(2, chart.DataRange.SeriesCount);
                Assert.True(chart.DataRange.StartRow > secondRange.CategoryEndRow);
                Assert.Equal(
                    "East",
                    document[secondRange.SheetName].CellAt(secondRange.CategoryStartRow, secondRange.CategoryStartColumn).GetValue<string>());
                Assert.NotNull(chartPart.ChartSpace.Descendants().FirstOrDefault(element =>
                    element.LocalName == "vendorState" && element.NamespaceUri == "urn:vendor:chart"));
                Xdr.OneCellAnchor anchor = Assert.Single(dashboard.WorksheetPart.DrawingsPart.WorksheetDrawing!
                    .Elements<Xdr.OneCellAnchor>());
                Assert.Equal("4", anchor.FromMarker!.ColumnId!.Text);
                Assert.Equal("3", anchor.FromMarker.RowId!.Text);

                dashboard.InsertRows(2, 2);
                Assert.Equal("5", anchor.FromMarker.RowId!.Text);
                var errors = document.ValidateOpenXml().ToArray();
                Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet dashboard = document["Dashboard"];
                ExcelModernChart chart = Assert.Single(dashboard.ModernCharts);
                Assert.Equal("NativeFunnel", chart.Name);
                Assert.Equal(ExcelModernChartType.Waterfall, chart.ChartType);
                Assert.Equal("Updated Funnel", chart.Title);
                Assert.NotNull(chart.DataRange);
                chart.SetChartType(ExcelModernChartType.Treemap);
                chart.UpdateData(new ExcelChartData(
                    new[] { "East", "West" },
                    new[] {
                        new ExcelChartSeries("Sales", new[] { 9d, 19d }),
                        new ExcelChartSeries("Plan", new[] { 10d, 20d })
                    }));
                Assert.Equal(ExcelModernChartType.Treemap, chart.ChartType);
                Assert.Equal(2, chart.DataRange!.CategoryCount);
                Assert.NotNull(dashboard.WorksheetPart.DrawingsPart!.ExtendedChartParts.Single()
                    .ChartSpace!.Descendants().FirstOrDefault(element => element.LocalName == "vendorState"));
                chart.Remove();
                Assert.Empty(dashboard.ModernCharts);
                Assert.Null(dashboard.WorksheetPart.DrawingsPart);
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Theory]
        [InlineData(ExcelModernChartType.Funnel)]
        [InlineData(ExcelModernChartType.Waterfall)]
        [InlineData(ExcelModernChartType.BoxWhisker)]
        [InlineData(ExcelModernChartType.Treemap)]
        [InlineData(ExcelModernChartType.Sunburst)]
        public void Test_ModernCharts_AllSupportedLayoutsProduceNativeChartEx(ExcelModernChartType chartType) {
            using ExcelDocument document = ExcelDocument.Create();
            ExcelSheet sheet = document.AddWorksheet("Dashboard");
            ExcelModernChart chart = sheet.AddModernChart(
                new ExcelChartData(
                    new[] { "A", "B" },
                    new[] { new ExcelChartSeries("Value", new[] { 1d, 2d }) }),
                1,
                1,
                chartType);

            Assert.Equal(chartType, chart.ChartType);
            Assert.Single(sheet.WorksheetPart.DrawingsPart!.ExtendedChartParts);
            Assert.Empty(sheet.WorksheetPart.DrawingsPart.ChartParts);
            var errors = document.ValidateOpenXml().ToArray();
            Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors));
        }

        [Fact]
        public void Test_ModernCharts_ImportedVisibleWorksheetDataIsInspectionOnly() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.ModernCharts.ImportedVisibleData.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet visibleData = document.AddWorksheet("VisibleData");
                visibleData.CellValue(1, 1, "Category");
                visibleData.CellValue(1, 2, "Value");
                visibleData.CellValue(2, 1, "Keep A");
                visibleData.CellValue(2, 2, 10d);
                visibleData.CellValue(3, 1, "Keep B");
                visibleData.CellValue(3, 2, 20d);
                ExcelSheet dashboard = document.AddWorksheet("Dashboard");
                dashboard.AddModernChart(
                    new ExcelChartData(
                        new[] { "A", "B" },
                        new[] { new ExcelChartSeries("Value", new[] { 1d, 2d }) }),
                    1,
                    1,
                    ExcelModernChartType.Funnel);
                ExtendedChartPart part = dashboard.WorksheetPart.DrawingsPart!.ExtendedChartParts.Single();
                var formulas = part.ChartSpace!.Descendants<Formula>().ToArray();
                Assert.Equal(2, formulas.Length);
                formulas[0].Text = "VisibleData!$A$2:$A$3";
                formulas[1].Text = "VisibleData!$B$2:$B$3";
                part.ChartSpace.Save();
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelModernChart chart = Assert.Single(document["Dashboard"].ModernCharts);
                Assert.Null(chart.DataRange);
                Assert.Throws<InvalidOperationException>(() => chart.UpdateData(
                    new ExcelChartData(
                        new[] { "Changed" },
                        new[] { new ExcelChartSeries("Value", new[] { 99d }) })));
                Assert.Equal("Keep A", document["VisibleData"].CellAt(2, 1).GetValue<string>());
                Assert.Equal(10d, document["VisibleData"].CellAt(2, 2).GetValue<double>());
            }
        }

        [Fact]
        public void Test_ModernCharts_VisibleReservedSheetNameDoesNotCaptureAuthoredData() {
            string filePath = Path.Combine(_directoryWithFiles, "Excel.ModernCharts.ReservedSheetCollision.xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet reserved = document.AddWorksheet("OfficeIMO_ChartData");
                reserved.CellValue(1, 1, "user data");
                ExcelSheet dashboard = document.AddWorksheet("Dashboard");
                ExcelModernChart chart = dashboard.AddModernChart(
                    new ExcelChartData(
                        new[] { "A", "B" },
                        new[] { new ExcelChartSeries("Value", new[] { 1d, 2d }) }),
                    1,
                    1,
                    ExcelModernChartType.Funnel);
                Assert.False(reserved.Hidden);
                Assert.Equal("user data", reserved.CellAt(1, 1).GetValue<string>());
                Assert.NotEqual(reserved.Name, chart.DataRange!.SheetName);
                Assert.True(document[chart.DataRange.SheetName].Hidden);
                document.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelModernChart chart = Assert.Single(document["Dashboard"].ModernCharts);
                Assert.NotNull(chart.DataRange);
                chart.UpdateData(new ExcelChartData(
                    new[] { "Changed" },
                    new[] { new ExcelChartSeries("Value", new[] { 3d }) }));
                Assert.Equal("user data", document["OfficeIMO_ChartData"].CellAt(1, 1).GetValue<string>());
                Assert.False(document["OfficeIMO_ChartData"].Hidden);
                Assert.Empty(document.ValidateOpenXml());
            }
        }
    }
}
