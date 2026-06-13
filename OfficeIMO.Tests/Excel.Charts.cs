using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel;
using Xunit;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public partial class Excel {
        private const int DefaultChartFontSize = 1100;

        private static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument spreadsheet, string sheetName) {
            var workbookPart = spreadsheet.WorkbookPart!;
            var sheet = workbookPart.Workbook.Sheets!
                .OfType<Sheet>()
                .First(sheet => string.Equals(sheet.Name?.Value, sheetName, StringComparison.Ordinal));
            return (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
        }

        private static WorksheetPart GetWorksheetPartWithCharts(SpreadsheetDocument spreadsheet) {
            var worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts
                .FirstOrDefault(part => part.DrawingsPart?.ChartParts.Any() == true);
            Assert.NotNull(worksheetPart);
            return worksheetPart!;
        }

        [Fact]
        public void Test_ExcelChartData_From_ReadOnlyListUsesIndexerWithoutSnapshotEnumeration() {
            var items = new ThrowOnEnumerateReadOnlyList<ChartProjectionRow>(
                new ChartProjectionRow("Q1", 10d),
                new ChartProjectionRow("Q2", 20d));

            var data = ExcelChartData.From(
                items,
                row => row.Category,
                new ExcelChartSeriesDefinition<ChartProjectionRow>("Sales", row => row.Value));

            Assert.Equal(new[] { "Q1", "Q2" }, data.Categories);
            Assert.Equal(new[] { 10d, 20d }, data.Series[0].Values);
        }

        [Fact]
        public void Test_ExcelChartData_From_StreamsEnumerableOnce() {
            var items = new SinglePassChartRows(
                new ChartProjectionRow("Q1", 10d),
                new ChartProjectionRow("Q2", 20d));

            var data = ExcelChartData.From(
                items,
                row => row.Category,
                new ExcelChartSeriesDefinition<ChartProjectionRow>("Sales", row => row.Value),
                new ExcelChartSeriesDefinition<ChartProjectionRow>("Target", row => row.Value + 5d));

            Assert.Equal(1, items.EnumerationCount);
            Assert.Equal(new[] { "Q1", "Q2" }, data.Categories);
            Assert.Equal(new[] { 10d, 20d }, data.Series[0].Values);
            Assert.Equal(new[] { 15d, 25d }, data.Series[1].Values);
        }

        [Fact]
        public void Test_ExcelChartData_PublicConstructorsSnapshotMutableInputs() {
            var categories = new List<string> { "Q1", "Q2" };
            var values = new List<double> { 10d, 20d };
            var sourceSeries = new List<ExcelChartSeries> {
                new ExcelChartSeries("Sales", values)
            };

            var data = new ExcelChartData(categories, sourceSeries);
            categories[0] = "Changed";
            values[0] = 99d;
            sourceSeries.Clear();

            Assert.Equal(new[] { "Q1", "Q2" }, data.Categories);
            Assert.Single(data.Series);
            Assert.Equal(new[] { 10d, 20d }, data.Series[0].Values);
        }

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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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

        private sealed class ChartProjectionRow {
            public ChartProjectionRow(string category, double value) {
                Category = category;
                Value = value;
            }

            public string Category { get; }

            public double Value { get; }
        }

        private sealed class SinglePassChartRows : System.Collections.Generic.IEnumerable<ChartProjectionRow> {
            private readonly ChartProjectionRow[] _items;

            public SinglePassChartRows(params ChartProjectionRow[] items) {
                _items = items;
            }

            public int EnumerationCount { get; private set; }

            public System.Collections.Generic.IEnumerator<ChartProjectionRow> GetEnumerator() {
                EnumerationCount++;
                if (EnumerationCount > 1) {
                    throw new InvalidOperationException("Chart projection should stream source rows once.");
                }

                return ((System.Collections.Generic.IEnumerable<ChartProjectionRow>)_items).GetEnumerator();
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }

        [Fact]
        public void Test_ExcelCharts_FluentRangeBuilder_CreatesChart() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.FluentRange.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                sheet.CellValue(1, 1, "Quarter");
                sheet.CellValue(1, 2, "Sales");
                sheet.CellValue(1, 3, "Target");
                sheet.CellValue(2, 1, "Q1");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(2, 3, 12);
                sheet.CellValue(3, 1, "Q2");
                sheet.CellValue(3, 2, 20);
                sheet.CellValue(3, 3, 22);

                sheet.Chart("A1:C3")
                    .Line()
                    .Title("Quarterly")
                    .Size(480, 320)
                    .At(1, 6);

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartPart = wsPart.DrawingsPart!.ChartParts.Single();
                var chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
                Assert.Equal("Quarterly", chart.Title!.Descendants<A.Text>().First().Text);
                Assert.NotNull(chart.PlotArea!.GetFirstChild<C.LineChart>());

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }
        }

        [Fact]
        public void Test_ExcelCharts_RadarChart_CanCreateAndUpdate() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Radar.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Coverage");
                sheet.CellValue(1, 1, "Area");
                sheet.CellValue(1, 2, "Current");
                sheet.CellValue(1, 3, "Target");
                sheet.CellValue(2, 1, "Mail");
                sheet.CellValue(2, 2, 3);
                sheet.CellValue(2, 3, 5);
                sheet.CellValue(3, 1, "DNS");
                sheet.CellValue(3, 2, 4);
                sheet.CellValue(3, 3, 5);
                sheet.CellValue(4, 1, "Identity");
                sheet.CellValue(4, 2, 2);
                sheet.CellValue(4, 3, 5);

                var chart = sheet.Chart("A1:C4")
                    .Radar()
                    .Title("Coverage")
                    .Size(520, 340)
                    .At(1, 5);

                chart.UpdateData(new ExcelChartData(
                    new[] { "Mail", "DNS", "Identity", "Device" },
                    new[] {
                        new ExcelChartSeries("Current", new[] { 4d, 3d, 5d, 2d }, ExcelChartType.Radar),
                        new ExcelChartSeries("Target", new[] { 5d, 5d, 5d, 4d }, ExcelChartType.Radar)
                    }));
                chart.SetSeriesLineColor("Current", "2563EB", widthPoints: 1.25)
                    .SetSeriesMarker("Target", C.MarkerStyleValues.Circle, size: 6, fillColor: "F97316", lineColor: "7C2D12");

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartPart = wsPart.DrawingsPart!.ChartParts.Single();
                var chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
                var plotArea = chart.GetFirstChild<C.PlotArea>()!;
                var radarChart = plotArea.GetFirstChild<C.RadarChart>();

                Assert.NotNull(radarChart);
                Assert.Equal(C.RadarStyleValues.Standard, radarChart!.RadarStyle?.Val?.Value);
                Assert.Equal("Coverage", chart.Title!.Descendants<A.Text>().First().Text);
                Assert.Equal(2, radarChart.Elements<C.RadarChartSeries>().Count());
                Assert.Equal(2, radarChart.Elements<C.AxisId>().Count());
                Assert.NotNull(plotArea.GetFirstChild<C.CategoryAxis>());
                Assert.NotNull(plotArea.GetFirstChild<C.ValueAxis>());

                var series = radarChart.Elements<C.RadarChartSeries>().First();
                var valuesCache = series.GetFirstChild<C.Values>()!
                    .GetFirstChild<C.NumberReference>()!
                    .GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("2563EB", series.OuterXml, StringComparison.OrdinalIgnoreCase);

                var targetSeries = radarChart.Elements<C.RadarChartSeries>().Skip(1).First();
                Assert.Equal(C.MarkerStyleValues.Circle, targetSeries.GetFirstChild<C.Marker>()?.Symbol?.Val?.Value);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_StockChart_CanCreateAndUpdate() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Stock.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Prices");
                sheet.CellValue(1, 1, "Date");
                sheet.CellValue(1, 2, "Open");
                sheet.CellValue(1, 3, "High");
                sheet.CellValue(1, 4, "Low");
                sheet.CellValue(1, 5, "Close");
                sheet.CellValue(2, 1, "2026-01-01");
                sheet.CellValue(2, 2, 102);
                sheet.CellValue(2, 3, 108);
                sheet.CellValue(2, 4, 99);
                sheet.CellValue(2, 5, 105);
                sheet.CellValue(3, 1, "2026-01-02");
                sheet.CellValue(3, 2, 105);
                sheet.CellValue(3, 3, 111);
                sheet.CellValue(3, 4, 101);
                sheet.CellValue(3, 5, 109);
                sheet.CellValue(4, 1, "2026-01-03");
                sheet.CellValue(4, 2, 109);
                sheet.CellValue(4, 3, 116);
                sheet.CellValue(4, 4, 106);
                sheet.CellValue(4, 5, 113);

                var chart = sheet.Chart("A1:E4")
                    .Stock()
                    .Title("Price Range")
                    .Size(560, 340)
                    .At(1, 7);

                chart.UpdateData(new ExcelChartData(
                    new[] { "2026-01-01", "2026-01-02", "2026-01-03", "2026-01-04" },
                    new[] {
                        new ExcelChartSeries("Open", new[] { 102d, 105d, 109d, 112d }, ExcelChartType.Stock),
                        new ExcelChartSeries("High", new[] { 108d, 111d, 116d, 118d }, ExcelChartType.Stock),
                        new ExcelChartSeries("Low", new[] { 99d, 101d, 106d, 110d }, ExcelChartType.Stock),
                        new ExcelChartSeries("Close", new[] { 105d, 109d, 113d, 116d }, ExcelChartType.Stock)
                    }));
                chart.SetSeriesLineColor("Close", "2563EB", widthPoints: 1.25);

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartPart = wsPart.DrawingsPart!.ChartParts.Single();
                var chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
                var plotArea = chart.GetFirstChild<C.PlotArea>()!;
                var stockChart = plotArea.GetFirstChild<C.StockChart>();

                Assert.NotNull(stockChart);
                Assert.Equal("Price Range", chart.Title!.Descendants<A.Text>().First().Text);
                Assert.Equal(4, stockChart!.Elements<C.LineChartSeries>().Count());
                Assert.NotNull(stockChart.GetFirstChild<C.HighLowLines>());
                Assert.NotNull(stockChart.GetFirstChild<C.UpDownBars>());
                Assert.Equal(2, stockChart.Elements<C.AxisId>().Count());
                Assert.NotNull(plotArea.GetFirstChild<C.CategoryAxis>());
                Assert.NotNull(plotArea.GetFirstChild<C.ValueAxis>());

                var closeSeries = stockChart.Elements<C.LineChartSeries>().Last();
                var valuesCache = closeSeries.GetFirstChild<C.Values>()!
                    .GetFirstChild<C.NumberReference>()!
                    .GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("2563EB", closeSeries.OuterXml, StringComparison.OrdinalIgnoreCase);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_SurfaceChart_CanCreateAndUpdate() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Surface.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Terrain");
                sheet.CellValue(1, 1, "Zone");
                sheet.CellValue(1, 2, "Low");
                sheet.CellValue(1, 3, "Mid");
                sheet.CellValue(1, 4, "High");
                sheet.CellValue(2, 1, "North");
                sheet.CellValue(2, 2, 12);
                sheet.CellValue(2, 3, 18);
                sheet.CellValue(2, 4, 24);
                sheet.CellValue(3, 1, "Central");
                sheet.CellValue(3, 2, 16);
                sheet.CellValue(3, 3, 21);
                sheet.CellValue(3, 4, 29);
                sheet.CellValue(4, 1, "South");
                sheet.CellValue(4, 2, 14);
                sheet.CellValue(4, 3, 20);
                sheet.CellValue(4, 4, 27);

                var chart = sheet.Chart("A1:D4")
                    .Surface()
                    .Title("Surface")
                    .Size(560, 340)
                    .At(1, 6);

                chart.UpdateData(new ExcelChartData(
                    new[] { "North", "Central", "South", "West" },
                    new[] {
                        new ExcelChartSeries("Low", new[] { 11d, 15d, 13d, 17d }, ExcelChartType.Surface),
                        new ExcelChartSeries("Mid", new[] { 18d, 22d, 19d, 23d }, ExcelChartType.Surface),
                        new ExcelChartSeries("High", new[] { 25d, 30d, 28d, 32d }, ExcelChartType.Surface)
                    }));
                chart.SetSeriesFillColor("High", "22C55E");

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartPart = wsPart.DrawingsPart!.ChartParts.Single();
                var chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
                var plotArea = chart.GetFirstChild<C.PlotArea>()!;
                var surfaceChart = plotArea.GetFirstChild<C.Surface3DChart>();

                Assert.NotNull(surfaceChart);
                Assert.Equal("Surface", chart.Title!.Descendants<A.Text>().First().Text);
                Assert.Equal(3, surfaceChart!.Elements<C.SurfaceChartSeries>().Count());
                Assert.Equal(3, surfaceChart.Elements<C.AxisId>().Count());
                Assert.NotNull(plotArea.GetFirstChild<C.CategoryAxis>());
                Assert.NotNull(plotArea.GetFirstChild<C.ValueAxis>());
                Assert.NotNull(plotArea.GetFirstChild<C.SeriesAxis>());

                var highSeries = surfaceChart.Elements<C.SurfaceChartSeries>().Last();
                var valuesCache = highSeries.GetFirstChild<C.Values>()!
                    .GetFirstChild<C.NumberReference>()!
                    .GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("22C55E", highSeries.OuterXml, StringComparison.OrdinalIgnoreCase);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_SurfaceVariantCharts_CanCreateAndUpdate() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SurfaceVariants.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Surface");
                sheet.CellValue(1, 1, "Bucket");
                sheet.CellValue(1, 2, "Low");
                sheet.CellValue(1, 3, "Mid");
                sheet.CellValue(1, 4, "High");
                sheet.CellValue(2, 1, "North");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(2, 3, 18);
                sheet.CellValue(2, 4, 25);
                sheet.CellValue(3, 1, "East");
                sheet.CellValue(3, 2, 14);
                sheet.CellValue(3, 3, 22);
                sheet.CellValue(3, 4, 30);
                sheet.CellValue(4, 1, "South");
                sheet.CellValue(4, 2, 12);
                sheet.CellValue(4, 3, 20);
                sheet.CellValue(4, 4, 28);

                var wireframeChart = sheet.Chart("A1:D4")
                    .SurfaceWireframe()
                    .Title("Wireframe Surface")
                    .Size(520, 340)
                    .At(1, 6);

                wireframeChart.UpdateData(new ExcelChartData(
                    new[] { "North", "East", "South", "West" },
                    new[] {
                        new ExcelChartSeries("Low", new[] { 11d, 15d, 13d, 17d }, ExcelChartType.SurfaceWireframe),
                        new ExcelChartSeries("Mid", new[] { 18d, 22d, 19d, 23d }, ExcelChartType.SurfaceWireframe),
                        new ExcelChartSeries("High", new[] { 25d, 30d, 28d, 32d }, ExcelChartType.SurfaceWireframe)
                    }));
                wireframeChart.SetSeriesFillColor("High", "22C55E");

                sheet.Chart("A1:D4")
                    .SurfaceContour()
                    .Title("Contour Surface")
                    .Size(520, 340)
                    .At(18, 6);

                sheet.Chart("A1:D4")
                    .SurfaceContourWireframe()
                    .Title("Wireframe Contour")
                    .Size(520, 340)
                    .At(35, 6);

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartParts = wsPart.DrawingsPart!.ChartParts.ToList();
                Assert.Equal(3, chartParts.Count);

                var plotAreas = chartParts
                    .Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!)
                    .ToList();

                var surface3DChart = plotAreas
                    .Select(plotArea => plotArea.GetFirstChild<C.Surface3DChart>())
                    .Single(chart => chart != null);
                var contourCharts = plotAreas
                    .Select(plotArea => plotArea.GetFirstChild<C.SurfaceChart>())
                    .Where(chart => chart != null)
                    .ToList();

                Assert.NotNull(surface3DChart);
                Assert.True(surface3DChart!.GetFirstChild<C.Wireframe>()!.Val!.Value);
                Assert.Equal(3, surface3DChart.Elements<C.SurfaceChartSeries>().Count());
                Assert.Equal(3, surface3DChart.Elements<C.AxisId>().Count());
                Assert.NotNull(plotAreas.First(plotArea => plotArea.GetFirstChild<C.Surface3DChart>() == surface3DChart).GetFirstChild<C.SeriesAxis>());

                var highSeries = surface3DChart.Elements<C.SurfaceChartSeries>().Last();
                var valuesCache = highSeries.GetFirstChild<C.Values>()!
                    .GetFirstChild<C.NumberReference>()!
                    .GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("22C55E", highSeries.OuterXml, StringComparison.OrdinalIgnoreCase);

                Assert.Equal(2, contourCharts.Count);
                Assert.Contains(contourCharts, chart => chart!.GetFirstChild<C.Wireframe>()!.Val!.Value == false);
                Assert.Contains(contourCharts, chart => chart!.GetFirstChild<C.Wireframe>()!.Val!.Value);
                foreach (var contourChart in contourCharts) {
                    Assert.Equal(3, contourChart!.Elements<C.SurfaceChartSeries>().Count());
                    Assert.Equal(3, contourChart.Elements<C.AxisId>().Count());
                    Assert.NotNull(plotAreas.First(plotArea => plotArea.GetFirstChild<C.SurfaceChart>() == contourChart).GetFirstChild<C.SeriesAxis>());
                }

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_Pie3DChart_CanCreateAndUpdate() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Pie3D.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Mix");
                sheet.CellValue(1, 1, "Category");
                sheet.CellValue(1, 2, "Share");
                sheet.CellValue(2, 1, "Mail");
                sheet.CellValue(2, 2, 42);
                sheet.CellValue(3, 1, "DNS");
                sheet.CellValue(3, 2, 28);
                sheet.CellValue(4, 1, "Identity");
                sheet.CellValue(4, 2, 30);

                var chart = sheet.Chart("A1:B4")
                    .Pie3D()
                    .Title("Workload Mix")
                    .Size(520, 340)
                    .At(1, 5);

                chart.UpdateData(new ExcelChartData(
                    new[] { "Mail", "DNS", "Identity", "Device" },
                    new[] {
                        new ExcelChartSeries("Share", new[] { 40d, 24d, 26d, 10d }, ExcelChartType.Pie3D)
                    }));
                chart.SetDataLabels(showValue: true, showCategoryName: true, showSeriesName: false, showLegendKey: false, showPercent: false)
                    .SetSeriesFillColor("Share", "2563EB");

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartPart = wsPart.DrawingsPart!.ChartParts.Single();
                var chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
                var plotArea = chart.GetFirstChild<C.PlotArea>()!;
                var pieChart = plotArea.GetFirstChild<C.Pie3DChart>();

                Assert.NotNull(pieChart);
                Assert.Equal("Workload Mix", chart.Title!.Descendants<A.Text>().First().Text);
                Assert.True(pieChart!.VaryColors?.Val?.Value);

                var series = Assert.Single(pieChart.Elements<C.PieChartSeries>());
                var valuesCache = series.GetFirstChild<C.Values>()!
                    .GetFirstChild<C.NumberReference>()!
                    .GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("2563EB", series.OuterXml, StringComparison.OrdinalIgnoreCase);

                var labels = pieChart.GetFirstChild<C.DataLabels>();
                Assert.True(labels!.GetFirstChild<C.ShowValue>()!.Val!.Value);
                Assert.True(labels.GetFirstChild<C.ShowCategoryName>()!.Val!.Value);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_OfPieCharts_CanCreateAndUpdate() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.OfPie.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Share");
                sheet.CellValue(1, 1, "Segment");
                sheet.CellValue(1, 2, "Share");
                sheet.CellValue(2, 1, "Core");
                sheet.CellValue(2, 2, 40);
                sheet.CellValue(3, 1, "Growth");
                sheet.CellValue(3, 2, 24);
                sheet.CellValue(4, 1, "Services");
                sheet.CellValue(4, 2, 26);
                sheet.CellValue(5, 1, "Other");
                sheet.CellValue(5, 2, 10);

                var pieChart = sheet.Chart("A1:B5")
                    .PieOfPie()
                    .Title("Pie of Pie")
                    .Size(520, 340)
                    .At(1, 5);

                pieChart.UpdateData(new ExcelChartData(
                    new[] { "Core", "Growth", "Services", "Other", "Legacy" },
                    new[] {
                        new ExcelChartSeries("Share", new[] { 38d, 25d, 24d, 9d, 4d }, ExcelChartType.PieOfPie)
                    }));
                pieChart.SetDataLabels(showValue: true, showCategoryName: true, showSeriesName: false, showLegendKey: false, showPercent: false)
                    .SetSeriesFillColor("Share", "2563EB");

                sheet.Chart("A1:B5")
                    .BarOfPie()
                    .Title("Bar of Pie")
                    .Size(520, 340)
                    .At(18, 5);

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartParts = wsPart.DrawingsPart!.ChartParts.ToList();
                Assert.Equal(2, chartParts.Count);

                var ofPieCharts = chartParts
                    .Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!.GetFirstChild<C.OfPieChart>())
                    .ToList();

                var pieChart = ofPieCharts.First(chart => chart?.GetFirstChild<C.OfPieType>()?.Val?.Value == C.OfPieValues.Pie);
                var barChart = ofPieCharts.First(chart => chart?.GetFirstChild<C.OfPieType>()?.Val?.Value == C.OfPieValues.Bar);

                Assert.NotNull(pieChart);
                Assert.Single(pieChart!.Elements<C.PieChartSeries>());
                Assert.NotNull(pieChart.GetFirstChild<C.SeriesLines>());
                Assert.Equal((ushort)75, pieChart.GetFirstChild<C.SecondPieSize>()!.Val!.Value);
                Assert.Equal(C.SplitValues.Position, pieChart.GetFirstChild<C.SplitType>()!.Val!.Value);

                var series = Assert.Single(pieChart.Elements<C.PieChartSeries>());
                var valuesCache = series.GetFirstChild<C.Values>()!
                    .GetFirstChild<C.NumberReference>()!
                    .GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)5, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("2563EB", series.OuterXml, StringComparison.OrdinalIgnoreCase);

                var labels = pieChart.GetFirstChild<C.DataLabels>();
                Assert.True(labels!.GetFirstChild<C.ShowValue>()!.Val!.Value);
                Assert.True(labels.GetFirstChild<C.ShowCategoryName>()!.Val!.Value);

                Assert.NotNull(barChart);
                Assert.Single(barChart!.Elements<C.PieChartSeries>());
                Assert.NotNull(barChart.GetFirstChild<C.SeriesLines>());

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_BarStacked100Variants_CanCreateAndUpdate() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.BarStacked100.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Mix");
                sheet.CellValue(1, 1, "Quarter");
                sheet.CellValue(1, 2, "Sales");
                sheet.CellValue(1, 3, "Target");
                sheet.CellValue(2, 1, "Q1");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(2, 3, 12);
                sheet.CellValue(3, 1, "Q2");
                sheet.CellValue(3, 2, 18);
                sheet.CellValue(3, 3, 20);
                sheet.CellValue(4, 1, "Q3");
                sheet.CellValue(4, 2, 24);
                sheet.CellValue(4, 3, 26);

                var columnChart = sheet.Chart("A1:C4")
                    .ColumnStacked100()
                    .Title("100% Stacked Columns")
                    .Size(520, 340)
                    .At(1, 5);

                columnChart.UpdateData(new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3", "Q4" },
                    new[] {
                        new ExcelChartSeries("Sales", new[] { 12d, 19d, 25d, 31d }, ExcelChartType.ColumnStacked100),
                        new ExcelChartSeries("Target", new[] { 14d, 21d, 28d, 34d }, ExcelChartType.ColumnStacked100)
                    }));
                columnChart.SetSeriesFillColor("Sales", "2563EB")
                    .SetDataLabels(showValue: true);

                sheet.Chart("A1:C4")
                    .BarStacked100()
                    .Title("100% Stacked Bars")
                    .Size(520, 340)
                    .At(18, 5);

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartParts = wsPart.DrawingsPart!.ChartParts.ToList();
                Assert.Equal(2, chartParts.Count);

                var barCharts = chartParts
                    .Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!.GetFirstChild<C.BarChart>())
                    .ToList();

                var columnChart = barCharts.First(chart => chart?.BarDirection?.Val?.Value == C.BarDirectionValues.Column);
                var barChart = barCharts.First(chart => chart?.BarDirection?.Val?.Value == C.BarDirectionValues.Bar);

                Assert.NotNull(columnChart);
                Assert.Equal(C.BarGroupingValues.PercentStacked, columnChart!.BarGrouping?.Val?.Value);
                Assert.Equal(2, columnChart.Elements<C.BarChartSeries>().Count());
                Assert.Equal(2, columnChart.Elements<C.AxisId>().Count());

                var series = columnChart.Elements<C.BarChartSeries>().First();
                var valuesCache = series.GetFirstChild<C.Values>()!
                    .GetFirstChild<C.NumberReference>()!
                    .GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("2563EB", series.OuterXml, StringComparison.OrdinalIgnoreCase);
                Assert.True(columnChart.GetFirstChild<C.DataLabels>()!.GetFirstChild<C.ShowValue>()!.Val!.Value);

                Assert.NotNull(barChart);
                Assert.Equal(C.BarGroupingValues.PercentStacked, barChart!.BarGrouping?.Val?.Value);
                Assert.Equal(2, barChart.Elements<C.AxisId>().Count());

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_Bar3DCharts_CanCreateAndUpdate() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Bar3D.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Sales");
                sheet.CellValue(1, 1, "Quarter");
                sheet.CellValue(1, 2, "Sales");
                sheet.CellValue(1, 3, "Target");
                sheet.CellValue(2, 1, "Q1");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(2, 3, 12);
                sheet.CellValue(3, 1, "Q2");
                sheet.CellValue(3, 2, 18);
                sheet.CellValue(3, 3, 20);
                sheet.CellValue(4, 1, "Q3");
                sheet.CellValue(4, 2, 24);
                sheet.CellValue(4, 3, 26);

                var columnChart = sheet.Chart("A1:C4")
                    .Column3DClustered()
                    .Title("3D Columns")
                    .Size(520, 340)
                    .At(1, 5);

                columnChart.UpdateData(new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3", "Q4" },
                    new[] {
                        new ExcelChartSeries("Sales", new[] { 12d, 19d, 25d, 31d }, ExcelChartType.Column3DClustered),
                        new ExcelChartSeries("Target", new[] { 14d, 21d, 28d, 34d }, ExcelChartType.Column3DClustered)
                    }));
                columnChart.SetSeriesFillColor("Sales", "2563EB")
                    .SetDataLabels(showValue: true);

                sheet.Chart("A1:C4")
                    .Bar3DStacked()
                    .Title("3D Bars")
                    .Size(520, 340)
                    .At(18, 5);

                sheet.Chart("A1:C4")
                    .Column3DStacked100()
                    .Title("3D 100% Columns")
                    .Size(520, 340)
                    .At(35, 5);

                sheet.Chart("A1:C4")
                    .Bar3DStacked100()
                    .Title("3D 100% Bars")
                    .Size(520, 340)
                    .At(52, 5);

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartParts = wsPart.DrawingsPart!.ChartParts.ToList();
                Assert.Equal(4, chartParts.Count);

                var plotAreas = chartParts
                    .Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!)
                    .ToList();

                var columnChart = plotAreas
                    .Select(plotArea => plotArea.GetFirstChild<C.Bar3DChart>())
                    .First(chart => chart?.BarDirection?.Val?.Value == C.BarDirectionValues.Column
                        && chart.BarGrouping?.Val?.Value == C.BarGroupingValues.Clustered);
                var barChart = plotAreas
                    .Select(plotArea => plotArea.GetFirstChild<C.Bar3DChart>())
                    .First(chart => chart?.BarDirection?.Val?.Value == C.BarDirectionValues.Bar
                        && chart.BarGrouping?.Val?.Value == C.BarGroupingValues.Stacked);
                var percentColumnChart = plotAreas
                    .Select(plotArea => plotArea.GetFirstChild<C.Bar3DChart>())
                    .First(chart => chart?.BarDirection?.Val?.Value == C.BarDirectionValues.Column
                        && chart.BarGrouping?.Val?.Value == C.BarGroupingValues.PercentStacked);
                var percentBarChart = plotAreas
                    .Select(plotArea => plotArea.GetFirstChild<C.Bar3DChart>())
                    .First(chart => chart?.BarDirection?.Val?.Value == C.BarDirectionValues.Bar
                        && chart.BarGrouping?.Val?.Value == C.BarGroupingValues.PercentStacked);

                Assert.NotNull(columnChart);
                Assert.Equal(C.BarGroupingValues.Clustered, columnChart!.BarGrouping?.Val?.Value);
                Assert.Equal(2, columnChart.Elements<C.BarChartSeries>().Count());
                Assert.Equal(3, columnChart.Elements<C.AxisId>().Count());
                Assert.Equal(C.ShapeValues.Box, columnChart.GetFirstChild<C.Shape>()?.Val?.Value);
                Assert.NotNull(plotAreas.First(plotArea => plotArea.GetFirstChild<C.Bar3DChart>() == columnChart).GetFirstChild<C.SeriesAxis>());

                var series = columnChart.Elements<C.BarChartSeries>().First();
                var valuesCache = series.GetFirstChild<C.Values>()!
                    .GetFirstChild<C.NumberReference>()!
                    .GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("2563EB", series.OuterXml, StringComparison.OrdinalIgnoreCase);
                Assert.True(columnChart.GetFirstChild<C.DataLabels>()!.GetFirstChild<C.ShowValue>()!.Val!.Value);

                Assert.NotNull(barChart);
                Assert.Equal(C.BarGroupingValues.Stacked, barChart!.BarGrouping?.Val?.Value);
                Assert.Equal(3, barChart.Elements<C.AxisId>().Count());

                Assert.NotNull(percentColumnChart);
                Assert.Equal(3, percentColumnChart!.Elements<C.AxisId>().Count());
                Assert.NotNull(plotAreas.First(plotArea => plotArea.GetFirstChild<C.Bar3DChart>() == percentColumnChart).GetFirstChild<C.SeriesAxis>());

                Assert.NotNull(percentBarChart);
                Assert.Equal(3, percentBarChart!.Elements<C.AxisId>().Count());
                Assert.NotNull(plotAreas.First(plotArea => plotArea.GetFirstChild<C.Bar3DChart>() == percentBarChart).GetFirstChild<C.SeriesAxis>());

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_LineStackedVariants_CanCreateAndUpdate() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.LineStacked.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Trend");
                sheet.CellValue(1, 1, "Month");
                sheet.CellValue(1, 2, "Actual");
                sheet.CellValue(1, 3, "Plan");
                sheet.CellValue(2, 1, "Jan");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(2, 3, 12);
                sheet.CellValue(3, 1, "Feb");
                sheet.CellValue(3, 2, 15);
                sheet.CellValue(3, 3, 17);
                sheet.CellValue(4, 1, "Mar");
                sheet.CellValue(4, 2, 20);
                sheet.CellValue(4, 3, 23);

                var stackedChart = sheet.Chart("A1:C4")
                    .LineStacked()
                    .Title("Stacked Line")
                    .Size(520, 340)
                    .At(1, 5);

                stackedChart.UpdateData(new ExcelChartData(
                    new[] { "Jan", "Feb", "Mar", "Apr" },
                    new[] {
                        new ExcelChartSeries("Actual", new[] { 11d, 16d, 22d, 27d }, ExcelChartType.LineStacked),
                        new ExcelChartSeries("Plan", new[] { 12d, 18d, 24d, 29d }, ExcelChartType.LineStacked)
                    }));
                stackedChart.SetSeriesLineColor("Actual", "2563EB", widthPoints: 1.25)
                    .SetSeriesMarker("Plan", C.MarkerStyleValues.Circle, size: 6, fillColor: "F97316", lineColor: "7C2D12")
                    .SetDataLabels(showValue: true);

                sheet.Chart("A1:C4")
                    .LineStacked100()
                    .Title("100% Stacked Line")
                    .Size(520, 340)
                    .At(18, 5);

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartParts = wsPart.DrawingsPart!.ChartParts.ToList();
                Assert.Equal(2, chartParts.Count);

                var lineCharts = chartParts
                    .Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!.GetFirstChild<C.LineChart>())
                    .ToList();

                var stackedChart = lineCharts.First(chart => chart?.Grouping?.Val?.Value == C.GroupingValues.Stacked);
                var percentChart = lineCharts.First(chart => chart?.Grouping?.Val?.Value == C.GroupingValues.PercentStacked);

                Assert.NotNull(stackedChart);
                Assert.Equal(2, stackedChart!.Elements<C.LineChartSeries>().Count());
                Assert.Equal(2, stackedChart.Elements<C.AxisId>().Count());

                var series = stackedChart.Elements<C.LineChartSeries>().First();
                var valuesCache = series.GetFirstChild<C.Values>()!
                    .GetFirstChild<C.NumberReference>()!
                    .GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("2563EB", series.OuterXml, StringComparison.OrdinalIgnoreCase);
                Assert.True(stackedChart.GetFirstChild<C.DataLabels>()!.GetFirstChild<C.ShowValue>()!.Val!.Value);

                var planSeries = stackedChart.Elements<C.LineChartSeries>().Skip(1).First();
                Assert.Equal(C.MarkerStyleValues.Circle, planSeries.GetFirstChild<C.Marker>()?.Symbol?.Val?.Value);

                Assert.NotNull(percentChart);
                Assert.Equal(2, percentChart!.Elements<C.LineChartSeries>().Count());
                Assert.Equal(2, percentChart.Elements<C.AxisId>().Count());

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Theory]
        [InlineData(ExcelChartType.Pie3D, "3-D pie")]
        [InlineData(ExcelChartType.PieOfPie, "pie-of-pie")]
        [InlineData(ExcelChartType.BarOfPie, "bar-of-pie")]
        public void Test_ExcelCharts_SingleSeriesPieVariantsRejectMultipleSeries(ExcelChartType type, string chartName) {
            string filePath = Path.Combine(_directoryWithFiles, $"ExcelCharts.{type}.MultipleSeries.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Share");
                sheet.CellValue(1, 1, "Segment");
                sheet.CellValue(1, 2, "Current");
                sheet.CellValue(1, 3, "Previous");
                sheet.CellValue(2, 1, "Core");
                sheet.CellValue(2, 2, 40);
                sheet.CellValue(2, 3, 35);
                sheet.CellValue(3, 1, "Growth");
                sheet.CellValue(3, 2, 24);
                sheet.CellValue(3, 3, 20);
                sheet.CellValue(4, 1, "Services");
                sheet.CellValue(4, 2, 26);
                sheet.CellValue(4, 3, 30);

                var exception = Assert.Throws<NotSupportedException>(() =>
                    sheet.AddChartFromRange("A1:C4", row: 1, column: 5, type: type));
                Assert.Contains(chartName, exception.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("exactly one series", exception.Message, StringComparison.OrdinalIgnoreCase);
            }
        }

        [Fact]
        public void Test_ExcelCharts_Line3DChart_CanCreateAndUpdate() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Line3D.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Trend");
                sheet.CellValue(1, 1, "Month");
                sheet.CellValue(1, 2, "Actual");
                sheet.CellValue(1, 3, "Plan");
                sheet.CellValue(2, 1, "Jan");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(2, 3, 12);
                sheet.CellValue(3, 1, "Feb");
                sheet.CellValue(3, 2, 15);
                sheet.CellValue(3, 3, 17);
                sheet.CellValue(4, 1, "Mar");
                sheet.CellValue(4, 2, 20);
                sheet.CellValue(4, 3, 23);

                var chart = sheet.Chart("A1:C4")
                    .Line3D()
                    .Title("3D Line")
                    .Size(520, 340)
                    .At(1, 5);

                chart.UpdateData(new ExcelChartData(
                    new[] { "Jan", "Feb", "Mar", "Apr" },
                    new[] {
                        new ExcelChartSeries("Actual", new[] { 11d, 16d, 22d, 27d }, ExcelChartType.Line3D),
                        new ExcelChartSeries("Plan", new[] { 12d, 18d, 24d, 29d }, ExcelChartType.Line3D)
                    }));
                chart.SetSeriesLineColor("Actual", "2563EB", widthPoints: 1.25)
                    .SetSeriesMarker("Plan", C.MarkerStyleValues.Circle, size: 6, fillColor: "F97316", lineColor: "7C2D12")
                    .SetDataLabels(showValue: true);

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartPart = wsPart.DrawingsPart!.ChartParts.Single();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
                var lineChart = plotArea.GetFirstChild<C.Line3DChart>();

                Assert.NotNull(lineChart);
                Assert.Equal(C.GroupingValues.Standard, lineChart!.Grouping?.Val?.Value);
                Assert.Equal(2, lineChart.Elements<C.LineChartSeries>().Count());
                Assert.Equal(3, lineChart.Elements<C.AxisId>().Count());
                Assert.NotNull(lineChart.GetFirstChild<C.GapDepth>());
                Assert.NotNull(plotArea.GetFirstChild<C.SeriesAxis>());

                var series = lineChart.Elements<C.LineChartSeries>().First();
                var valuesCache = series.GetFirstChild<C.Values>()!
                    .GetFirstChild<C.NumberReference>()!
                    .GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("2563EB", series.OuterXml, StringComparison.OrdinalIgnoreCase);
                Assert.True(lineChart.GetFirstChild<C.DataLabels>()!.GetFirstChild<C.ShowValue>()!.Val!.Value);

                var planSeries = lineChart.Elements<C.LineChartSeries>().Skip(1).First();
                Assert.Equal(C.MarkerStyleValues.Circle, planSeries.GetFirstChild<C.Marker>()?.Symbol?.Val?.Value);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_AreaStackedVariants_CanCreateAndUpdate() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.AreaStacked.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Mix");
                sheet.CellValue(1, 1, "Month");
                sheet.CellValue(1, 2, "Services");
                sheet.CellValue(1, 3, "Licenses");
                sheet.CellValue(2, 1, "Jan");
                sheet.CellValue(2, 2, 35);
                sheet.CellValue(2, 3, 18);
                sheet.CellValue(3, 1, "Feb");
                sheet.CellValue(3, 2, 42);
                sheet.CellValue(3, 3, 23);
                sheet.CellValue(4, 1, "Mar");
                sheet.CellValue(4, 2, 48);
                sheet.CellValue(4, 3, 29);

                var stackedChart = sheet.Chart("A1:C4")
                    .AreaStacked()
                    .Title("Stacked Area")
                    .Size(520, 340)
                    .At(1, 5);

                stackedChart.UpdateData(new ExcelChartData(
                    new[] { "Jan", "Feb", "Mar", "Apr" },
                    new[] {
                        new ExcelChartSeries("Services", new[] { 36d, 44d, 50d, 54d }, ExcelChartType.AreaStacked),
                        new ExcelChartSeries("Licenses", new[] { 19d, 25d, 31d, 34d }, ExcelChartType.AreaStacked)
                    }));
                stackedChart.SetSeriesFillColor("Services", "2563EB")
                    .SetDataLabels(showValue: true);

                sheet.Chart("A1:C4")
                    .AreaStacked100()
                    .Title("100% Stacked Area")
                    .Size(520, 340)
                    .At(18, 5);

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartParts = wsPart.DrawingsPart!.ChartParts.ToList();
                Assert.Equal(2, chartParts.Count);

                var areaCharts = chartParts
                    .Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!.GetFirstChild<C.AreaChart>())
                    .ToList();

                var stackedChart = areaCharts.First(chart => chart?.Grouping?.Val?.Value == C.GroupingValues.Stacked);
                var percentChart = areaCharts.First(chart => chart?.Grouping?.Val?.Value == C.GroupingValues.PercentStacked);

                Assert.NotNull(stackedChart);
                Assert.Equal(2, stackedChart!.Elements<C.AreaChartSeries>().Count());
                Assert.Equal(2, stackedChart.Elements<C.AxisId>().Count());

                var series = stackedChart.Elements<C.AreaChartSeries>().First();
                var valuesCache = series.GetFirstChild<C.Values>()!
                    .GetFirstChild<C.NumberReference>()!
                    .GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("2563EB", series.OuterXml, StringComparison.OrdinalIgnoreCase);
                Assert.True(stackedChart.GetFirstChild<C.DataLabels>()!.GetFirstChild<C.ShowValue>()!.Val!.Value);

                Assert.NotNull(percentChart);
                Assert.Equal(2, percentChart!.Elements<C.AreaChartSeries>().Count());
                Assert.Equal(2, percentChart.Elements<C.AxisId>().Count());

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_Area3DCharts_CanCreateAndUpdate() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Area3D.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Volume");
                sheet.CellValue(1, 1, "Month");
                sheet.CellValue(1, 2, "Actual");
                sheet.CellValue(1, 3, "Plan");
                sheet.CellValue(2, 1, "Jan");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(2, 3, 12);
                sheet.CellValue(3, 1, "Feb");
                sheet.CellValue(3, 2, 16);
                sheet.CellValue(3, 3, 18);
                sheet.CellValue(4, 1, "Mar");
                sheet.CellValue(4, 2, 21);
                sheet.CellValue(4, 3, 24);

                var areaChart = sheet.Chart("A1:C4")
                    .Area3D()
                    .Title("3D Area")
                    .Size(520, 340)
                    .At(1, 5);

                areaChart.UpdateData(new ExcelChartData(
                    new[] { "Jan", "Feb", "Mar", "Apr" },
                    new[] {
                        new ExcelChartSeries("Actual", new[] { 11d, 17d, 22d, 28d }, ExcelChartType.Area3D),
                        new ExcelChartSeries("Plan", new[] { 13d, 19d, 25d, 31d }, ExcelChartType.Area3D)
                    }));
                areaChart.SetSeriesFillColor("Actual", "2563EB")
                    .SetDataLabels(showValue: true);

                sheet.Chart("A1:C4")
                    .Area3DStacked()
                    .Title("3D Stacked Area")
                    .Size(520, 340)
                    .At(18, 5);

                sheet.Chart("A1:C4")
                    .Area3DStacked100()
                    .Title("3D 100% Stacked Area")
                    .Size(520, 340)
                    .At(35, 5);

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartParts = wsPart.DrawingsPart!.ChartParts.ToList();
                Assert.Equal(3, chartParts.Count);

                var plotAreas = chartParts
                    .Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!)
                    .ToList();

                var standardChart = plotAreas
                    .Select(plotArea => plotArea.GetFirstChild<C.Area3DChart>())
                    .First(chart => chart?.Grouping?.Val?.Value == C.GroupingValues.Standard);
                var stackedChart = plotAreas
                    .Select(plotArea => plotArea.GetFirstChild<C.Area3DChart>())
                    .First(chart => chart?.Grouping?.Val?.Value == C.GroupingValues.Stacked);
                var percentChart = plotAreas
                    .Select(plotArea => plotArea.GetFirstChild<C.Area3DChart>())
                    .First(chart => chart?.Grouping?.Val?.Value == C.GroupingValues.PercentStacked);

                Assert.NotNull(standardChart);
                Assert.Equal(2, standardChart!.Elements<C.AreaChartSeries>().Count());
                Assert.Equal(3, standardChart.Elements<C.AxisId>().Count());
                Assert.NotNull(standardChart.GetFirstChild<C.GapDepth>());
                Assert.NotNull(plotAreas.First(plotArea => plotArea.GetFirstChild<C.Area3DChart>() == standardChart).GetFirstChild<C.SeriesAxis>());

                var series = standardChart.Elements<C.AreaChartSeries>().First();
                var valuesCache = series.GetFirstChild<C.Values>()!
                    .GetFirstChild<C.NumberReference>()!
                    .GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("2563EB", series.OuterXml, StringComparison.OrdinalIgnoreCase);
                Assert.True(standardChart.GetFirstChild<C.DataLabels>()!.GetFirstChild<C.ShowValue>()!.Val!.Value);

                Assert.NotNull(stackedChart);
                Assert.Equal(3, stackedChart!.Elements<C.AxisId>().Count());

                Assert.NotNull(percentChart);
                Assert.Equal(3, percentChart!.Elements<C.AxisId>().Count());

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_RecipeHelpers_CreateExpectedChartTypes() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.RecipeHelpers.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Dashboard");
                sheet.CellValue(1, 1, "Month");
                sheet.CellValue(1, 2, "Revenue");
                sheet.CellValue(2, 1, "Jan");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(3, 1, "Feb");
                sheet.CellValue(3, 2, 16);
                sheet.CellValue(4, 1, "Mar");
                sheet.CellValue(4, 2, 13);
                sheet.AddTable("A1:B4", hasHeader: true, name: "RevenueData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium9);

                sheet.AddRevenueTrendChart("A1:B4", row: 1, column: 5);
                sheet.AddTopNBarChart("A1:B4", row: 18, column: 5, title: "Top Revenue");
                sheet.AddVarianceColumnChart("A1:B4", row: 35, column: 5, title: "Revenue Variance");
                sheet.AddKpiScorecardChart("A1:B4", row: 52, column: 5, title: "Revenue KPI");
                sheet.AddContributionChart("A1:B4", row: 69, column: 5, title: "Revenue Mix");
                sheet.ChartFromTable("RevenueData")
                    .StatusBreakdown("Revenue Mix")
                    .At(86, 5);
                sheet.ChartFromTable("RevenueData")
                    .VarianceWaterfall("Revenue Bridge")
                    .At(103, 5);

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var charts = wsPart.DrawingsPart!.ChartParts
                    .Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!)
                    .ToList();
                var plotAreas = charts.Select(chart => chart.GetFirstChild<C.PlotArea>()!).ToList();

                Assert.Equal(7, plotAreas.Count);
                Assert.Contains(plotAreas, plotArea => plotArea.GetFirstChild<C.LineChart>() != null);
                Assert.Contains(plotAreas, plotArea => plotArea.GetFirstChild<C.DoughnutChart>() != null);
                Assert.Contains(plotAreas, plotArea => plotArea.Elements<C.BarChart>()
                    .Any(chart => chart.BarDirection?.Val?.Value == C.BarDirectionValues.Bar));
                Assert.Contains(plotAreas, plotArea => plotArea.Elements<C.BarChart>()
                    .Any(chart => chart.BarDirection?.Val?.Value == C.BarDirectionValues.Column));
                Assert.Contains(plotAreas, plotArea => plotArea.Elements<C.BarChart>()
                    .Any(chart => chart.BarDirection?.Val?.Value == C.BarDirectionValues.Column
                        && chart.BarGrouping?.Val?.Value == C.BarGroupingValues.Stacked));
                Assert.Contains(charts, chart => chart.GetFirstChild<C.Legend>() == null
                    && chart.GetFirstChild<C.PlotArea>()!.Elements<C.BarChart>().Any());
                Assert.Contains(plotAreas.SelectMany(plotArea => plotArea.Descendants<C.DataLabels>()), labels =>
                    labels.GetFirstChild<C.ShowCategoryName>()?.Val?.Value == true
                    && labels.GetFirstChild<C.ShowPercent>()?.Val?.Value == true);
                Assert.Contains(plotAreas.SelectMany(plotArea => plotArea.Descendants<C.DataLabels>()), labels =>
                    labels.GetFirstChild<C.ShowValue>()?.Val?.Value == true
                    && labels.GetFirstChild<C.DataLabelPosition>()?.Val?.Value == C.DataLabelPositionValues.OutsideEnd);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }
        }

        [Fact]
        public void Test_ExcelCharts_ChartLayoutPlacesRecipeChartsWithoutOverlap() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.ChartLayout.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Dashboard");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(1, 2, "Rules");
                sheet.CellValue(2, 1, "High");
                sheet.CellValue(2, 2, 5);
                sheet.CellValue(3, 1, "Medium");
                sheet.CellValue(3, 2, 3);
                sheet.AddTable("A1:B3", hasHeader: true, name: "StatusData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium4);

                ExcelChartGridLayout layout = sheet.ChartLayout(row: 12, column: 4, widthPixels: 520, heightPixels: 300);
                ExcelChartPlacement statusPlacement = layout.Next();
                ExcelChartPlacement contributionPlacement = layout.Next();

                sheet.AddStatusBreakdownChart("A1:B3", statusPlacement.Row, statusPlacement.Column, widthPixels: statusPlacement.WidthPixels, heightPixels: statusPlacement.HeightPixels);
                sheet.AddContributionChart("A1:B3", contributionPlacement.Row, contributionPlacement.Column, widthPixels: contributionPlacement.WidthPixels, heightPixels: contributionPlacement.HeightPixels);
                document.Save();

                Assert.Equal(12, statusPlacement.Row);
                Assert.Equal(4, statusPlacement.Column);
                Assert.Equal(12, contributionPlacement.Row);
                Assert.Equal(13, contributionPlacement.Column);
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var anchors = wsPart.DrawingsPart!.WorksheetDrawing!
                    .Elements<Xdr.OneCellAnchor>()
                    .ToList();

                Assert.Equal(2, anchors.Count);
                Assert.Equal("3", anchors[0].FromMarker!.ColumnId!.Text);
                Assert.Equal("12", anchors[1].FromMarker!.ColumnId!.Text);
                Assert.Equal(4953000L, anchors[0].Extent!.Cx!.Value);
                Assert.Equal(2857500L, anchors[0].Extent!.Cy!.Value);
                Assert.Equal(4953000L, anchors[1].Extent!.Cx!.Value);
                Assert.Equal(2857500L, anchors[1].Extent!.Cy!.Value);
            }
        }

        [Fact]
        public void Test_ExcelCharts_PieOutsideEndDataLabelsArePreserved() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.PieOutsideEndLabels.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Pie");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(1, 2, "Rules");
                sheet.CellValue(2, 1, "High");
                sheet.CellValue(2, 2, 5);
                sheet.CellValue(3, 1, "Medium");
                sheet.CellValue(3, 2, 3);

                sheet.AddChartFromRange("A1:B3", row: 1, column: 4, widthPixels: 480, heightPixels: 300, type: ExcelChartType.Pie, title: "Status")
                    .SetDataLabels(
                        showValue: false,
                        showCategoryName: true,
                        showSeriesName: false,
                        showLegendKey: false,
                        showPercent: true,
                        position: C.DataLabelPositionValues.OutsideEnd,
                        numberFormat: "0%");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                C.DataLabelPosition position = wsPart.DrawingsPart!.ChartParts
                    .SelectMany(part => part.ChartSpace.Descendants<C.DataLabelPosition>())
                    .Single();

                Assert.Equal(C.DataLabelPositionValues.OutsideEnd, position.Val!.Value);
            }
        }

        [Fact]
        public void Test_ExcelCharts_PieBestFitDataLabelsArePreserved() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.PieBestFitLabels.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Pie");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(1, 2, "Rules");
                sheet.CellValue(2, 1, "High");
                sheet.CellValue(2, 2, 5);
                sheet.CellValue(3, 1, "Medium");
                sheet.CellValue(3, 2, 3);

                sheet.AddChartFromRange("A1:B3", row: 1, column: 4, widthPixels: 480, heightPixels: 300, type: ExcelChartType.Pie, title: "Status")
                    .SetDataLabels(
                        showValue: false,
                        showCategoryName: true,
                        showSeriesName: false,
                        showLegendKey: false,
                        showPercent: true,
                        position: C.DataLabelPositionValues.BestFit,
                        numberFormat: "0%");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                C.DataLabelPosition position = wsPart.DrawingsPart!.ChartParts
                    .SelectMany(part => part.ChartSpace.Descendants<C.DataLabelPosition>())
                    .Single();

                Assert.Equal(C.DataLabelPositionValues.BestFit, position.Val!.Value);
            }
        }

        [Fact]
        public void Test_ExcelCharts_RecipeCustomizationDoesNotLeakToNextBuilderChart() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.RecipeBuilderReuse.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Dashboard");
                sheet.CellValue(1, 1, "Month");
                sheet.CellValue(1, 2, "Revenue");
                sheet.CellValue(2, 1, "Jan");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(3, 1, "Feb");
                sheet.CellValue(3, 2, 16);
                sheet.CellValue(4, 1, "Mar");
                sheet.CellValue(4, 2, 13);
                sheet.AddTable("A1:B4", hasHeader: true, name: "RevenueData", style: OfficeIMO.Excel.TableStyle.TableStyleMedium9);

                var builder = sheet.ChartFromTable("RevenueData");
                builder.KpiScorecard("Revenue KPI").At(1, 5);
                builder.ColumnClustered().Title("Revenue").At(18, 5);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var charts = wsPart.DrawingsPart!.ChartParts
                    .Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!)
                    .ToList();

                Assert.Equal(2, charts.Count);
                Assert.Single(charts, chart => chart.GetFirstChild<C.Legend>() == null);
                Assert.Single(charts, chart => chart.GetFirstChild<C.Legend>() != null);
            }
        }

        [Fact]
        public void Test_ExcelCharts_SeriesStyling_Validates() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SeriesStyling.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3", "Q4" },
                    new[] {
                        new ExcelChartSeries("Revenue", new[] { 120d, 180d, 260d, 320d }),
                        new ExcelChartSeries("Costs", new[] { 85d, 94d, 132d, 150d })
                    });

                var chart = sheet.AddChart(data, row: 1, column: 6, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.ColumnClustered, title: "Quarterly");
                chart.SetSeriesFillColor(0, "2563EB")
                    .SetSeriesLineColor(1, "F97316", widthPoints: 0.5);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var barSeries = chartPart.ChartSpace.Descendants<C.BarChartSeries>().First();
                var children = barSeries.ChildElements.Select(child => child.GetType()).ToList();

                Assert.True(
                    children.IndexOf(typeof(C.ChartShapeProperties)) < children.IndexOf(typeof(C.CategoryAxisData)),
                    "Series shape properties must be emitted before category/value data to satisfy the chart schema.");
                Assert.Contains("2563EB", chartPart.ChartSpace.OuterXml, StringComparison.OrdinalIgnoreCase);
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_UpdateData_NoHeaders_StartsBelowRow1() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Update.NoHeaders.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                sheet.CellValues(new[] {
                    (9, 1, (object)"sentinel"),
                    (9, 2, (object)"sentinel"),
                    (10, 1, (object)"A"),
                    (11, 1, (object)"B"),
                    (12, 1, (object)"C"),
                    (10, 2, (object)1d),
                    (11, 2, (object)2d),
                    (12, 2, (object)3d)
                });

                var chart = sheet.AddChartFromRange("A10:B12", row: 1, column: 4, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.ColumnClustered, hasHeaders: false, title: "No Headers");

                var updated = new ExcelChartData(
                    new[] { "D", "E", "F" },
                    new[] { new ExcelChartSeries("Series 1", new[] { 4d, 5d, 6d }) });

                chart.UpdateData(updated);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                Assert.Equal("sentinel", GetCellValue(spreadsheet, wsPart, "A9"));
                Assert.Equal("sentinel", GetCellValue(spreadsheet, wsPart, "B9"));
                Assert.Equal("D", GetCellValue(spreadsheet, wsPart, "A10"));
                Assert.Equal("E", GetCellValue(spreadsheet, wsPart, "A11"));
                Assert.Equal("F", GetCellValue(spreadsheet, wsPart, "A12"));
                Assert.Equal("4", GetCellValue(spreadsheet, wsPart, "B10"));
                Assert.Equal("5", GetCellValue(spreadsheet, wsPart, "B11"));
                Assert.Equal("6", GetCellValue(spreadsheet, wsPart, "B12"));

                var chartPart = wsPart.DrawingsPart!.ChartParts.Single();
                var seriesText = chartPart.ChartSpace.Descendants<C.BarChartSeries>().Single().GetFirstChild<C.SeriesText>()!;
                Assert.DoesNotContain(seriesText.ChildElements, child => child.LocalName == "strLit");
                Assert.Equal("Series 1", seriesText.GetFirstChild<C.NumericValue>()!.Text);

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }
        }

        [Fact]
        public void Test_ExcelCharts_UpdateData_WithQuotedBangSheetName() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Update.QuotedBangSheet.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Report!2026's");
                sheet.Cell(1, 1, "Month");
                sheet.Cell(1, 2, "Sales");
                sheet.Cell(2, 1, "Jan");
                sheet.Cell(3, 1, "Feb");
                sheet.Cell(2, 2, 1d);
                sheet.Cell(3, 2, 2d);

                var chart = sheet.AddChartFromRange("A1:B3", row: 2, column: 6, widthPixels: 480, heightPixels: 320,
                    type: ExcelChartType.Line, hasHeaders: true, title: "Monthly");

                var updated = new ExcelChartData(
                    new[] { "Jan", "Feb", "Mar" },
                    new[] { new ExcelChartSeries("Sales", new[] { 1d, 2d, 3d }) });

                chart.UpdateData(updated);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                var workbookPart = spreadsheet.WorkbookPart!;
                var hostSheet = workbookPart.Workbook.Sheets!.Elements<Sheet>()
                    .First(s => s.Name!.Value == "Report!2026's");
                var wsPart = (WorksheetPart)workbookPart.GetPartById(hostSheet.Id!);

                Assert.Equal("Jan", GetCellValue(spreadsheet, wsPart, "A2"));
                Assert.Equal("Feb", GetCellValue(spreadsheet, wsPart, "A3"));
                Assert.Equal("Mar", GetCellValue(spreadsheet, wsPart, "A4"));
                Assert.Equal("1", GetCellValue(spreadsheet, wsPart, "B2"));
                Assert.Equal("2", GetCellValue(spreadsheet, wsPart, "B3"));
                Assert.Equal("3", GetCellValue(spreadsheet, wsPart, "B4"));

                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
                var lineChart = plotArea.GetFirstChild<C.LineChart>()!;
                var series = lineChart.Elements<C.LineChartSeries>().First();
                string valuesFormula = series.GetFirstChild<C.Values>()!
                    .GetFirstChild<C.NumberReference>()!
                    .Formula!.Text!;
                Assert.Equal("'Report!2026''s'!$B$2:$B$4", valuesFormula);
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;

                Assert.NotNull(plotArea.GetFirstChild<C.ScatterChart>());
                Assert.Equal(2, plotArea.Elements<C.ValueAxis>().Count());
            }

            using (var document = ExcelDocument.Load(filePath, readOnly: true)) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_Scatter_WritesNumericCategoriesToChartDataSheet() {
            const string chartDataSheetName = "OfficeIMO_ChartData";
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Scatter.NumericCategories.xlsx");

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
                WorksheetPart chartHostSheet = spreadsheet.WorkbookPart!.WorksheetParts
                    .First(p => p.DrawingsPart?.ChartParts.Any() == true);
                var chartPart = chartHostSheet.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
                var scatter = plotArea.GetFirstChild<C.ScatterChart>()!;
                var series = scatter.Elements<C.ScatterChartSeries>().First();
                string formula = series.GetFirstChild<C.XValues>()!
                    .GetFirstChild<C.NumberReference>()!
                    .Formula!.Text!;

                int bang = formula.LastIndexOf('!');
                Assert.True(bang > 0);
                string sheetName = formula.Substring(0, bang).Trim().Trim('\'');
                Assert.Equal(chartDataSheetName, sheetName);

                string a1Range = formula.Substring(bang + 1).Replace("$", string.Empty);
                Assert.True(A1.TryParseRange(a1Range, out int r1, out int c1, out int r2, out int c2));
                Assert.Equal(c1, c2);

                var dataSheet = spreadsheet.WorkbookPart.Workbook.Sheets!
                    .OfType<Sheet>()
                    .First(s => s.Name?.Value == chartDataSheetName);
                var dataSheetPart = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(dataSheet.Id!);

                for (int r = r1; r <= r2; r++) {
                    string cellRef = $"{A1.ColumnIndexToLetters(c1)}{r}";
                    var cell = dataSheetPart.Worksheet.Descendants<Cell>()
                        .First(c => c.CellReference != null && c.CellReference.Value == cellRef);
                    Assert.NotEqual(CellValues.SharedString, cell.DataType?.Value);
                }
            }
        }

        [Fact]
        public void Test_ExcelCharts_ComboScatter_ValidatesWithOpenXmlValidator() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.ComboScatter.Validation.xlsx");

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

                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                OpenXmlValidator validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));

                WorksheetPart worksheetPart = GetWorksheetPartWithCharts(spreadsheet);
                var lineChart = worksheetPart.DrawingsPart!.ChartParts
                    .SelectMany(part => part.ChartSpace.Descendants<C.LineChart>())
                    .First();
                var lineSeries = lineChart.Elements<C.LineChartSeries>().First();
                var pointLabel = lineSeries.GetFirstChild<C.DataLabels>()!
                    .Elements<C.DataLabel>()
                    .First(label => label.GetFirstChild<C.Index>()?.Val?.Value == 2U);
                Assert.Equal(
                    C.DataLabelPositionValues.Top,
                    pointLabel.GetFirstChild<C.DataLabelPosition>()!.Val!.Value);
            }

            if (IsWindowsPlatform()) {
                AssertWorkbookOpensViaExcelComWhenAvailable(filePath,
                    "Excel repaired the combo/scatter chart workbook. Line chart point labels must not emit outEnd positions.");
            }
        }

        [Fact]
        public void Test_ExcelCharts_RangeSeriesReadOnlyListUsesIndexerWithoutSnapshotEnumeration() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.RangeSeriesReadOnlyList.xlsx");
            var ranges = new ThrowOnEnumerateReadOnlyList<ExcelChartSeriesRange>(
                new ExcelChartSeriesRange("Series 1", "A2:A4", "B2:B4"));

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorkSheet("Summary");
                sheet.CellValues(new[] {
                    (1, 1, (object)"X"), (1, 2, (object)"Y"),
                    (2, 1, (object)1d), (2, 2, (object)2d),
                    (3, 1, (object)2d), (3, 2, (object)4d),
                    (4, 1, (object)3d), (4, 2, (object)3d)
                }, null);

                sheet.AddScatterChartFromRanges(ranges, row: 1, column: 4, widthPixels: 480, heightPixels: 320, title: "Scatter");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                Assert.NotNull(wsPart.DrawingsPart);
                Assert.Single(wsPart.DrawingsPart!.ChartParts);
                Assert.Single(wsPart.DrawingsPart.WorksheetDrawing!.Elements<Xdr.OneCellAnchor>());
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                WorksheetPart wsPart = GetWorksheetPartByName(spreadsheet, "Data");
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
                WorksheetPart wsPart = GetWorksheetPartByName(spreadsheet, "Data");
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                Assert.Equal(DefaultChartFontSize, categoryTitleProps!.FontSize!.Value);
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
                Assert.Equal(DefaultChartFontSize, valueTitleProps!.FontSize!.Value);

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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                Assert.Equal(DefaultChartFontSize, pointTextProps?.FontSize?.Value);
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
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

        private static string GetCellValue(SpreadsheetDocument document, WorksheetPart worksheetPart, string cellReference) {
            var cell = worksheetPart.Worksheet.Descendants<Cell>()
                .First(c => c.CellReference != null && c.CellReference.Value == cellReference);
            var value = cell.CellValue?.Text ?? string.Empty;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
                var table = document.WorkbookPart?.SharedStringTablePart?.SharedStringTable;
                if (table != null && int.TryParse(value, out int id)) {
                    return table.ChildElements[id].InnerText;
                }
            }
            return value;
        }
    }
}
