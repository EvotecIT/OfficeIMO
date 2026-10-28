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

namespace OfficeIMO.Tests
{
    public partial class Excel
    {
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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3", "Q4" },
                    new[] {
                        new ExcelChartSeries("Sales", new[] { 10d, 20d, 25d, 30d }),
                        new ExcelChartSeries("Target", new[] { 12d, 22d, 24d, 32d })
                    });

                sheet.AddChart(data, row: 1, column: 6, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Quarterly");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                Assert.NotNull(wsPart.DrawingsPart);
                Assert.True(wsPart.DrawingsPart!.ChartParts.Any());

                var hiddenSheets = spreadsheet.WorkbookPart!.Workbook!.Sheets!
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
        public void Test_ExcelCharts_DataPointStyling_WritesPointOverride() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.DataPointStyle.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] {
                        new ExcelChartSeries("Sales", new[] { 10d, 20d, 15d })
                    });

                ExcelChart chart = sheet.AddChart(data, row: 1, column: 6, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Quarterly");

                chart.SetDataPointColor("Sales", pointIndex: 1, fillColor: OfficeColor.Parse("#C00000"), lineColor: OfficeColor.Parse("#7030A0"), lineWidth: OfficeLength.Points(1.25));
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                ChartPart chartPart = wsPart.DrawingsPart!.ChartParts.Single();
                C.DataPoint point = chartPart.ChartSpace!.Descendants<C.DataPoint>().Single();
                Assert.Equal(1U, point.GetFirstChild<C.Index>()!.Val!.Value);
                Assert.Contains(point.Descendants<A.RgbColorModelHex>(), color => color.Val!.Value == "C00000");
                Assert.Contains(point.Descendants<A.RgbColorModelHex>(), color => color.Val!.Value == "7030A0");
                Assert.Equal(15875, point.Descendants<A.Outline>().Single().Width!.Value);
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
        public void Test_ExcelCharts_FluentRangeBuilder_CreatesChart()
        {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.FluentRange.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Summary");
                sheet.CellValue(1, 1, "Quarter");
                sheet.CellValue(1, 2, "Sales");
                sheet.CellValue(1, 3, "Target");
                sheet.CellValue(2, 1, "Q1");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(2, 3, 12);
                sheet.CellValue(3, 1, "Q2");
                sheet.CellValue(3, 2, 20);
                sheet.CellValue(3, 3, 22);
                sheet.AddChartFromRange("A1:C3", row: 1, column: 6, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320), title: "Quarterly", includeCachedData: true, type: OfficeChartType.Line, hasHeaders: true);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
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
        public void Test_ExcelCharts_RadarChart_CanCreateAndUpdate()
        {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Radar.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Coverage");
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
                var chart = sheet.AddChartFromRange("A1:C4", row: 1, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "Coverage", includeCachedData: true, type: OfficeChartType.Radar, hasHeaders: true);
                chart.UpdateData(new ExcelChartData(new[] { "Mail", "DNS", "Identity", "Device" }, new[] { new ExcelChartSeries("Current", new[] { 4d, 3d, 5d, 2d }, OfficeChartType.Radar), new ExcelChartSeries("Target", new[] { 5d, 5d, 5d, 4d }, OfficeChartType.Radar) }));
                chart.SetSeriesLineColor("Current", OfficeColor.Parse("2563EB"), width: OfficeLength.Points(1.25)).SetSeriesMarker("Target", OfficeChartMarkerShape.Circle, size: OfficeLength.Points(6), fillColor: OfficeColor.Parse("F97316"), lineColor: OfficeColor.Parse("7C2D12"));
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
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
                var valuesCache = series.GetFirstChild<C.Values>()!.GetFirstChild<C.NumberReference>()!.GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("2563EB", series.OuterXml, StringComparison.OrdinalIgnoreCase);
                var targetSeries = radarChart.Elements<C.RadarChartSeries>().Skip(1).First();
                Assert.Equal(C.MarkerStyleValues.Circle, targetSeries.GetFirstChild<C.Marker>()?.Symbol?.Val?.Value);
                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_StockChart_CanCreateAndUpdate()
        {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Stock.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Prices");
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
                var chart = sheet.AddChartFromRange("A1:E4", row: 1, column: 7, width: OfficeLength.Pixels(560), height: OfficeLength.Pixels(340), title: "Price Range", includeCachedData: true, type: OfficeChartType.Stock, hasHeaders: true);
                chart.UpdateData(new ExcelChartData(new[] { "2026-01-01", "2026-01-02", "2026-01-03", "2026-01-04" }, new[] { new ExcelChartSeries("Open", new[] { 102d, 105d, 109d, 112d }, OfficeChartType.Stock), new ExcelChartSeries("High", new[] { 108d, 111d, 116d, 118d }, OfficeChartType.Stock), new ExcelChartSeries("Low", new[] { 99d, 101d, 106d, 110d }, OfficeChartType.Stock), new ExcelChartSeries("Close", new[] { 105d, 109d, 113d, 116d }, OfficeChartType.Stock) }));
                chart.SetSeriesLineColor("Close", OfficeColor.Parse("2563EB"), width: OfficeLength.Points(1.25));
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
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
                var valuesCache = closeSeries.GetFirstChild<C.Values>()!.GetFirstChild<C.NumberReference>()!.GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("2563EB", closeSeries.OuterXml, StringComparison.OrdinalIgnoreCase);
                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_SurfaceChart_CanCreateAndUpdate()
        {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Surface.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Terrain");
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
                var chart = sheet.AddChartFromRange("A1:D4", row: 1, column: 6, width: OfficeLength.Pixels(560), height: OfficeLength.Pixels(340), title: "Surface", includeCachedData: true, type: OfficeChartType.Surface, hasHeaders: true);
                chart.UpdateData(new ExcelChartData(new[] { "North", "Central", "South", "West" }, new[] { new ExcelChartSeries("Low", new[] { 11d, 15d, 13d, 17d }, OfficeChartType.Surface), new ExcelChartSeries("Mid", new[] { 18d, 22d, 19d, 23d }, OfficeChartType.Surface), new ExcelChartSeries("High", new[] { 25d, 30d, 28d, 32d }, OfficeChartType.Surface) }));
                chart.SetSeriesFillColor("High", OfficeColor.Parse("22C55E"));
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
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
                var valuesCache = highSeries.GetFirstChild<C.Values>()!.GetFirstChild<C.NumberReference>()!.GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("22C55E", highSeries.OuterXml, StringComparison.OrdinalIgnoreCase);
                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_SurfaceVariantCharts_CanCreateAndUpdate()
        {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SurfaceVariants.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Surface");
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
                var wireframeChart = sheet.AddChartFromRange("A1:D4", row: 1, column: 6, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "Wireframe Surface", includeCachedData: true, type: OfficeChartType.SurfaceWireframe, hasHeaders: true);
                wireframeChart.UpdateData(new ExcelChartData(new[] { "North", "East", "South", "West" }, new[] { new ExcelChartSeries("Low", new[] { 11d, 15d, 13d, 17d }, OfficeChartType.SurfaceWireframe), new ExcelChartSeries("Mid", new[] { 18d, 22d, 19d, 23d }, OfficeChartType.SurfaceWireframe), new ExcelChartSeries("High", new[] { 25d, 30d, 28d, 32d }, OfficeChartType.SurfaceWireframe) }));
                wireframeChart.SetSeriesFillColor("High", OfficeColor.Parse("22C55E"));
                sheet.AddChartFromRange("A1:D4", row: 18, column: 6, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "Contour Surface", includeCachedData: true, type: OfficeChartType.SurfaceContour, hasHeaders: true);
                sheet.AddChartFromRange("A1:D4", row: 35, column: 6, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "Wireframe Contour", includeCachedData: true, type: OfficeChartType.SurfaceContourWireframe, hasHeaders: true);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartParts = wsPart.DrawingsPart!.ChartParts.ToList();
                Assert.Equal(3, chartParts.Count);
                var plotAreas = chartParts.Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!).ToList();
                var surface3DChart = plotAreas.Select(plotArea => plotArea.GetFirstChild<C.Surface3DChart>()).Single(chart => chart != null);
                var contourCharts = plotAreas.Select(plotArea => plotArea.GetFirstChild<C.SurfaceChart>()).Where(chart => chart != null).ToList();
                Assert.NotNull(surface3DChart);
                Assert.True(surface3DChart!.GetFirstChild<C.Wireframe>()!.Val!.Value);
                Assert.Equal(3, surface3DChart.Elements<C.SurfaceChartSeries>().Count());
                Assert.Equal(3, surface3DChart.Elements<C.AxisId>().Count());
                Assert.NotNull(plotAreas.First(plotArea => plotArea.GetFirstChild<C.Surface3DChart>() == surface3DChart).GetFirstChild<C.SeriesAxis>());
                var highSeries = surface3DChart.Elements<C.SurfaceChartSeries>().Last();
                var valuesCache = highSeries.GetFirstChild<C.Values>()!.GetFirstChild<C.NumberReference>()!.GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("22C55E", highSeries.OuterXml, StringComparison.OrdinalIgnoreCase);
                Assert.Equal(2, contourCharts.Count);
                Assert.Contains(contourCharts, chart => chart!.GetFirstChild<C.Wireframe>()!.Val!.Value == false);
                Assert.Contains(contourCharts, chart => chart!.GetFirstChild<C.Wireframe>()!.Val!.Value);
                foreach (var contourChart in contourCharts)
                {
                    Assert.Equal(3, contourChart!.Elements<C.SurfaceChartSeries>().Count());
                    Assert.Equal(3, contourChart.Elements<C.AxisId>().Count());
                    Assert.NotNull(plotAreas.First(plotArea => plotArea.GetFirstChild<C.SurfaceChart>() == contourChart).GetFirstChild<C.SeriesAxis>());
                }

                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_Pie3DChart_CanCreateAndUpdate()
        {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Pie3D.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Mix");
                sheet.CellValue(1, 1, "Category");
                sheet.CellValue(1, 2, "Share");
                sheet.CellValue(2, 1, "Mail");
                sheet.CellValue(2, 2, 42);
                sheet.CellValue(3, 1, "DNS");
                sheet.CellValue(3, 2, 28);
                sheet.CellValue(4, 1, "Identity");
                sheet.CellValue(4, 2, 30);
                var chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "Workload Mix", includeCachedData: true, type: OfficeChartType.Pie3D, hasHeaders: true);
                chart.UpdateData(new ExcelChartData(new[] { "Mail", "DNS", "Identity", "Device" }, new[] { new ExcelChartSeries("Share", new[] { 40d, 24d, 26d, 10d }, OfficeChartType.Pie3D) }));
                chart.SetDataLabels(showValue: true, showCategoryName: true, showSeriesName: false, showLegendKey: false, showPercent: false).SetSeriesFillColor("Share", OfficeColor.Parse("2563EB"));
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartPart = wsPart.DrawingsPart!.ChartParts.Single();
                var chart = chartPart.ChartSpace.GetFirstChild<C.Chart>()!;
                var plotArea = chart.GetFirstChild<C.PlotArea>()!;
                var pieChart = plotArea.GetFirstChild<C.Pie3DChart>();
                Assert.NotNull(pieChart);
                Assert.Equal("Workload Mix", chart.Title!.Descendants<A.Text>().First().Text);
                Assert.True(pieChart!.VaryColors?.Val?.Value);
                var series = Assert.Single(pieChart.Elements<C.PieChartSeries>());
                var valuesCache = series.GetFirstChild<C.Values>()!.GetFirstChild<C.NumberReference>()!.GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("2563EB", series.OuterXml, StringComparison.OrdinalIgnoreCase);
                var labels = pieChart.GetFirstChild<C.DataLabels>();
                Assert.True(labels!.GetFirstChild<C.ShowValue>()!.Val!.Value);
                Assert.True(labels.GetFirstChild<C.ShowCategoryName>()!.Val!.Value);
                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_OfPieCharts_CanCreateAndUpdate()
        {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.OfPie.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Share");
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
                var pieChart = sheet.AddChartFromRange("A1:B5", row: 1, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "Pie of Pie", includeCachedData: true, type: OfficeChartType.PieOfPie, hasHeaders: true);
                pieChart.UpdateData(new ExcelChartData(new[] { "Core", "Growth", "Services", "Other", "Legacy" }, new[] { new ExcelChartSeries("Share", new[] { 38d, 25d, 24d, 9d, 4d }, OfficeChartType.PieOfPie) }));
                pieChart.SetDataLabels(showValue: true, showCategoryName: true, showSeriesName: false, showLegendKey: false, showPercent: false).SetSeriesFillColor("Share", OfficeColor.Parse("2563EB"));
                sheet.AddChartFromRange("A1:B5", row: 18, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "Bar of Pie", includeCachedData: true, type: OfficeChartType.BarOfPie, hasHeaders: true);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartParts = wsPart.DrawingsPart!.ChartParts.ToList();
                Assert.Equal(2, chartParts.Count);
                var ofPieCharts = chartParts.Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!.GetFirstChild<C.OfPieChart>()).ToList();
                var pieChart = ofPieCharts.First(chart => chart?.GetFirstChild<C.OfPieType>()?.Val?.Value == C.OfPieValues.Pie);
                var barChart = ofPieCharts.First(chart => chart?.GetFirstChild<C.OfPieType>()?.Val?.Value == C.OfPieValues.Bar);
                Assert.NotNull(pieChart);
                Assert.Single(pieChart!.Elements<C.PieChartSeries>());
                Assert.NotNull(pieChart.GetFirstChild<C.SeriesLines>());
                Assert.Equal((ushort)75, pieChart.GetFirstChild<C.SecondPieSize>()!.Val!.Value);
                Assert.Equal(C.SplitValues.Position, pieChart.GetFirstChild<C.SplitType>()!.Val!.Value);
                var series = Assert.Single(pieChart.Elements<C.PieChartSeries>());
                var valuesCache = series.GetFirstChild<C.Values>()!.GetFirstChild<C.NumberReference>()!.GetFirstChild<C.NumberingCache>()!;
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

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_BarStacked100Variants_CanCreateAndUpdate()
        {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.BarStacked100.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Mix");
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
                var columnChart = sheet.AddChartFromRange("A1:C4", row: 1, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "100% Stacked Columns", includeCachedData: true, type: OfficeChartType.ColumnStacked100, hasHeaders: true);
                columnChart.UpdateData(new ExcelChartData(new[] { "Q1", "Q2", "Q3", "Q4" }, new[] { new ExcelChartSeries("Sales", new[] { 12d, 19d, 25d, 31d }, OfficeChartType.ColumnStacked100), new ExcelChartSeries("Target", new[] { 14d, 21d, 28d, 34d }, OfficeChartType.ColumnStacked100) }));
                columnChart.SetSeriesFillColor("Sales", OfficeColor.Parse("2563EB")).SetDataLabels(showValue: true);
                sheet.AddChartFromRange("A1:C4", row: 18, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "100% Stacked Bars", includeCachedData: true, type: OfficeChartType.BarStacked100, hasHeaders: true);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartParts = wsPart.DrawingsPart!.ChartParts.ToList();
                Assert.Equal(2, chartParts.Count);
                var barCharts = chartParts.Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!.GetFirstChild<C.BarChart>()).ToList();
                var columnChart = barCharts.First(chart => chart?.BarDirection?.Val?.Value == C.BarDirectionValues.Column);
                var barChart = barCharts.First(chart => chart?.BarDirection?.Val?.Value == C.BarDirectionValues.Bar);
                Assert.NotNull(columnChart);
                Assert.Equal(C.BarGroupingValues.PercentStacked, columnChart!.BarGrouping?.Val?.Value);
                Assert.Equal(2, columnChart.Elements<C.BarChartSeries>().Count());
                Assert.Equal(2, columnChart.Elements<C.AxisId>().Count());
                var series = columnChart.Elements<C.BarChartSeries>().First();
                var valuesCache = series.GetFirstChild<C.Values>()!.GetFirstChild<C.NumberReference>()!.GetFirstChild<C.NumberingCache>()!;
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

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_Bar3DCharts_CanCreateAndUpdate()
        {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Bar3D.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Sales");
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
                var columnChart = sheet.AddChartFromRange("A1:C4", row: 1, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "3D Columns", includeCachedData: true, type: OfficeChartType.Column3DClustered, hasHeaders: true);
                columnChart.UpdateData(new ExcelChartData(new[] { "Q1", "Q2", "Q3", "Q4" }, new[] { new ExcelChartSeries("Sales", new[] { 12d, 19d, 25d, 31d }, OfficeChartType.Column3DClustered), new ExcelChartSeries("Target", new[] { 14d, 21d, 28d, 34d }, OfficeChartType.Column3DClustered) }));
                columnChart.SetSeriesFillColor("Sales", OfficeColor.Parse("2563EB")).SetDataLabels(showValue: true);
                sheet.AddChartFromRange("A1:C4", row: 18, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "3D Bars", includeCachedData: true, type: OfficeChartType.Bar3DStacked, hasHeaders: true);
                sheet.AddChartFromRange("A1:C4", row: 35, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "3D 100% Columns", includeCachedData: true, type: OfficeChartType.Column3DStacked100, hasHeaders: true);
                sheet.AddChartFromRange("A1:C4", row: 52, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "3D 100% Bars", includeCachedData: true, type: OfficeChartType.Bar3DStacked100, hasHeaders: true);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartParts = wsPart.DrawingsPart!.ChartParts.ToList();
                Assert.Equal(4, chartParts.Count);
                var plotAreas = chartParts.Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!).ToList();
                var columnChart = plotAreas.Select(plotArea => plotArea.GetFirstChild<C.Bar3DChart>()).First(chart => chart?.BarDirection?.Val?.Value == C.BarDirectionValues.Column && chart.BarGrouping?.Val?.Value == C.BarGroupingValues.Clustered);
                var barChart = plotAreas.Select(plotArea => plotArea.GetFirstChild<C.Bar3DChart>()).First(chart => chart?.BarDirection?.Val?.Value == C.BarDirectionValues.Bar && chart.BarGrouping?.Val?.Value == C.BarGroupingValues.Stacked);
                var percentColumnChart = plotAreas.Select(plotArea => plotArea.GetFirstChild<C.Bar3DChart>()).First(chart => chart?.BarDirection?.Val?.Value == C.BarDirectionValues.Column && chart.BarGrouping?.Val?.Value == C.BarGroupingValues.PercentStacked);
                var percentBarChart = plotAreas.Select(plotArea => plotArea.GetFirstChild<C.Bar3DChart>()).First(chart => chart?.BarDirection?.Val?.Value == C.BarDirectionValues.Bar && chart.BarGrouping?.Val?.Value == C.BarGroupingValues.PercentStacked);
                Assert.NotNull(columnChart);
                Assert.Equal(C.BarGroupingValues.Clustered, columnChart!.BarGrouping?.Val?.Value);
                Assert.Equal(2, columnChart.Elements<C.BarChartSeries>().Count());
                Assert.Equal(3, columnChart.Elements<C.AxisId>().Count());
                Assert.Equal(C.ShapeValues.Box, columnChart.GetFirstChild<C.Shape>()?.Val?.Value);
                Assert.NotNull(plotAreas.First(plotArea => plotArea.GetFirstChild<C.Bar3DChart>() == columnChart).GetFirstChild<C.SeriesAxis>());
                var series = columnChart.Elements<C.BarChartSeries>().First();
                var valuesCache = series.GetFirstChild<C.Values>()!.GetFirstChild<C.NumberReference>()!.GetFirstChild<C.NumberingCache>()!;
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

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_LineStackedVariants_CanCreateAndUpdate()
        {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.LineStacked.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Trend");
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
                var stackedChart = sheet.AddChartFromRange("A1:C4", row: 1, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "Stacked Line", includeCachedData: true, type: OfficeChartType.LineStacked, hasHeaders: true);
                stackedChart.UpdateData(new ExcelChartData(new[] { "Jan", "Feb", "Mar", "Apr" }, new[] { new ExcelChartSeries("Actual", new[] { 11d, 16d, 22d, 27d }, OfficeChartType.LineStacked), new ExcelChartSeries("Plan", new[] { 12d, 18d, 24d, 29d }, OfficeChartType.LineStacked) }));
                stackedChart.SetSeriesLineColor("Actual", OfficeColor.Parse("2563EB"), width: OfficeLength.Points(1.25)).SetSeriesMarker("Plan", OfficeChartMarkerShape.Circle, size: OfficeLength.Points(6), fillColor: OfficeColor.Parse("F97316"), lineColor: OfficeColor.Parse("7C2D12")).SetDataLabels(showValue: true);
                sheet.AddChartFromRange("A1:C4", row: 18, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "100% Stacked Line", includeCachedData: true, type: OfficeChartType.LineStacked100, hasHeaders: true);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartParts = wsPart.DrawingsPart!.ChartParts.ToList();
                Assert.Equal(2, chartParts.Count);
                var lineCharts = chartParts.Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!.GetFirstChild<C.LineChart>()).ToList();
                var stackedChart = lineCharts.First(chart => chart?.Grouping?.Val?.Value == C.GroupingValues.Stacked);
                var percentChart = lineCharts.First(chart => chart?.Grouping?.Val?.Value == C.GroupingValues.PercentStacked);
                Assert.NotNull(stackedChart);
                Assert.Equal(2, stackedChart!.Elements<C.LineChartSeries>().Count());
                Assert.Equal(2, stackedChart.Elements<C.AxisId>().Count());
                var series = stackedChart.Elements<C.LineChartSeries>().First();
                var valuesCache = series.GetFirstChild<C.Values>()!.GetFirstChild<C.NumberReference>()!.GetFirstChild<C.NumberingCache>()!;
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

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Theory]
        [InlineData(OfficeChartType.Pie3D, "3-D pie")]
        [InlineData(OfficeChartType.PieOfPie, "pie-of-pie")]
        [InlineData(OfficeChartType.BarOfPie, "bar-of-pie")]
        public void Test_ExcelCharts_SingleSeriesPieVariantsRejectMultipleSeries(OfficeChartType type, string chartName) {
            string filePath = Path.Combine(_directoryWithFiles, $"ExcelCharts.{type}.MultipleSeries.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Share");
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
        public void Test_ExcelCharts_Line3DChart_CanCreateAndUpdate()
        {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Line3D.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Trend");
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
                var chart = sheet.AddChartFromRange("A1:C4", row: 1, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "3D Line", includeCachedData: true, type: OfficeChartType.Line3D, hasHeaders: true);
                chart.UpdateData(new ExcelChartData(new[] { "Jan", "Feb", "Mar", "Apr" }, new[] { new ExcelChartSeries("Actual", new[] { 11d, 16d, 22d, 27d }, OfficeChartType.Line3D), new ExcelChartSeries("Plan", new[] { 12d, 18d, 24d, 29d }, OfficeChartType.Line3D) }));
                chart.SetSeriesLineColor("Actual", OfficeColor.Parse("2563EB"), width: OfficeLength.Points(1.25)).SetSeriesMarker("Plan", OfficeChartMarkerShape.Circle, size: OfficeLength.Points(6), fillColor: OfficeColor.Parse("F97316"), lineColor: OfficeColor.Parse("7C2D12")).SetDataLabels(showValue: true);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
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
                var valuesCache = series.GetFirstChild<C.Values>()!.GetFirstChild<C.NumberReference>()!.GetFirstChild<C.NumberingCache>()!;
                Assert.Equal((uint)4, valuesCache.PointCount!.Val!.Value);
                Assert.Contains("2563EB", series.OuterXml, StringComparison.OrdinalIgnoreCase);
                Assert.True(lineChart.GetFirstChild<C.DataLabels>()!.GetFirstChild<C.ShowValue>()!.Val!.Value);
                var planSeries = lineChart.Elements<C.LineChartSeries>().Skip(1).First();
                Assert.Equal(C.MarkerStyleValues.Circle, planSeries.GetFirstChild<C.Marker>()?.Symbol?.Val?.Value);
                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_AreaStackedVariants_CanCreateAndUpdate()
        {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.AreaStacked.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Mix");
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
                var stackedChart = sheet.AddChartFromRange("A1:C4", row: 1, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "Stacked Area", includeCachedData: true, type: OfficeChartType.AreaStacked, hasHeaders: true);
                stackedChart.UpdateData(new ExcelChartData(new[] { "Jan", "Feb", "Mar", "Apr" }, new[] { new ExcelChartSeries("Services", new[] { 36d, 44d, 50d, 54d }, OfficeChartType.AreaStacked), new ExcelChartSeries("Licenses", new[] { 19d, 25d, 31d, 34d }, OfficeChartType.AreaStacked) }));
                stackedChart.SetSeriesFillColor("Services", OfficeColor.Parse("2563EB")).SetDataLabels(showValue: true);
                sheet.AddChartFromRange("A1:C4", row: 18, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "100% Stacked Area", includeCachedData: true, type: OfficeChartType.AreaStacked100, hasHeaders: true);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartParts = wsPart.DrawingsPart!.ChartParts.ToList();
                Assert.Equal(2, chartParts.Count);
                var areaCharts = chartParts.Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!.GetFirstChild<C.AreaChart>()).ToList();
                var stackedChart = areaCharts.First(chart => chart?.Grouping?.Val?.Value == C.GroupingValues.Stacked);
                var percentChart = areaCharts.First(chart => chart?.Grouping?.Val?.Value == C.GroupingValues.PercentStacked);
                Assert.NotNull(stackedChart);
                Assert.Equal(2, stackedChart!.Elements<C.AreaChartSeries>().Count());
                Assert.Equal(2, stackedChart.Elements<C.AxisId>().Count());
                var series = stackedChart.Elements<C.AreaChartSeries>().First();
                var valuesCache = series.GetFirstChild<C.Values>()!.GetFirstChild<C.NumberReference>()!.GetFirstChild<C.NumberingCache>()!;
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

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_Area3DCharts_CanCreateAndUpdate()
        {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Area3D.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Volume");
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
                var areaChart = sheet.AddChartFromRange("A1:C4", row: 1, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "3D Area", includeCachedData: true, type: OfficeChartType.Area3D, hasHeaders: true);
                areaChart.UpdateData(new ExcelChartData(new[] { "Jan", "Feb", "Mar", "Apr" }, new[] { new ExcelChartSeries("Actual", new[] { 11d, 17d, 22d, 28d }, OfficeChartType.Area3D), new ExcelChartSeries("Plan", new[] { 13d, 19d, 25d, 31d }, OfficeChartType.Area3D) }));
                areaChart.SetSeriesFillColor("Actual", OfficeColor.Parse("2563EB")).SetDataLabels(showValue: true);
                sheet.AddChartFromRange("A1:C4", row: 18, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "3D Stacked Area", includeCachedData: true, type: OfficeChartType.Area3DStacked, hasHeaders: true);
                sheet.AddChartFromRange("A1:C4", row: 35, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(340), title: "3D 100% Stacked Area", includeCachedData: true, type: OfficeChartType.Area3DStacked100, hasHeaders: true);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartParts = wsPart.DrawingsPart!.ChartParts.ToList();
                Assert.Equal(3, chartParts.Count);
                var plotAreas = chartParts.Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!).ToList();
                var standardChart = plotAreas.Select(plotArea => plotArea.GetFirstChild<C.Area3DChart>()).First(chart => chart?.Grouping?.Val?.Value == C.GroupingValues.Standard);
                var stackedChart = plotAreas.Select(plotArea => plotArea.GetFirstChild<C.Area3DChart>()).First(chart => chart?.Grouping?.Val?.Value == C.GroupingValues.Stacked);
                var percentChart = plotAreas.Select(plotArea => plotArea.GetFirstChild<C.Area3DChart>()).First(chart => chart?.Grouping?.Val?.Value == C.GroupingValues.PercentStacked);
                Assert.NotNull(standardChart);
                Assert.Equal(2, standardChart!.Elements<C.AreaChartSeries>().Count());
                Assert.Equal(3, standardChart.Elements<C.AxisId>().Count());
                Assert.NotNull(standardChart.GetFirstChild<C.GapDepth>());
                Assert.NotNull(plotAreas.First(plotArea => plotArea.GetFirstChild<C.Area3DChart>() == standardChart).GetFirstChild<C.SeriesAxis>());
                var series = standardChart.Elements<C.AreaChartSeries>().First();
                var valuesCache = series.GetFirstChild<C.Values>()!.GetFirstChild<C.NumberReference>()!.GetFirstChild<C.NumberingCache>()!;
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

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly }))
            {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_RecipeHelpers_CreateExpectedChartTypes()
        {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.RecipeHelpers.xlsx");
            using (var document = ExcelDocument.Create(filePath))
            {
                var sheet = document.AddWorksheet("Dashboard");
                sheet.CellValue(1, 1, "Month");
                sheet.CellValue(1, 2, "Revenue");
                sheet.CellValue(2, 1, "Jan");
                sheet.CellValue(2, 2, 10);
                sheet.CellValue(3, 1, "Feb");
                sheet.CellValue(3, 2, 16);
                sheet.CellValue(4, 1, "Mar");
                sheet.CellValue(4, 2, 13);
                sheet.AddTable("A1:B4", hasHeader: true, name: "RevenueData", style: OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium9);
                sheet.AddRevenueTrendChart("A1:B4", row: 1, column: 5);
                sheet.AddTopNBarChart("A1:B4", row: 18, column: 5, title: "Top Revenue");
                sheet.AddVarianceColumnChart("A1:B4", row: 35, column: 5, title: "Revenue Variance");
                sheet.AddKpiScorecardChart("A1:B4", row: 52, column: 5, title: "Revenue KPI");
                sheet.AddContributionChart("A1:B4", row: 69, column: 5, title: "Revenue Mix");
                sheet.AddStatusBreakdownChart(sheet.GetTableRange("RevenueData")!, row: 86, column: 5, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(320), title: "Revenue Mix", includeCachedData: true, hasHeaders: true);
                sheet.AddVarianceWaterfallChart(sheet.GetTableRange("RevenueData")!, row: 103, column: 5, width: OfficeLength.Pixels(720), height: OfficeLength.Pixels(360), title: "Revenue Bridge", includeCachedData: true, hasHeaders: true);
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false))
            {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var charts = wsPart.DrawingsPart!.ChartParts.Select(part => part.ChartSpace.GetFirstChild<C.Chart>()!).ToList();
                var plotAreas = charts.Select(chart => chart.GetFirstChild<C.PlotArea>()!).ToList();
                Assert.Equal(7, plotAreas.Count);
                Assert.Contains(plotAreas, plotArea => plotArea.GetFirstChild<C.LineChart>() != null);
                Assert.Contains(plotAreas, plotArea => plotArea.GetFirstChild<C.DoughnutChart>() != null);
                Assert.Contains(plotAreas, plotArea => plotArea.Elements<C.BarChart>().Any(chart => chart.BarDirection?.Val?.Value == C.BarDirectionValues.Bar));
                Assert.Contains(plotAreas, plotArea => plotArea.Elements<C.BarChart>().Any(chart => chart.BarDirection?.Val?.Value == C.BarDirectionValues.Column));
                Assert.Contains(plotAreas, plotArea => plotArea.Elements<C.BarChart>().Any(chart => chart.BarDirection?.Val?.Value == C.BarDirectionValues.Column && chart.BarGrouping?.Val?.Value == C.BarGroupingValues.Stacked));
                Assert.Contains(charts, chart => chart.GetFirstChild<C.Legend>() == null && chart.GetFirstChild<C.PlotArea>()!.Elements<C.BarChart>().Any());
                Assert.Contains(plotAreas.SelectMany(plotArea => plotArea.Descendants<C.DataLabels>()), labels => labels.GetFirstChild<C.ShowCategoryName>()?.Val?.Value == true && labels.GetFirstChild<C.ShowPercent>()?.Val?.Value == true);
                Assert.Contains(plotAreas.SelectMany(plotArea => plotArea.Descendants<C.DataLabels>()), labels => labels.GetFirstChild<C.ShowValue>()?.Val?.Value == true && labels.GetFirstChild<C.DataLabelPosition>()?.Val?.Value == C.DataLabelPositionValues.OutsideEnd);
                OpenXmlValidator validator = new OpenXmlValidator();
                var errors = validator.Validate(spreadsheet).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }
        }

        [Fact]
        public void Test_ExcelCharts_ChartLayoutPlacesRecipeChartsWithoutOverlap() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.ChartLayout.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Dashboard");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(1, 2, "Rules");
                sheet.CellValue(2, 1, "High");
                sheet.CellValue(2, 2, 5);
                sheet.CellValue(3, 1, "Medium");
                sheet.CellValue(3, 2, 3);
                sheet.AddTable("A1:B3", hasHeader: true, name: "StatusData", style: OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium4);

                ExcelChartGridLayout layout = sheet.ChartLayout(row: 12, column: 4, width: OfficeLength.Pixels(520), height: OfficeLength.Pixels(300));
                ExcelChartPlacement statusPlacement = layout.Next();
                ExcelChartPlacement contributionPlacement = layout.Next();

                sheet.AddStatusBreakdownChart("A1:B3", statusPlacement.Row, statusPlacement.Column, width: statusPlacement.Width, height: statusPlacement.Height);
                sheet.AddContributionChart("A1:B3", contributionPlacement.Row, contributionPlacement.Column, width: contributionPlacement.Width, height: contributionPlacement.Height);
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
                var sheet = document.AddWorksheet("Pie");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(1, 2, "Rules");
                sheet.CellValue(2, 1, "High");
                sheet.CellValue(2, 2, 5);
                sheet.CellValue(3, 1, "Medium");
                sheet.CellValue(3, 2, 3);

                sheet.AddChartFromRange("A1:B3", row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(300), type: OfficeChartType.Pie, title: "Status")
                    .SetDataLabels(
                        showValue: false,
                        showCategoryName: true,
                        showSeriesName: false,
                        showLegendKey: false,
                        showPercent: true,
                        position: OfficeChartDataLabelPosition.OutsideEnd,
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
                var sheet = document.AddWorksheet("Pie");
                sheet.CellValue(1, 1, "Status");
                sheet.CellValue(1, 2, "Rules");
                sheet.CellValue(2, 1, "High");
                sheet.CellValue(2, 2, 5);
                sheet.CellValue(3, 1, "Medium");
                sheet.CellValue(3, 2, 3);

                sheet.AddChartFromRange("A1:B3", row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(300), type: OfficeChartType.Pie, title: "Status")
                    .SetDataLabels(
                        showValue: false,
                        showCategoryName: true,
                        showSeriesName: false,
                        showLegendKey: false,
                        showPercent: true,
                        position: OfficeChartDataLabelPosition.BestFit,
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
        public void Test_ExcelCharts_SeriesStyling_Validates() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SeriesStyling.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3", "Q4" },
                    new[] {
                        new ExcelChartSeries("Revenue", new[] { 120d, 180d, 260d, 320d }),
                        new ExcelChartSeries("Costs", new[] { 85d, 94d, 132d, 150d })
                    });

                var chart = sheet.AddChart(data, row: 1, column: 6, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Quarterly");
                chart.SetSeriesFillColor(0, OfficeColor.Parse("2563EB"))
                    .SetSeriesLineColor(1, OfficeColor.Parse("F97316"), width: OfficeLength.Points(0.5));
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

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_UpdateData() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Update.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Jan", "Feb" },
                    new[] { new ExcelChartSeries("Sales", new[] { 1d, 2d }) });

                var chart = sheet.AddChart(data, row: 2, column: 6, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.Line, title: "Monthly");

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

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_UpdateData_NoHeaders_StartsBelowRow1() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Update.NoHeaders.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Summary");
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

                var chart = sheet.AddChartFromRange("A10:B12", row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, hasHeaders: false, title: "No Headers");

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
                var sheet = document.AddWorksheet("Report!2026's");
                sheet.Cell(1, 1, "Month");
                sheet.Cell(1, 2, "Sales");
                sheet.Cell(2, 1, "Jan");
                sheet.Cell(3, 1, "Feb");
                sheet.Cell(2, 2, 1d);
                sheet.Cell(3, 2, 2d);

                var chart = sheet.AddChartFromRange("A1:B3", row: 2, column: 6, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.Line, hasHeaders: true, title: "Monthly");

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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "1", "2", "3" },
                    new[] { new ExcelChartSeries("Series 1", new[] { 2d, 4d, 6d }, OfficeChartType.Scatter) });

                sheet.AddChart(data, row: 1, column: 6, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.Scatter, title: "Scatter");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, false)) {
                WorksheetPart wsPart = GetWorksheetPartWithCharts(spreadsheet);
                var chartPart = wsPart.DrawingsPart!.ChartParts.First();
                var plotArea = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;

                Assert.NotNull(plotArea.GetFirstChild<C.ScatterChart>());
                Assert.Equal(2, plotArea.Elements<C.ValueAxis>().Count());
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                Assert.Empty(document.ValidateOpenXml());
            }
        }

        [Fact]
        public void Test_ExcelCharts_Scatter_WritesNumericCategoriesToChartDataSheet() {
            const string chartDataSheetName = "OfficeIMO_ChartData";
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Scatter.NumericCategories.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "1", "2", "3" },
                    new[] { new ExcelChartSeries("Series 1", new[] { 2d, 4d, 6d }, OfficeChartType.Scatter) });

                sheet.AddChart(data, row: 1, column: 6, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.Scatter, title: "Scatter");
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
                Assert.True(ExcelA1.TryParseRange(a1Range, out int r1, out int c1, out int r2, out int c2));
                Assert.Equal(c1, c2);

                var dataSheet = spreadsheet.WorkbookPart.Workbook.Sheets!
                    .OfType<Sheet>()
                    .First(s => s.Name?.Value == chartDataSheetName);
                var dataSheetPart = (WorksheetPart)spreadsheet.WorkbookPart.GetPartById(dataSheet.Id!);

                for (int r = r1; r <= r2; r++) {
                    string cellRef = $"{ExcelA1.ColumnIndexToLetters(c1)}{r}";
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
                var sheet = document.AddWorksheet("Summary");

                var comboData = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3", "Q4" },
                    new[] {
                        new ExcelChartSeries("Sales", new[] { 10d, 20d, 25d, 30d }, OfficeChartType.ColumnClustered, OfficeChartAxisGroup.Primary),
                        new ExcelChartSeries("Trend", new[] { 12d, 18d, 28d, 35d }, OfficeChartType.Line, OfficeChartAxisGroup.Secondary)
                    });

                var comboChart = sheet.AddChart(comboData, row: 2, column: 6, width: OfficeLength.Pixels(640), height: OfficeLength.Pixels(360),
                    type: OfficeChartType.ColumnClustered, title: "Sales vs Trend");
                comboChart.ApplyStylePreset()
                          .SetSeriesMarker(1, OfficeChartMarkerShape.Circle, size: OfficeLength.Points(6), lineColor: OfficeColor.Parse("4472C4"));
                comboChart.SetValueAxisNumberFormat("0.00", sourceLinked: false, axisGroup: OfficeChartAxisGroup.Secondary)
                          .SetSeriesDataLabels(1, showValue: true, position: OfficeChartDataLabelPosition.Top, numberFormat: "0.0")
                          .SetSeriesDataLabelTextStyle(1, fontSize: OfficeLength.Points(9), color: OfficeColor.Parse("1F4E79"))
                          .SetSeriesDataLabelShapeStyle(1, fillColor: OfficeColor.Parse("FFFFFF"), lineColor: OfficeColor.Parse("1F4E79"), lineWidth: OfficeLength.Points(0.5));
                comboChart.SetSeriesDataLabelLeaderLines(1, showLeaderLines: true, lineColor: OfficeColor.Parse("1F4E79"), lineWidth: OfficeLength.Points(0.5));
                comboChart.SetLegend(OfficeChartLegendPosition.Right)
                          .SetTitleTextStyle(fontSize: OfficeLength.Points(14), bold: true, color: OfficeColor.Parse("1F4E79"))
                          .SetLegendTextStyle(fontSize: OfficeLength.Points(9), color: OfficeColor.Parse("404040"))
                          .SetCategoryAxisTitle("Quarter")
                          .SetValueAxisTitle("Revenue")
                          .SetCategoryAxisTitleTextStyle(fontSize: OfficeLength.Points(10), bold: true)
                          .SetValueAxisLabelTextStyle(fontSize: OfficeLength.Points(9), color: OfficeColor.Parse("404040"))
                          .SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: OfficeColor.Parse("C0C0C0"), lineWidth: OfficeLength.Points(0.75))
                          .SetCategoryAxisLabelRotation(45)
                          .SetValueAxisTickLabelPosition(OfficeChartAxisTickLabelPosition.Low);
                comboChart.SetCategoryAxisReverseOrder()
                          .SetValueAxisScale(minimum: 0, maximum: 40, majorUnit: 10, minorUnit: 5);
                comboChart.SetValueAxisCrossing(OfficeChartAxisCrossingPosition.Maximum)
                          .SetCategoryAxisCrossing(OfficeChartAxisCrossingPosition.Minimum)
                          .SetValueAxisCrossBetween(OfficeChartAxisCrossBetween.Between)
                          .SetValueAxisDisplayUnits(OfficeChartDisplayUnit.Thousands, "Thousands USD", showLabel: true);
                comboChart.SetChartAreaStyle(fillColor: OfficeColor.Parse("F2F2F2"), lineColor: OfficeColor.Parse("404040"), lineWidth: OfficeLength.Points(1))
                          .SetPlotAreaStyle(fillColor: OfficeColor.Parse("FFFFFF"), lineColor: OfficeColor.Parse("BFBFBF"), lineWidth: OfficeLength.Points(0.75));
                comboChart.SetSeriesTrendline(1, OfficeChartTrendlineType.Linear, displayEquation: true, displayRSquared: true,
                    lineColor: OfficeColor.Parse("A5A5A5"), lineWidth: OfficeLength.Points(1));
                var labelTemplate = new ExcelChartDataLabelTemplate {
                    ShowValue = true,
                    Position = OfficeChartDataLabelPosition.Top,
                    NumberFormat = "0.0",
                    FontSize = OfficeLength.Points(9),
                    TextColor = OfficeColor.Parse("404040"),
                    Separator = " - "
                };
                comboChart.SetSeriesDataLabelTemplate(1, labelTemplate)
                          .SetSeriesDataLabelForPoint(1, 2, showValue: true, position: OfficeChartDataLabelPosition.OutsideEnd,
                            numberFormat: "0.00")
                          .SetSeriesDataLabelSeparatorForPoint(1, 2, " | ")
                          .SetSeriesDataLabelTextStyleForPoint(1, 2, fontSize: OfficeLength.Points(11), bold: true, color: OfficeColor.Parse("FF0000"));

                var scatterData = new ExcelChartData(
                    new[] { "1", "2", "3", "4" },
                    new[] { new ExcelChartSeries("Points", new[] { 2d, 4d, 3d, 5d }, OfficeChartType.Scatter) });

                var scatterChart = sheet.AddChart(scatterData, row: 22, column: 6, width: OfficeLength.Pixels(640), height: OfficeLength.Pixels(360),
                    type: OfficeChartType.Scatter, title: "Scatter Sample");
                scatterChart.SetScatterXAxisScale(minimum: 1, maximum: 10, majorUnit: 1, logScale: true);
                scatterChart.SetScatterYAxisScale(minimum: 0, maximum: 6, majorUnit: 1);
                scatterChart.SetScatterYAxisCrossing(OfficeChartAxisCrossingPosition.Minimum, crossesAt: 2d);

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
                }, row: 38, column: 6, width: OfficeLength.Pixels(640), height: OfficeLength.Pixels(360), title: "Scatter (Ranges)");

                sheet.AddBubbleChartFromRanges(new[] {
                    new ExcelChartSeriesRange("Bubbles", "A31:A33", "B31:B33", "D31:D33")
                }, row: 54, column: 6, width: OfficeLength.Pixels(640), height: OfficeLength.Pixels(360), title: "Bubble");

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
                var sheet = document.AddWorksheet("Summary");
                sheet.CellValues(new[] {
                    (1, 1, (object)"X"), (1, 2, (object)"Y"),
                    (2, 1, (object)1d), (2, 2, (object)2d),
                    (3, 1, (object)2d), (3, 2, (object)4d),
                    (4, 1, (object)3d), (4, 2, (object)3d)
                }, null);

                sheet.AddScatterChartFromRanges(ranges, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320), title: "Scatter");
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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] {
                        new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }, OfficeChartType.ColumnClustered, OfficeChartAxisGroup.Primary),
                        new ExcelChartSeries("Trend", new[] { 12d, 18d, 28d }, OfficeChartType.Line, OfficeChartAxisGroup.Secondary)
                    });

                sheet.AddChart(data, row: 1, column: 6, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Combo");
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
        public void Test_ExcelCharts_DuplicatePrimaryAxisReferenceDoesNotEscapeSnapshot() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.DuplicatePrimaryAxisReference.xlsx");
            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(new[] { "Q1", "Q2" }, new[] {
                    new ExcelChartSeries("Sales", new[] { 10d, 20d }, OfficeChartType.ColumnClustered, OfficeChartAxisGroup.Primary),
                    new ExcelChartSeries("Trend", new[] { 12d, 18d }, OfficeChartType.Line, OfficeChartAxisGroup.Secondary)
                });
                sheet.AddChart(data, row: 1, column: 6, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Combo");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                ChartPart chartPart = GetWorksheetPartWithCharts(spreadsheet).DrawingsPart!.ChartParts.First();
                C.PlotArea plot = chartPart.ChartSpace.GetFirstChild<C.Chart>()!.GetFirstChild<C.PlotArea>()!;
                C.BarChart primary = plot.GetFirstChild<C.BarChart>()!;
                uint primaryValueId = plot.Elements<C.ValueAxis>().Single(axis =>
                    axis.AxisPosition?.Val?.Value == C.AxisPositionValues.Left).AxisId!.Val!.Value;
                primary.Append(new C.AxisId { Val = primaryValueId });
                chartPart.ChartSpace.Save();
            }

            using var loaded = ExcelDocument.Load(filePath, new ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
            ExcelChart chart = Assert.Single(loaded.Sheets.Single(sheet => sheet.Name == "Summary").Charts);
            Assert.Null(Record.Exception(() => chart.TryGetSnapshot(out _)));
        }

        [Fact]
        public void Test_ExcelCharts_SeriesDataLabels_AndSecondaryAxisFormat() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SeriesLabels.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] {
                        new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }, OfficeChartType.ColumnClustered, OfficeChartAxisGroup.Primary),
                        new ExcelChartSeries("Trend", new[] { 12d, 18d, 28d }, OfficeChartType.Line, OfficeChartAxisGroup.Secondary)
                    });

                var chart = sheet.AddChart(data, row: 1, column: 6, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Combo");
                chart.SetValueAxisNumberFormat("0.00", sourceLinked: false, axisGroup: OfficeChartAxisGroup.Secondary)
                     .SetSeriesDataLabels(1, showValue: true, position: OfficeChartDataLabelPosition.Top, numberFormat: "0.0");
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
                var sheet = document.AddWorksheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"X"), (1, 2, "Y1"),
                    (2, 1, 1d), (2, 2, 2d),
                    (3, 1, 2d), (3, 2, 4d),
                    (4, 1, 3d), (4, 2, 6d)
                }, null);

                sheet.AddScatterChartFromRanges(new[] {
                    new ExcelChartSeriesRange("Series 1", "A2:A4", "B2:B4")
                }, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320), title: "Scatter Ranges");

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
                var sheet = document.AddWorksheet("Data");
                sheet.CellValues(new[] {
                    (1, 1, (object)"X"), (1, 2, "Y"), (1, 3, "Size"),
                    (2, 1, 1d), (2, 2, 2d), (2, 3, 4d),
                    (3, 1, 2d), (3, 2, 3d), (3, 3, 5d),
                    (4, 1, 3d), (4, 2, 4d), (4, 3, 6d)
                }, null);

                sheet.AddBubbleChartFromRanges(new[] {
                    new ExcelChartSeriesRange("Bubbles", "A2:A4", "B2:B4", "C2:C4")
                }, row: 8, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320), title: "Bubble Ranges");

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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2" },
                    new[] { new ExcelChartSeries("Sales", new[] { 1d, 2d }) });

                sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Styled");
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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2" },
                    new[] { new ExcelChartSeries("Sales", new[] { 1d, 2d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Labels");
                chart.SetSeriesDataLabels(0, showValue: true)
                     .SetSeriesDataLabelTextStyle(0, fontSize: OfficeLength.Points(12), bold: true, color: OfficeColor.Parse("FF0000"), fontName: "Calibri");
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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2" },
                    new[] { new ExcelChartSeries("Sales", new[] { 1d, 2d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Label Shapes");
                chart.SetSeriesDataLabels(0, showValue: true)
                     .SetSeriesDataLabelShapeStyle(0, fillColor: OfficeColor.Parse("FFFFCC"), lineColor: OfficeColor.Parse("000000"), lineWidth: OfficeLength.Points(1.5));
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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2" },
                    new[] { new ExcelChartSeries("Sales", new[] { 1d, 2d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.Pie, title: "Leader Lines");
                chart.SetSeriesDataLabels(0, showValue: true)
                     .SetSeriesDataLabelLeaderLines(0, showLeaderLines: true, lineColor: OfficeColor.Parse("000000"), lineWidth: OfficeLength.Points(1));
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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2" },
                    new[] { new ExcelChartSeries("Sales", new[] { 1d, 2d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Styled");
                chart.SetLegend(OfficeChartLegendPosition.Right)
                     .SetTitleTextStyle(fontSize: OfficeLength.Points(14), bold: true, color: OfficeColor.Parse("1F4E79"))
                     .SetLegendTextStyle(fontSize: OfficeLength.Points(9), italic: true, color: OfficeColor.Parse("404040"), fontName: "Calibri");
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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2" },
                    new[] { new ExcelChartSeries("Sales", new[] { 1d, 2d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Axis Styles");
                chart.SetCategoryAxisTitle("Quarter")
                     .SetValueAxisTitle("Revenue")
                     .SetCategoryAxisTitleTextStyle(fontSize: OfficeLength.Points(11), bold: true, color: OfficeColor.Parse("006100"))
                     .SetValueAxisTitleTextStyle(fontSize: OfficeLength.Points(11), bold: true, color: OfficeColor.Parse("006100"))
                     .SetCategoryAxisLabelTextStyle(fontSize: OfficeLength.Points(9), color: OfficeColor.Parse("404040"))
                     .SetValueAxisLabelTextStyle(fontSize: OfficeLength.Points(9), italic: true, fontName: "Calibri");
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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Gridlines");
                chart.SetValueAxisGridlines(showMajor: true, showMinor: true, lineColor: OfficeColor.Parse("C0C0C0"), lineWidth: OfficeLength.Points(0.75))
                     .SetCategoryAxisLabelRotation(45)
                     .SetValueAxisTickLabelPosition(OfficeChartAxisTickLabelPosition.Low);
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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Axis Crossing");
                chart.SetValueAxisCrossing(OfficeChartAxisCrossingPosition.Maximum)
                     .SetCategoryAxisCrossing(OfficeChartAxisCrossingPosition.Minimum)
                     .SetValueAxisCrossBetween(OfficeChartAxisCrossBetween.Between)
                     .SetValueAxisDisplayUnits(OfficeChartDisplayUnit.Thousands, "Thousands USD", showLabel: true);
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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Axis Crossing At");
                chart.SetValueAxisCrossing(OfficeChartAxisCrossingPosition.AutoZero, crossesAt: 2.5);
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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Axis Scale");
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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "1", "2", "3", "4" },
                    new[] { new ExcelChartSeries("Points", new[] { 2d, 4d, 3d, 5d }, OfficeChartType.Scatter) });

                var chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.Scatter, title: "Scatter Axis");
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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "1", "2", "3", "4" },
                    new[] { new ExcelChartSeries("Points", new[] { 2d, 4d, 3d, 5d }, OfficeChartType.Scatter) });

                var chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.Scatter, title: "Scatter Axis");
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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "1", "2", "3" },
                    new[] { new ExcelChartSeries("Points", new[] { 2d, 4d, 3d }, OfficeChartType.Scatter) });

                var chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.Scatter, title: "Scatter Crossing");
                chart.SetScatterXAxisScale(minimum: 1, maximum: 10, logScale: true);

                Assert.Throws<ArgumentOutOfRangeException>(() =>
                    chart.SetScatterXAxisCrossing(OfficeChartAxisCrossingPosition.AutoZero, crossesAt: 0));
            }
        }

        [Fact]
        public void Test_ExcelCharts_ScatterYAxisCrossing() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.ScatterYAxisCrossing.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "1", "2", "3" },
                    new[] { new ExcelChartSeries("Points", new[] { 2d, 4d, 3d }, OfficeChartType.Scatter) });

                var chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.Scatter, title: "Scatter Crossing");
                chart.SetScatterYAxisCrossing(OfficeChartAxisCrossingPosition.Minimum, crossesAt: 2d);
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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Label Template");
                var template = new ExcelChartDataLabelTemplate {
                    ShowValue = true,
                    Position = OfficeChartDataLabelPosition.Top,
                    NumberFormat = "0.0",
                    FontSize = OfficeLength.Points(9),
                    TextColor = OfficeColor.Parse("404040"),
                    Separator = " - ",
                    FillColor = OfficeColor.Parse("FFFFFF"),
                    LineColor = OfficeColor.Parse("000000"),
                    LineWidth = OfficeLength.Points(0.5)
                };
                chart.SetSeriesDataLabelTemplate(0, template)
                     .SetSeriesDataLabelForPoint(0, 1, showValue: true, position: OfficeChartDataLabelPosition.OutsideEnd,
                        numberFormat: "0.00")
                     .SetSeriesDataLabelSeparatorForPoint(0, 1, " | ")
                     .SetSeriesDataLabelTextStyleForPoint(0, 1, fontSize: OfficeLength.Points(11), bold: true, color: OfficeColor.Parse("FF0000"))
                     .SetSeriesDataLabelShapeStyleForPoint(0, 1, fillColor: OfficeColor.Parse("FFFFCC"), lineColor: OfficeColor.Parse("000000"),
                        lineWidth: OfficeLength.Points(1));
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
        public void Test_ExcelCharts_ImageExportStylesFollowSeriesIndexNotXmlOrder() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SeriesStyleIndexOrder.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2" },
                    new[] {
                        new ExcelChartSeries("First", new[] { 10d, 20d }),
                        new ExcelChartSeries("Second", new[] { 30d, 40d })
                    });

                ExcelChart chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Series Style Index");
                chart.SetSeriesFillColor(0, OfficeColor.Parse("FF0000"))
                     .SetSeriesFillColor(1, OfficeColor.Parse("0000FF"));
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = GetWorksheetPartWithCharts(spreadsheet);
                ChartPart chartPart = worksheetPart.DrawingsPart!.ChartParts.First();
                C.BarChartSeries[] series = chartPart.ChartSpace.GetFirstChild<C.Chart>()!
                    .GetFirstChild<C.PlotArea>()!
                    .GetFirstChild<C.BarChart>()!
                    .Elements<C.BarChartSeries>()
                    .ToArray();
                Assert.Equal(2, series.Length);
                series[0].GetFirstChild<C.Index>()!.Val = 1U;
                series[0].GetFirstChild<C.Order>()!.Val = 1U;
                series[1].GetFirstChild<C.Index>()!.Val = 0U;
                series[1].GetFirstChild<C.Order>()!.Val = 0U;
                chartPart.ChartSpace.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                ExcelChart chart = Assert.Single(document.Sheets.Single(sheet => sheet.Name == "Summary").Charts);
                Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));
                Assert.Equal("0000FF", snapshot.Data.Series[0].SeriesColorArgb);
                Assert.Equal("FF0000", snapshot.Data.Series[1].SeriesColorArgb);
            }
        }

        [Fact]
        public void Test_ExcelCharts_SnapshotPreservesSparseCachedValuePointIndexes() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.SparseCachedValues.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }) });

                sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Sparse cache");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = GetWorksheetPartWithCharts(spreadsheet);
                ChartPart chartPart = worksheetPart.DrawingsPart!.ChartParts.First();
                C.NumberReference valueReference = chartPart.ChartSpace.GetFirstChild<C.Chart>()!
                    .GetFirstChild<C.PlotArea>()!
                    .GetFirstChild<C.BarChart>()!
                    .Elements<C.BarChartSeries>()
                    .First()
                    .GetFirstChild<C.Values>()!
                    .GetFirstChild<C.NumberReference>()!;
                valueReference.Formula!.Text = "Summary!#REF!";
                C.NumberingCache cache = valueReference.GetFirstChild<C.NumberingCache>()!;
                cache.PointCount!.Val = 3U;
                cache.RemoveAllChildren<C.NumericPoint>();
                cache.Append(
                    new C.NumericPoint(new C.NumericValue("10")) { Index = 0U },
                    new C.NumericPoint(new C.NumericValue("30")) { Index = 2U });
                chartPart.ChartSpace.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                ExcelChart chart = Assert.Single(document.Sheets.Single(sheet => sheet.Name == "Summary").Charts);
                Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));

                double[] values = Assert.Single(snapshot.Data.Series).Values.ToArray();
                Assert.Equal(3, values.Length);
                Assert.Equal(10d, values[0]);
                Assert.True(double.IsNaN(values[1]));
                Assert.Equal(30d, values[2]);
            }
        }

        [Fact]
        public void Test_ExcelCharts_HorizontalSnapshotReadsNonAdjacentValueRowFromSeriesFormula() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.HorizontalNonAdjacentValues.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Summary");
                sheet.CellValue(1, 2, "Q1");
                sheet.CellValue(1, 3, "Q2");
                sheet.CellValue(1, 4, "Q3");
                sheet.CellValue(2, 1, "Wrong row");
                sheet.CellValue(2, 2, 999);
                sheet.CellValue(2, 3, 999);
                sheet.CellValue(2, 4, 999);
                sheet.CellValue(3, 1, "Actual");
                sheet.CellValue(3, 2, 10);
                sheet.CellValue(3, 3, 20);
                sheet.CellValue(3, 4, 30);

                var range = new ExcelChartDataRange(
                    "Summary",
                    startRow: 1,
                    startColumn: 1,
                    categoryCount: 3,
                    seriesCount: 1,
                    hasHeaderRow: true,
                    ExcelChartDataOrientation.Horizontal);
                sheet.AddChart(range, row: 5, column: 1, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Non-adjacent row");
                document.Save();
            }

            using (var spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                WorksheetPart worksheetPart = GetWorksheetPartWithCharts(spreadsheet);
                ChartPart chartPart = worksheetPart.DrawingsPart!.ChartParts.First();
                C.BarChartSeries series = chartPart.ChartSpace.GetFirstChild<C.Chart>()!
                    .GetFirstChild<C.PlotArea>()!
                    .GetFirstChild<C.BarChart>()!
                    .Elements<C.BarChartSeries>()
                    .First();

                series.GetFirstChild<C.SeriesText>()!
                    .GetFirstChild<C.StringReference>()!
                    .Formula!.Text = "Summary!$A$3";
                series.GetFirstChild<C.Values>()!
                    .GetFirstChild<C.NumberReference>()!
                    .Formula!.Text = "Summary!$B$3:$D$3";
                chartPart.ChartSpace.Save();
            }

            using (var document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                ExcelChart chart = Assert.Single(document.Sheets.Single(sheet => sheet.Name == "Summary").Charts);
                Assert.True(chart.TryGetSnapshot(out ExcelChartInfo snapshot));

                Assert.Equal(new[] { "Q1", "Q2", "Q3" }, snapshot.Data.Categories);
                ExcelChartSeries series = Assert.Single(snapshot.Data.Series);
                Assert.Equal("Actual", series.Name);
                Assert.Equal(new[] { 10d, 20d, 30d }, series.Values);
            }
        }

        [Fact]
        public void Test_ExcelCharts_SeriesTrendline() {
            string filePath = Path.Combine(_directoryWithFiles, "ExcelCharts.Trendline.xlsx");

            using (var document = ExcelDocument.Create(filePath)) {
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2", "Q3" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d, 30d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.Line, title: "Trendline");
                chart.SetSeriesTrendline(0, OfficeChartTrendlineType.Polynomial, order: 2,
                    displayEquation: true, displayRSquared: true, lineColor: OfficeColor.Parse("FF0000"), lineWidth: OfficeLength.Points(1.5));
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
                var sheet = document.AddWorksheet("Summary");
                var data = new ExcelChartData(
                    new[] { "Q1", "Q2" },
                    new[] { new ExcelChartSeries("Sales", new[] { 10d, 20d }) });

                var chart = sheet.AddChart(data, row: 1, column: 4, width: OfficeLength.Pixels(480), height: OfficeLength.Pixels(320),
                    type: OfficeChartType.ColumnClustered, title: "Area Style");
                chart.SetChartAreaStyle(fillColor: OfficeColor.Parse("F2F2F2"), lineColor: OfficeColor.Parse("404040"), lineWidth: OfficeLength.Points(1.25))
                     .SetPlotAreaStyle(fillColor: OfficeColor.Parse("FFFFFF"), lineColor: OfficeColor.Parse("00B0F0"), lineWidth: OfficeLength.Points(0.5));
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