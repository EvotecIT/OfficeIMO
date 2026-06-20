using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using System.Globalization;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Excel {

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Exports_Worksheet_Chart_Snapshots() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfCharts.xlsx");

        byte[] bytes;
        byte[] disabledBytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Charts")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Category");
            sheet.Cell(1, 2, "Actual");
            sheet.Cell(1, 3, "Target");
            sheet.Cell(2, 1, "Jan");
            sheet.Cell(2, 2, 12);
            sheet.Cell(2, 3, 10);
            sheet.Cell(3, 1, "Feb");
            sheet.Cell(3, 2, 18);
            sheet.Cell(3, 3, 16);
            sheet.Cell(4, 1, "Mar");
            sheet.Cell(4, 2, 24);
            sheet.Cell(4, 3, 20);
            sheet.AddChartFromRange("A1:C4", row: 1, column: 5, widthPixels: 360, heightPixels: 220, type: ExcelChartType.ColumnClustered, title: "Revenue Chart");

            ExcelChart chart = Assert.Single(sheet.Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal("Revenue Chart", snapshot.Title);
            Assert.Equal(ExcelChartType.ColumnClustered, snapshot.ChartType);
            Assert.Equal(3, snapshot.Data.Categories.Count);
            Assert.Equal(2, snapshot.Data.Series.Count);
            MethodInfo createSnapshot = typeof(ExcelPdfConverterExtensions).GetMethod("CreateOfficeChartSnapshot", BindingFlags.NonPublic | BindingFlags.Static)!;
            var sharedSnapshot = Assert.IsType<OfficeChartSnapshot>(createSnapshot.Invoke(null, new object[] { snapshot, new ExcelPdfSaveOptions() }));
            Assert.Equal("Revenue Chart", sharedSnapshot.Title);

            document.Save(false);

            var options = new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(480, 360),
                Margins = PdfCore.PageMargins.Uniform(24)
            };
            bytes = document.SaveAsPdf(options);
            disabledBytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                UseWorksheetCharts = false,
                PageSize = new PdfCore.PageSize(480, 360),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Revenue Chart", text);
        Assert.Single(Regex.Matches(text, "Revenue Chart").Cast<Match>());
        Assert.Contains("Actual", text);
        Assert.Contains("Target", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.122 0.306 0.475 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.184 0.435 0.243 rg", rawPdf, StringComparison.Ordinal);

        using PdfPigDocument disabledPdf = PdfPigDocument.Open(new MemoryStream(disabledBytes));
        Assert.DoesNotContain("Revenue Chart", disabledPdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Applies_Shared_Chart_Style_Options() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfStyledCharts.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Charts")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Category");
            sheet.Cell(1, 2, "Actual");
            sheet.Cell(1, 3, "Target");
            sheet.Cell(2, 1, "Jan");
            sheet.Cell(2, 2, 12);
            sheet.Cell(2, 3, 10);
            sheet.Cell(3, 1, "Feb");
            sheet.Cell(3, 2, 18);
            sheet.Cell(3, 3, 16);
            sheet.Cell(4, 1, "Mar");
            sheet.Cell(4, 2, 24);
            sheet.Cell(4, 3, 20);
            sheet.AddChartFromRange("A1:C4", row: 1, column: 5, widthPixels: 360, heightPixels: 220, type: ExcelChartType.ColumnClustered, title: "Styled Excel Chart");

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(480, 360),
                Margins = PdfCore.PageMargins.Uniform(24),
                ChartStyle = new OfficeChartStyle(
                    palette: new[] {
                        OfficeColor.FromRgb(18, 52, 86),
                        OfficeColor.FromRgb(120, 40, 160)
                    },
                    backgroundColor: OfficeColor.FromRgb(242, 248, 255),
                    titleColor: OfficeColor.FromRgb(200, 10, 10))
            });
        }

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.071 0.204 0.337 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.471 0.157 0.627 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.949 0.973 1 rg", rawPdf, StringComparison.Ordinal);

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Styled Excel Chart", text);
        Assert.Contains("Actual", text);
        Assert.Contains("Target", text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Warns_When_Shared_Chart_Quality_Preflight_Finds_Issues() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfChartQualityWarnings.xlsx");

        byte[] bytes;
        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            HeaderRowCount = 1,
            PageSize = new PdfCore.PageSize(480, 360),
            Margins = PdfCore.PageMargins.Uniform(24),
            ChartLayout = new OfficeChartLayout(maximumCategoryAxisLabels: 12, preventLabelOverlap: false)
        };

        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Charts")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Month");
            sheet.Cell(1, 2, "Actual");
            for (int row = 2; row <= 13; row++) {
                sheet.Cell(row, 1, "M" + (row - 1).ToString("00", CultureInfo.InvariantCulture));
                sheet.Cell(row, 2, row * 3);
            }

            sheet.AddChartFromRange("A1:B13", row: 1, column: 4, widthPixels: 300, heightPixels: 180, type: ExcelChartType.Line, title: "Dense Month Chart");
            document.Save(false);

            bytes = document.SaveAsPdf(options);
        }

        ExcelPdfExportWarning warning = Assert.Single(options.Warnings, item => item.Feature == "chart-quality");
        Assert.Equal("Charts", warning.SheetName);
        Assert.Contains("Dense Month Chart", warning.Message);
        Assert.Contains("TextOverlap", warning.Message);

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Contains("Dense Month Chart", pdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Rejects_Empty_Chart_Snapshots_Before_Shared_Rendering() {
        var data = new ExcelChartData(Array.Empty<string>(), new[] {
            new ExcelChartSeries("Empty", Array.Empty<double>())
        });
        var snapshot = new ExcelChartSnapshot(
            "EmptyChart",
            "Empty Chart",
            ExcelChartType.Line,
            data,
            rowIndex: 1,
            columnIndex: 1,
            widthPixels: 320,
            heightPixels: 180);

        MethodInfo method = typeof(ExcelPdfConverterExtensions).GetMethod("HasRenderableChartData", BindingFlags.NonPublic | BindingFlags.Static)!;

        Assert.False((bool)method.Invoke(null, new object[] { snapshot })!);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Exports_Pie_And_Doughnut_Chart_Snapshots() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfPieDoughnutCharts.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Charts")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Category");
            sheet.Cell(1, 2, "Control Share");
            sheet.Cell(2, 1, "Compliant");
            sheet.Cell(2, 2, 62);
            sheet.Cell(3, 1, "Partial");
            sheet.Cell(3, 2, 21);
            sheet.Cell(4, 1, "Non-compliant");
            sheet.Cell(4, 2, 11);
            sheet.Cell(5, 1, "Not assessed");
            sheet.Cell(5, 2, 6);
            sheet.AddChartFromRange("A1:B5", row: 1, column: 4, widthPixels: 280, heightPixels: 180, type: ExcelChartType.Pie, title: "Control Status Pie");
            sheet.AddChartFromRange("A1:B5", row: 12, column: 4, widthPixels: 280, heightPixels: 180, type: ExcelChartType.Doughnut, title: "Control Status Doughnut");

            List<ExcelChart> charts = sheet.Charts.ToList();
            Assert.Equal(2, charts.Count);
            Assert.All(charts, chart => Assert.True(chart.TryGetSnapshot(out _)));
            Assert.Equal(ExcelChartType.Pie, charts[0].ChartType);
            Assert.Equal(ExcelChartType.Doughnut, charts[1].ChartType);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(480, 520),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("Control Status Pie", text);
        Assert.Contains("Control Status Doughnut", text);
        Assert.Contains("Compliant", text);
        Assert.Contains("Non-compliant", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.122 0.306 0.475 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.184 0.435 0.243 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.722 0.353 0.137 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Exports_Area_Chart_Snapshots() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfAreaCharts.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Charts")) {
            ExcelSheet sheet = document.Sheets[0];
            sheet.Cell(1, 1, "Quarter");
            sheet.Cell(1, 2, "Services");
            sheet.Cell(1, 3, "Licenses");
            sheet.Cell(2, 1, "Q1");
            sheet.Cell(2, 2, 36);
            sheet.Cell(2, 3, 19);
            sheet.Cell(3, 1, "Q2");
            sheet.Cell(3, 2, 44);
            sheet.Cell(3, 3, 25);
            sheet.Cell(4, 1, "Q3");
            sheet.Cell(4, 2, 50);
            sheet.Cell(4, 3, 31);
            sheet.Cell(5, 1, "Q4");
            sheet.Cell(5, 2, 54);
            sheet.Cell(5, 3, 34);
            sheet.AddChartFromRange("A1:C5", row: 1, column: 5, widthPixels: 360, heightPixels: 220, type: ExcelChartType.Area, title: "Revenue Area");
            sheet.AddChartFromRange("A1:C5", row: 14, column: 5, widthPixels: 360, heightPixels: 220, type: ExcelChartType.AreaStacked100, title: "Revenue Mix Area");

            List<ExcelChart> charts = sheet.Charts.ToList();
            Assert.Equal(2, charts.Count);
            Assert.All(charts, chart => Assert.True(chart.TryGetSnapshot(out _)));
            Assert.Equal(ExcelChartType.Area, charts[0].ChartType);
            Assert.Equal(ExcelChartType.AreaStacked100, charts[1].ChartType);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(520, 620),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Concat(pdf.GetPages().Select(page => page.Text));
        Assert.Contains("Revenue Area", text);
        Assert.Contains("Revenue Mix Area", text);
        Assert.Contains("Services", text);
        Assert.Contains("Licenses", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.122 0.306 0.475 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.184 0.435 0.243 RG", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Preserves_Negative_Line_Chart_Values() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfNegativeLineChart.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Charts")) {
            ExcelSheet sheet = document.Sheets[0];
            var data = new ExcelChartData(
                new[] { "Low", "Zero", "High" },
                new[] {
                    new ExcelChartSeries("Profit", new[] { -10D, 0D, 10D }, ExcelChartType.Line)
                });

            sheet.AddChart(data, row: 1, column: 5, widthPixels: 360, heightPixels: 220, type: ExcelChartType.Line, title: "Profit Trend");

            ExcelChart chart = Assert.Single(sheet.Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal(ExcelChartType.Line, snapshot.ChartType);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(480, 360),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Profit Trend", text);
        Assert.Contains("Profit", text);

        MethodInfo rangeMethod = typeof(OfficeChartDrawingRenderer).GetMethod("GetFiniteSeriesRange", BindingFlags.NonPublic | BindingFlags.Static)!;
        object range = rangeMethod.Invoke(null, new object[] { new List<OfficeChartSeries> { new OfficeChartSeries("Profit", new[] { -10D, 0D, 10D }) } })!;
        double min = (double)range.GetType().GetProperty("Min", BindingFlags.Instance | BindingFlags.Public)!.GetValue(range)!;
        double max = (double)range.GetType().GetProperty("Max", BindingFlags.Instance | BindingFlags.Public)!.GetValue(range)!;

        MethodInfo plotYMethod = typeof(OfficeChartDrawingRenderer)
            .GetMethods(BindingFlags.NonPublic | BindingFlags.Static)
            .Single(method => method.Name == "ToPlotY" && method.GetParameters().Length == 5);
        double negativeY = (double)plotYMethod.Invoke(null, new object[] { -10D, min, max, 0D, 100D })!;
        double zeroY = (double)plotYMethod.Invoke(null, new object[] { 0D, min, max, 0D, 100D })!;
        double positiveY = (double)plotYMethod.Invoke(null, new object[] { 10D, min, max, 0D, 100D })!;

        Assert.Equal(-10D, min);
        Assert.Equal(10D, max);
        Assert.True(negativeY > zeroY && zeroY > positiveY, "Expected negative, zero, and positive line chart values to map to separate vertical positions.");
    }

    [Fact]
    public void SharedChartRenderer_UsesDisplayedAxisRangeForPositiveLineCharts() {
        var snapshot = new OfficeChartSnapshot(
            "PositiveLine",
            null,
            OfficeChartKind.Line,
            new OfficeChartData(
                new[] { "Low", "High" },
                new[] { new OfficeChartSeries("Value", new[] { 10D, 20D }) }),
            360D,
            220D);

        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot);
        OfficeDrawingShape firstMarker = drawing.Shapes
            .Where(shape => shape.Shape.Kind == OfficeShapeKind.Ellipse && Math.Abs(shape.Shape.Width - 4D) < 0.01D)
            .OrderBy(shape => shape.X)
            .First();

        Assert.InRange(firstMarker.Y, 90D, 110D);
    }

    [Fact]
    public void SharedChartRenderer_UsesDisplayedAxisRangeForPositiveAreaCharts() {
        var snapshot = new OfficeChartSnapshot(
            "PositiveArea",
            null,
            OfficeChartKind.Area,
            new OfficeChartData(
                new[] { "Low", "High" },
                new[] { new OfficeChartSeries("Value", new[] { 10D, 20D }) }),
            360D,
            220D);

        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot);

        Assert.Contains(drawing.Shapes, shape => shape.Shape.Kind == OfficeShapeKind.Polygon && shape.Y >= 0D && shape.Y < drawing.Height);
    }

    [Fact]
    public void SharedChartRenderer_RendersDoughnutRingForEachSeries() {
        var snapshot = new OfficeChartSnapshot(
            "MultiDoughnut",
            null,
            OfficeChartKind.Doughnut,
            new OfficeChartData(
                new[] { "A", "B", "C" },
                new[] {
                    new OfficeChartSeries("Outer", new[] { 3D, 2D, 1D }),
                    new OfficeChartSeries("Inner", new[] { 1D, 2D, 3D })
                }),
            360D,
            220D);

        OfficeDrawing drawing = OfficeChartDrawingRenderer.Render(snapshot);
        int slices = drawing.Shapes.Count(shape => shape.Shape.Kind == OfficeShapeKind.Polygon);

        Assert.Equal(6, slices);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_WarnsAndSkipsMixedSeriesChartTypes() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfMixedSeriesChart.xlsx");

        byte[] bytes;
        var options = new ExcelPdfSaveOptions {
            IncludeSheetHeadings = false,
            HeaderRowCount = 1,
            PageSize = new PdfCore.PageSize(480, 360),
            Margins = PdfCore.PageMargins.Uniform(24)
        };
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Charts")) {
            ExcelSheet sheet = document.Sheets[0];
            var data = new ExcelChartData(
                new[] { "Q1", "Q2", "Q3" },
                new[] {
                    new ExcelChartSeries("Sales", new[] { 12D, 18D, 24D }, ExcelChartType.ColumnClustered, ExcelChartAxisGroup.Primary),
                    new ExcelChartSeries("Trend", new[] { 10D, 16D, 22D }, ExcelChartType.Line, ExcelChartAxisGroup.Secondary)
                });

            sheet.AddChart(data, row: 1, column: 5, widthPixels: 360, heightPixels: 220, type: ExcelChartType.ColumnClustered, title: "Sales vs Trend");
            document.Save(false);

            bytes = document.SaveAsPdf(options);
        }

        ExcelPdfExportWarning warning = Assert.Single(options.Warnings, item => item.Feature == "WorksheetChart");
        Assert.Contains("mixed per-series chart types", warning.Message, StringComparison.Ordinal);

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.DoesNotContain("Sales vs Trend", pdf.GetPage(1).Text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Maps_Stacked_Line_Charts_To_Shared_Stacked_Line_Renderer() {
        MethodInfo method = typeof(ExcelPdfConverterExtensions).GetMethod("TryMapChartKind", BindingFlags.NonPublic | BindingFlags.Static)!;

        object?[] stackedArgs = { ExcelChartType.LineStacked, null };
        object?[] percentArgs = { ExcelChartType.LineStacked100, null };

        Assert.True((bool)method.Invoke(null, stackedArgs)!);
        Assert.True((bool)method.Invoke(null, percentArgs)!);
        Assert.Equal(OfficeChartKind.LineStacked, stackedArgs[1]);
        Assert.Equal(OfficeChartKind.LineStacked100, percentArgs[1]);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Exports_Negative_Area_Chart_Values() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfNegativeAreaChart.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Charts")) {
            ExcelSheet sheet = document.Sheets[0];
            var data = new ExcelChartData(
                new[] { "Q1", "Q2", "Q3" },
                new[] {
                    new ExcelChartSeries("Delta", new[] { -6D, 0D, 9D }, ExcelChartType.Area)
                });

            sheet.AddChart(data, row: 1, column: 5, widthPixels: 360, heightPixels: 220, type: ExcelChartType.Area, title: "Delta Area");
            ExcelChart chart = Assert.Single(sheet.Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal(ExcelChartType.Area, snapshot.ChartType);

            document.Save(false);
            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(480, 360),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Delta Area", text);
        Assert.Contains("Delta", text);

        MethodInfo rangeMethod = typeof(OfficeChartDrawingRenderer).GetMethod("GetFiniteSeriesRange", BindingFlags.NonPublic | BindingFlags.Static)!;
        object range = rangeMethod.Invoke(null, new object[] { new List<OfficeChartSeries> { new OfficeChartSeries("Delta", new[] { -6D, 0D, 9D }) } })!;
        Assert.Equal(-6D, (double)range.GetType().GetProperty("Min", BindingFlags.Instance | BindingFlags.Public)!.GetValue(range)!);
        Assert.Equal(9D, (double)range.GetType().GetProperty("Max", BindingFlags.Instance | BindingFlags.Public)!.GetValue(range)!);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Exports_Scatter_Chart_Snapshots() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfScatterCharts.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Charts")) {
            ExcelSheet sheet = document.Sheets[0];
            var data = new ExcelChartData(
                new[] { "1", "2", "4", "8", "16" },
                new[] {
                    new ExcelChartSeries("Latency", new[] { 9D, 7D, 5.5D, 4.2D, 3.8D }, ExcelChartType.Scatter),
                    new ExcelChartSeries("Throughput", new[] { 2D, 3.5D, 6D, 7.5D, 9D }, ExcelChartType.Scatter)
                });

            sheet.AddChart(data, row: 1, column: 5, widthPixels: 360, heightPixels: 220, type: ExcelChartType.Scatter, title: "Scale Scatter");

            ExcelChart chart = Assert.Single(sheet.Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal(ExcelChartType.Scatter, snapshot.ChartType);
            Assert.Equal(5, snapshot.Data.Categories.Count);
            Assert.Equal(2, snapshot.Data.Series.Count);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(480, 360),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Scale Scatter", text);
        Assert.Contains("Latency", text);
        Assert.Contains("Throughput", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.122 0.306 0.475 RG", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.184 0.435 0.243 RG", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Exports_Radar_Chart_Snapshots() {
        string workbookPath = Path.Combine(_directoryWithFiles, "ExcelPdfRadarCharts.xlsx");

        byte[] bytes;
        using (ExcelDocument document = ExcelDocument.Create(workbookPath, "Charts")) {
            ExcelSheet sheet = document.Sheets[0];
            var data = new ExcelChartData(
                new[] { "Quality", "Speed", "Cost", "Coverage", "Risk" },
                new[] {
                    new ExcelChartSeries("Current", new[] { 7D, 6D, 5D, 8D, 4D }, ExcelChartType.Radar),
                    new ExcelChartSeries("Target", new[] { 9D, 8D, 7D, 9D, 6D }, ExcelChartType.Radar)
                });

            sheet.AddChart(data, row: 1, column: 5, widthPixels: 360, heightPixels: 220, type: ExcelChartType.Radar, title: "Capability Radar");

            ExcelChart chart = Assert.Single(sheet.Charts);
            Assert.True(chart.TryGetSnapshot(out ExcelChartSnapshot snapshot));
            Assert.Equal(ExcelChartType.Radar, snapshot.ChartType);
            Assert.Equal(5, snapshot.Data.Categories.Count);
            Assert.Equal(2, snapshot.Data.Series.Count);

            document.Save(false);

            bytes = document.SaveAsPdf(new ExcelPdfSaveOptions {
                IncludeSheetHeadings = false,
                HeaderRowCount = 1,
                PageSize = new PdfCore.PageSize(480, 360),
                Margins = PdfCore.PageMargins.Uniform(24)
            });
        }

        using PdfPigDocument pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = pdf.GetPage(1).Text;
        Assert.Contains("Capability Radar", text);
        Assert.Contains("Current", text);
        Assert.Contains("Target", text);

        string rawPdf = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.122 0.306 0.475 rg", rawPdf, StringComparison.Ordinal);
        Assert.Contains("0.184 0.435 0.243 rg", rawPdf, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_ExcelWorkbook_Preserves_Negative_Radar_Chart_Values() {
        var series = new List<ExcelChartSeries> {
            new ExcelChartSeries("Delta", new[] { -10D, -2D, 0D, 10D }, ExcelChartType.Radar)
        };

        MethodInfo rangeMethod = typeof(OfficeChartDrawingRenderer).GetMethod("GetRadarValueRange", BindingFlags.NonPublic | BindingFlags.Static)!;
        object range = rangeMethod.Invoke(null, new object[] { series.Select(item => new OfficeChartSeries(item.Name, item.Values)).ToList() })!;
        double min = (double)range.GetType().GetProperty("Min", BindingFlags.Instance | BindingFlags.Public)!.GetValue(range)!;
        double max = (double)range.GetType().GetProperty("Max", BindingFlags.Instance | BindingFlags.Public)!.GetValue(range)!;

        MethodInfo ratioMethod = typeof(OfficeChartDrawingRenderer).GetMethod("ToRadarRadiusRatio", BindingFlags.NonPublic | BindingFlags.Static)!;
        double negativeRatio = (double)ratioMethod.Invoke(null, new object[] { -2D, min, max })!;
        double zeroRatio = (double)ratioMethod.Invoke(null, new object[] { 0D, min, max })!;
        double positiveRatio = (double)ratioMethod.Invoke(null, new object[] { 10D, min, max })!;

        Assert.Equal(-10D, min);
        Assert.Equal(10D, max);
        Assert.True(negativeRatio > 0D, "Expected below-zero radar values inside the axis range to render away from the center.");
        Assert.True(negativeRatio < zeroRatio && zeroRatio < positiveRatio, "Expected signed radar values to keep their axis order.");
    }

}
