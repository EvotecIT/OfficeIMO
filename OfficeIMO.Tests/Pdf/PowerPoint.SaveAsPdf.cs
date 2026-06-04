using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Drawing;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PowerPointSaveAsPdfTests {
    [Fact]
    public void SaveAsPdf_PowerPointPresentation_MapsSlideSizeTextShapeAndPictureToCanvasPdf() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 180);
        PowerPointSlide slide = presentation.Slides[0];

        PowerPointAutoShape panel = slide.AddRectanglePoints(20, 24, 120, 48);
        panel.FillColor = "EAF4FF";
        panel.OutlineColor = "1E5A96";
        panel.OutlineWidthPoints = 1.5D;

        PowerPointTextBox textBox = slide.AddTextBoxPoints("Premium Slide", 32, 36, 150, 36);
        textBox.FillColor = "FFFFFF";
        textBox.OutlineColor = "94A3B8";
        textBox.FontSize = 14;
        textBox.Color = "123456";
        textBox.Rotation = 0D;

        slide.AddPicture(new MemoryStream(CreateMinimalRgbPng()), OfficeIMO.PowerPoint.ImagePartType.Png, PowerPointUnits.FromPoints(210), PowerPointUnits.FromPoints(42), PowerPointUnits.FromPoints(50), PowerPointUnits.FromPoints(30));

        byte[] bytes = presentation.SaveAsPdf();
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);

        Assert.Equal(1, info.PageCount);
        PdfCore.PdfPageInfo page = Assert.Single(info.Pages);
        Assert.Equal(320D, page.Width);
        Assert.Equal(180D, page.Height);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Premium Slide", text, StringComparison.Ordinal);

        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("20 108 120 48 re", raw, StringComparison.Ordinal);
        Assert.Contains("/Im1 Do", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesTextRunHyperlinks() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointTextBox textBox = presentation.Slides[0].AddTextBoxPoints(string.Empty, 24, 32, 150, 38);
        textBox.SetParagraphs(new[] { string.Empty });
        PowerPointTextRun run = textBox.Paragraphs[0].AddRun("OfficeIMO");
        run.SetHyperlink("https://officeimo.net/");

        byte[] bytes = presentation.SaveAsPdf();
        PdfCore.PdfDocumentInfo info = PdfCore.PdfInspector.Inspect(bytes);

        Assert.Equal(new[] { "https://officeimo.net/" }, info.LinkUris);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_MapsTextBoxVerticalAlignmentToSharedCanvasTextBox() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(360, 200);
        PowerPointSlide slide = presentation.Slides[0];

        PowerPointTextBox top = slide.AddTextBoxPoints("TopPpt", 20, 30, 90, 90);
        top.TextVerticalAlignment = TextAnchoringTypeValues.Top;
        top.FontSize = 10;
        top.FillColor = "FFFFFF";
        top.FillTransparency = 100;

        PowerPointTextBox middle = slide.AddTextBoxPoints("MiddlePpt", 130, 30, 90, 90);
        middle.TextVerticalAlignment = TextAnchoringTypeValues.Center;
        middle.FontSize = 10;
        middle.FillColor = "FFFFFF";
        middle.FillTransparency = 100;

        PowerPointTextBox bottom = slide.AddTextBoxPoints("BottomPpt", 240, 30, 90, 90);
        bottom.TextVerticalAlignment = TextAnchoringTypeValues.Bottom;
        bottom.FontSize = 10;
        bottom.FillColor = "FFFFFF";
        bottom.FillTransparency = 100;

        byte[] bytes = presentation.SaveAsPdf();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        var page = pdf.GetPage(1);
        double topY = FindWordStartY(page, "TopPpt");
        double middleY = FindWordStartY(page, "MiddlePpt");
        double bottomY = FindWordStartY(page, "BottomPpt");

        Assert.True(topY > middleY + 30D, $"Expected PowerPoint center-anchored text to render lower than top-anchored text. Top: {topY:0.##}, middle: {middleY:0.##}.");
        Assert.True(middleY > bottomY + 30D, $"Expected PowerPoint bottom-anchored text to render lower than center-anchored text. Middle: {middleY:0.##}, bottom: {bottomY:0.##}.");
    }

    [Fact]
    public void ToPdfDocument_PowerPointPresentation_WarnsForUnsupportedShapes() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.Slides[0].AddShape(ShapeTypeValues.Triangle, PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(20), PowerPointUnits.FromPoints(50), PowerPointUnits.FromPoints(40));
        var options = new PowerPointPdfSaveOptions();

        presentation.ToPdfDocument(options).ToBytes();

        PowerPointPdfExportWarning warning = Assert.Single(options.Warnings);
        Assert.Equal(1, warning.SlideNumber);
        Assert.Equal("unsupported-auto-shape", warning.Code);
    }

    [Fact]
    public void ToPdfDocument_PowerPointPresentation_WarnsAndSkipsOffSlideShapes() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        presentation.Slides[0].AddRectanglePoints(-10, 20, 80, 40);
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.SaveAsPdf(options);

        PowerPointPdfExportWarning warning = Assert.Single(options.Warnings);
        Assert.Equal(1, warning.SlideNumber);
        Assert.Equal("invalid-shape-bounds", warning.Code);
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_DegradesTinyTextBoxMarginsInsteadOfThrowing() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(160, 100);
        PowerPointTextBox textBox = presentation.Slides[0].AddTextBoxPoints("Tiny", 20, 20, 6, 6);
        textBox.FillTransparency = 100;
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.SaveAsPdf(options);

        PowerPointPdfExportWarning warning = Assert.Single(options.Warnings);
        Assert.Equal("text-box-padding", warning.Code);
        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Equal(1, pdf.NumberOfPages);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersSolidSlideBackground() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        presentation.Slides[0].BackgroundColor = "112233";
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.SaveAsPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.067 0.133 0.2 rg", raw, StringComparison.Ordinal);
        Assert.Contains("0 0 240 160 re", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersGradientSlideBackground() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        presentation.Slides[0].SetBackgroundGradient("112233", "445566", 45D);
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.SaveAsPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("/SH1 sh", raw, StringComparison.Ordinal);
        Assert.Contains("/Shading", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersImageSlideBackground() {
        string imagePath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), Guid.NewGuid() + ".png");
        try {
            File.WriteAllBytes(imagePath, CreateMinimalRgbPng());
            using var stream = new MemoryStream();
            using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
            presentation.SlideSize.SetSizePoints(240, 160);
            presentation.Slides[0].SetBackgroundImage(imagePath);
            var options = new PowerPointPdfSaveOptions();

            byte[] bytes = presentation.SaveAsPdf(options);

            Assert.Empty(options.Warnings);
            string raw = Encoding.ASCII.GetString(bytes);
            Assert.Contains("/Im1 Do", raw, StringComparison.Ordinal);
            Assert.Contains("240 0 0 160 0 0 cm", raw, StringComparison.Ordinal);
        } finally {
            if (File.Exists(imagePath)) {
                File.Delete(imagePath);
            }
        }
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersTablesThroughSharedPdfCanvasTable() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(260, 180);
        PowerPointTable table = presentation.Slides[0].AddTablePoints(2, 2, 30, 34, 150, 70);
        table.FirstRow = true;
        table.BandedRows = false;
        table.SetColumnWidthsPoints(90, 60);
        table.SetRowHeightsPoints(28, 42);

        PowerPointTableCell header = table.GetCell(0, 0);
        header.Text = "Metric";
        header.FillColor = "D9EAF7";
        header.Bold = true;

        PowerPointTableCell headerScore = table.GetCell(0, 1);
        headerScore.Text = "Score";
        headerScore.FillColor = "D9EAF7";
        headerScore.HorizontalAlignment = TextAlignmentTypeValues.Center;

        PowerPointTableCell body = table.GetCell(1, 0);
        body.Text = "Quality";
        body.PaddingLeftPoints = 8D;
        body.BorderColor = "1E5A96";

        PowerPointTableCell score = table.GetCell(1, 1);
        score.Text = "99";
        score.FillColor = "EAF4FF";
        score.HorizontalAlignment = TextAlignmentTypeValues.Center;
        score.VerticalAlignment = TextAnchoringTypeValues.Center;

        var options = new PowerPointPdfSaveOptions();
        byte[] bytes = presentation.SaveAsPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("30 76 150 70 re", raw, StringComparison.Ordinal);
        Assert.Contains("120 146 m", raw, StringComparison.Ordinal);
        Assert.Contains("120 76 l", raw, StringComparison.Ordinal);
        Assert.Contains("30 118 m", raw, StringComparison.Ordinal);
        Assert.Contains("180 118 l", raw, StringComparison.Ordinal);
        Assert.Contains("120 76 60 42 re", raw, StringComparison.Ordinal);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Metric", text, StringComparison.Ordinal);
        Assert.Contains("Quality", text, StringComparison.Ordinal);
        Assert.Contains("99", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersChartsThroughSharedVectorRenderer() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 240);
        var data = new PowerPointChartData(
            new[] { "Q1", "Q2", "Q3", "Q4" },
            new[] {
                new PowerPointChartSeries("Sales", new[] { 12D, 18D, 24D, 30D }),
                new PowerPointChartSeries("Target", new[] { 15D, 20D, 22D, 28D })
            });
        PowerPointChart chart = presentation.Slides[0].AddChartPoints(data, 40, 32, 240, 172);
        chart.SetTitle("Revenue Mix");
        var options = new PowerPointPdfSaveOptions {
            ChartLayout = new OfficeChartLayout(preventLabelOverlap: false)
        };

        byte[] bytes = presentation.SaveAsPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("40 36 240 172 re", raw, StringComparison.Ordinal);
        Assert.Contains("0.122 0.306 0.475 rg", raw, StringComparison.Ordinal);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Revenue Mix", text, StringComparison.Ordinal);
        Assert.Contains("Sales", text, StringComparison.Ordinal);
        Assert.Contains("Target", text, StringComparison.Ordinal);
        Assert.Contains("Q1", text, StringComparison.Ordinal);
        Assert.Contains("Q4", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesHorizontalStackedBarChartKind() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 240);
        var data = new PowerPointChartData(
            new[] { "North", "South" },
            new[] {
                new PowerPointChartSeries("Won", new[] { 10D, 12D }),
                new PowerPointChartSeries("Open", new[] { 4D, 6D })
            });
        PowerPointChart chart = presentation.Slides[0].AddChartPoints(data, 40, 32, 240, 172);
        SetBarChartShape(chart, C.BarDirectionValues.Bar, C.BarGroupingValues.Stacked);

        Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot));
        Assert.Equal(PowerPointChartSnapshotKind.StackedBar, snapshot.ChartKind);

        byte[] bytes = presentation.SaveAsPdf();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Won", text, StringComparison.Ordinal);
        Assert.Contains("Open", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesStackedLineChartKind() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 240);
        var data = new PowerPointChartData(
            new[] { "Q1", "Q2", "Q3" },
            new[] {
                new PowerPointChartSeries("Actual", new[] { 10D, 12D, 16D }),
                new PowerPointChartSeries("Target", new[] { 3D, 4D, 5D })
            });
        PowerPointChart chart = presentation.Slides[0].AddLineChartPoints(data, 40, 32, 240, 172);
        SetLineChartGrouping(chart, C.GroupingValues.PercentStacked);

        Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot));
        Assert.Equal(PowerPointChartSnapshotKind.StackedLine100, snapshot.ChartKind);

        byte[] bytes = presentation.SaveAsPdf();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Actual", text, StringComparison.Ordinal);
        Assert.Contains("Target", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_RendersZeroThicknessLineAutoShapes() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(240, 160);
        PowerPointSlide slide = presentation.Slides[0];
        slide.AddShapePoints(ShapeTypeValues.Line, 20, 40, 100, 0).Stroke("1E5A96", 1.5D);
        slide.AddShapePoints(ShapeTypeValues.Line, 140, 30, 0, 80).Stroke("C00000", 1.5D);
        var options = new PowerPointPdfSaveOptions();

        byte[] bytes = presentation.SaveAsPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("20 120 m", raw, StringComparison.Ordinal);
        Assert.Contains("120 120 l", raw, StringComparison.Ordinal);
        Assert.Contains("140 130 m", raw, StringComparison.Ordinal);
        Assert.Contains("140 50 l", raw, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_PreservesScatterSeriesXValues() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 240);
        var data = new PowerPointScatterChartData(new[] {
            new PowerPointScatterChartSeries("Actual", new[] { 1D, 2D, 3D }, new[] { 10D, 12D, 14D }),
            new PowerPointScatterChartSeries("Forecast", new[] { 1.5D, 2.5D }, new[] { 11D, 13D })
        });
        PowerPointChart chart = presentation.Slides[0].AddScatterChartPoints(data, 40, 32, 240, 172);

        Assert.True(chart.TryGetSnapshot(out PowerPointChartSnapshot snapshot));
        PowerPointChartSeries forecast = Assert.Single(snapshot.Data.Series, series => series.Name == "Forecast");
        Assert.Equal(new[] { 1.5D, 2.5D }, forecast.XValues);

        byte[] bytes = presentation.SaveAsPdf();

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Actual", text, StringComparison.Ordinal);
        Assert.Contains("Forecast", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_AppliesSharedChartStyleAndLayoutOptions() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(360, 240);
        var data = new PowerPointChartData(
            new[] { "M01", "M02", "M03", "M04", "M05", "M06", "M07", "M08" },
            new[] {
                new PowerPointChartSeries("Actual", new[] { 12D, 18D, 24D, 30D, 34D, 38D, 41D, 43D }),
                new PowerPointChartSeries("Target", new[] { 15D, 20D, 22D, 28D, 31D, 36D, 39D, 44D })
            });
        PowerPointChart chart = presentation.Slides[0].AddChartPoints(data, 38, 30, 270, 176);
        chart.SetTitle("Styled Slide Chart");
        var options = new PowerPointPdfSaveOptions {
            ChartStyle = new OfficeChartStyle(
                palette: new[] {
                    OfficeColor.FromRgb(18, 52, 86),
                    OfficeColor.FromRgb(120, 40, 160)
                },
                backgroundColor: OfficeColor.FromRgb(242, 248, 255),
                titleColor: OfficeColor.FromRgb(200, 10, 10)),
            ChartLayout = new OfficeChartLayout(maximumCategoryAxisLabels: 2)
        };

        byte[] bytes = presentation.SaveAsPdf(options);

        Assert.Empty(options.Warnings);
        string raw = Encoding.ASCII.GetString(bytes);
        Assert.Contains("0.071 0.204 0.337 rg", raw, StringComparison.Ordinal);
        Assert.Contains("0.471 0.157 0.627 rg", raw, StringComparison.Ordinal);
        Assert.Contains("0.949 0.973 1 rg", raw, StringComparison.Ordinal);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        string text = string.Join("", pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("Styled Slide Chart", text, StringComparison.Ordinal);
        Assert.Contains("Actual", text, StringComparison.Ordinal);
        Assert.Contains("Target", text, StringComparison.Ordinal);
        Assert.Contains("M01", text, StringComparison.Ordinal);
        Assert.Contains("M05", text, StringComparison.Ordinal);
        Assert.DoesNotContain("M02", text, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_PowerPointPresentation_WarnsWhenSharedChartQualityPreflightFindsIssues() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 220);
        var data = new PowerPointChartData(
            new[] { "M01", "M02", "M03", "M04", "M05", "M06", "M07", "M08", "M09", "M10", "M11", "M12" },
            new[] {
                new PowerPointChartSeries("Actual", new[] { 12D, 18D, 24D, 30D, 34D, 38D, 41D, 43D, 44D, 45D, 46D, 47D })
            });
        PowerPointChart chart = presentation.Slides[0].AddChartPoints(data, 32, 28, 240, 150);
        chart.SetTitle("Dense Slide Chart");
        var options = new PowerPointPdfSaveOptions {
            ChartLayout = new OfficeChartLayout(maximumCategoryAxisLabels: 12, preventLabelOverlap: false)
        };

        byte[] bytes = presentation.SaveAsPdf(options);

        PowerPointPdfExportWarning warning = Assert.Single(options.Warnings, item => item.Code == "chart-quality");
        Assert.Equal(1, warning.SlideNumber);
        Assert.Contains("Dense Slide Chart", warning.Message, StringComparison.Ordinal);
        Assert.Contains("TextOverlap", warning.Message, StringComparison.Ordinal);

        using var pdf = PdfPigDocument.Open(new MemoryStream(bytes));
        Assert.Contains("Dense Slide Chart", pdf.GetPage(1).Text, StringComparison.Ordinal);
    }

    private static byte[] CreateMinimalRgbPng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10,
            0, 0, 0, 13,
            73, 72, 68, 82,
            0, 0, 0, 1,
            0, 0, 0, 1,
            8, 2, 0, 0, 0,
            0, 0, 0, 0,
            0, 0, 0, 12,
            73, 68, 65, 84,
            0x78, 0x9C, 0x63, 0xF8, 0xCF, 0xC0, 0x00, 0x00, 0x03, 0x01, 0x01, 0x00,
            0, 0, 0, 0,
            0, 0, 0, 0,
            73, 69, 78, 68,
            0, 0, 0, 0
        };
    }

    private static void SetBarChartShape(PowerPointChart chart, C.BarDirectionValues direction, C.BarGroupingValues grouping) {
        MethodInfo method = typeof(PowerPointChart).GetMethod("GetChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!;
        var chartPart = (ChartPart)method.Invoke(chart, Array.Empty<object>())!;
        C.BarChart barChart = chartPart.ChartSpace!.Descendants<C.BarChart>().Single();
        barChart.GetFirstChild<C.BarDirection>()!.Val = direction;
        barChart.GetFirstChild<C.BarGrouping>()!.Val = grouping;
        chartPart.ChartSpace.Save();
    }

    private static void SetLineChartGrouping(PowerPointChart chart, C.GroupingValues grouping) {
        MethodInfo method = typeof(PowerPointChart).GetMethod("GetChartPart", BindingFlags.NonPublic | BindingFlags.Instance)!;
        var chartPart = (ChartPart)method.Invoke(chart, Array.Empty<object>())!;
        C.LineChart lineChart = chartPart.ChartSpace!.Descendants<C.LineChart>().Single();
        C.Grouping chartGrouping = lineChart.GetFirstChild<C.Grouping>() ?? lineChart.PrependChild(new C.Grouping());
        chartGrouping.Val = grouping;
        chartPart.ChartSpace.Save();
    }

    private static double FindWordStartY(UglyToad.PdfPig.Content.Page page, string word) {
        var lines = page.Letters
            .Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1));

        foreach (var line in lines) {
            var ordered = line.OrderBy(letter => letter.StartBaseLine.X).ToList();
            string text = string.Concat(ordered.Select(letter => letter.Value));
            int index = text.IndexOf(word, StringComparison.Ordinal);
            if (index >= 0) {
                return ordered[index].StartBaseLine.Y;
            }
        }

        throw new InvalidOperationException("Could not find word '" + word + "' in rendered PDF text.");
    }
}
