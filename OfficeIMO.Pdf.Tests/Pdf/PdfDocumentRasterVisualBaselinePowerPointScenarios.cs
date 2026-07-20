using System.IO;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using DocumentFormat.OpenXml.Drawing;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfDocumentRasterVisualBaselineTests {
    [Fact]
    public void NativePowerPointSlide_MatchesPopplerRasterBaseline() {
        AssertScenarioRasterBaseline("native-powerpoint-slide", CreateNativePowerPointSlide);
    }

    [Fact]
    public void NativePowerPointDenseLayout_MatchesPopplerRasterBaseline() {
        AssertScenarioRasterBaseline("native-powerpoint-dense-layout", CreateNativePowerPointDenseLayout);
    }

    private static byte[] CreateNativePowerPointSlide() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 180);
        PowerPointSlide slide = presentation.AddSlide();

        PowerPointAutoShape background = slide.AddRectanglePoints(0, 0, 320, 180);
        background.FillColor = "F6FAFC";
        background.OutlineColor = "F6FAFC";

        PowerPointAutoShape panel = slide.AddShapePoints(ShapeTypeValues.RoundRectangle, 14, 14, 132, 72);
        panel.FillColor = "FFFFFF";
        panel.OutlineColor = "B9D7EA";
        panel.OutlineWidthPoints = 1.2D;

        PowerPointTextBox title = slide.AddTextBoxPoints("PowerPoint PDF Visual Gate", 22, 20, 122, 32);
        title.FontName = "Georgia";
        title.FontSize = 12;
        title.Color = "123456";
        title.FillTransparency = 100;

        PowerPointTextBox body = slide.AddTextBoxPoints(string.Empty, 22, 52, 118, 30);
        body.FontSize = 8;
        body.Color = "334155";
        body.FillTransparency = 100;
        body.SetBullets(new[] { "Text placement", "Tables and charts" });

        slide.AddPicture(new MemoryStream(CreatePowerPointVisualGatePng()), ImagePartType.Png, PowerPointUnits.FromPoints(250), PowerPointUnits.FromPoints(22), PowerPointUnits.FromPoints(42), PowerPointUnits.FromPoints(30));

        PowerPointTable table = slide.AddTablePoints(2, 2, 20, 98, 120, 48);
        table.ApplyToCells(cell => cell.FontSize = 8);
        table.SetColumnWidthsPoints(70, 50);
        table.SetRowHeightsPoints(20, 28);
        PowerPointTableCell metric = table.GetCell(0, 0);
        metric.Text = "Metric";
        metric.Bold = true;
        metric.FillColor = "D9EAF7";
        PowerPointTableCell score = table.GetCell(0, 1);
        score.Text = "Score";
        score.Bold = true;
        score.FillColor = "D9EAF7";
        table.GetCell(1, 0).Text = "Quality";
        PowerPointTableCell value = table.GetCell(1, 1);
        value.Text = "99";
        value.HorizontalAlignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center;

        var chartData = new PowerPointChartData(
            new[] { "Q1", "Q2", "Q3" },
            new[] {
                new PowerPointChartSeries("Actual", new[] { 12D, 19D, 27D }),
                new PowerPointChartSeries("Target", new[] { 14D, 21D, 25D })
            });
        PowerPointChart chart = slide.AddChartPoints(chartData, 166, 78, 132, 86);
        chart.SetTitle("Revenue");

        PowerPointAutoShape rule = slide.AddShapePoints(ShapeTypeValues.Line, 154, 20, 80, 0);
        rule.OutlineColor = "1E5A96";
        rule.OutlineWidthPoints = 1.5D;

        var options = new PowerPointPdfSaveOptions {
            ChartStyle = new OfficeChartStyle(
                palette: new[] {
                    OfficeColor.FromRgb(30, 90, 150),
                    OfficeColor.FromRgb(120, 40, 160)
                },
                backgroundColor: OfficeColor.White,
                titleColor: OfficeColor.FromRgb(18, 52, 86)),
            ChartLayout = new OfficeChartLayout(maximumCategoryAxisLabels: 3)
        };

        presentation.Save();
        WriteReviewArtifact("native-powerpoint-slide.pptx", stream.ToArray());
        return presentation.ToPdf(options);
    }

    private static byte[] CreateNativePowerPointDenseLayout() {
        using var stream = new MemoryStream();
        using PowerPointPresentation presentation = PowerPointPresentation.Create(stream);
        presentation.SlideSize.SetSizePoints(320, 180);
        PowerPointSlide slide = presentation.AddSlide();

        PowerPointAutoShape background = slide.AddRectanglePoints(0, 0, 320, 180);
        background.FillColor = "FFFDF7";
        background.OutlineColor = "FFFDF7";

        PowerPointTextBox title = slide.AddTextBoxPoints("Dense Layout Gate", 18, 16, 170, 22);
        title.FontName = "Arial";
        title.FontSize = 14;
        title.Color = "1A1F2B";
        title.Bold = true;
        title.FillTransparency = 100;

        PowerPointTextBox list = slide.AddTextBoxPoints(string.Empty, 20, 44, 128, 62);
        list.FontSize = 8;
        list.Color = "334155";
        list.FillTransparency = 100;
        list.SetBullets(
            new[] { "Explicit hanging indent", "Tight but readable prefix" },
            configure: paragraph => {
                paragraph.SetLeftMarginPoints(22);
                paragraph.SetHangingPoints(10);
            });

        PowerPointTable table = slide.AddTablePoints(2, 2, 20, 118, 132, 40);
        table.SetColumnWidthsPoints(76, 56);
        table.SetRowHeightsPoints(18, 22);
        table.GetCell(0, 0).Text = "Section";
        table.GetCell(0, 1).Text = "Status";
        table.GetCell(0, 0).Bold = true;
        table.GetCell(0, 1).Bold = true;
        table.GetCell(1, 0).Text = "Packed";
        table.GetCell(1, 1).Text = "Watch";
        table.GetCell(1, 1).HorizontalAlignment = DocumentFormat.OpenXml.Drawing.TextAlignmentTypeValues.Center;

        byte[] logoBytes = File.ReadAllBytes(System.IO.Path.Combine(GetTestsProjectRoot(), "Images", "EvotecLogo.png"));
        PowerPointPicture picture = slide.AddPicture(new MemoryStream(logoBytes), ImagePartType.Png, PowerPointUnits.FromPoints(220), PowerPointUnits.FromPoints(18), PowerPointUnits.FromPoints(64), PowerPointUnits.FromPoints(30));
        picture.AltText = "OfficeIMO visual gate logo";

        var chartData = new PowerPointChartData(
            new[] { "A", "B", "C", "D" },
            new[] {
                new PowerPointChartSeries("Current", new[] { 18D, 24D, 16D, 30D })
            });
        PowerPointChart chart = slide.AddChartPoints(chartData, 166, 82, 126, 74);
        chart.SetTitle("Flow");

        var options = new PowerPointPdfSaveOptions {
            ChartStyle = new OfficeChartStyle(
                palette: new[] {
                    OfficeColor.FromRgb(34, 126, 102)
                },
                backgroundColor: OfficeColor.White,
                titleColor: OfficeColor.FromRgb(26, 31, 43)),
            ChartLayout = new OfficeChartLayout(maximumCategoryAxisLabels: 4)
        };

        presentation.Save();
        WriteReviewArtifact("native-powerpoint-dense-layout.pptx", stream.ToArray());
        return presentation.ToPdf(options);
    }

    private static byte[] CreatePowerPointVisualGatePng() {
        return new byte[] {
            137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82,
            0, 0, 0, 2, 0, 0, 0, 2, 8, 2, 0, 0, 0, 253, 212, 154, 115,
            0, 0, 0, 22, 73, 68, 65, 84, 120, 156, 99, 100, 248, 207, 192,
            192, 240, 159, 129, 129, 225, 63, 3, 3, 0, 26, 8, 3, 253, 160,
            177, 180, 62, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130
        };
    }
}
