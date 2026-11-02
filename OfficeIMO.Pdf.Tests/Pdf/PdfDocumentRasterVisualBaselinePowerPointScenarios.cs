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
        using PowerPointDocument presentation = PowerPointDocument.Create(stream);
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(320),OfficeIMO.OfficeLength.Points(180));
        PowerPointSlide slide = presentation.AddSlide();

        PowerPointAutoShape background = slide.AddRectangle(OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(320), OfficeIMO.OfficeLength.Points(180));
        background.FillColor = "F6FAFC";
        background.OutlineColor = "F6FAFC";

        PowerPointAutoShape panel = slide.AddShape(OfficePresetShapeType.RoundRectangle, OfficeIMO.OfficeLength.Points(14), OfficeIMO.OfficeLength.Points(14), OfficeIMO.OfficeLength.Points(132), OfficeIMO.OfficeLength.Points(72));
        panel.FillColor = "FFFFFF";
        panel.OutlineColor = "B9D7EA";
        panel.OutlineWidth = global::OfficeIMO.OfficeLength.Points(1.2D);

        PowerPointTextBox title = slide.AddTextBox("PowerPoint PDF Visual Gate",global::OfficeIMO.OfficeLength.Points(22),global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(122),global::OfficeIMO.OfficeLength.Points(32));
        title.FontName = "Georgia";
        title.FontSize = global::OfficeIMO.OfficeLength.Points(12);
        title.Color = "123456";
        title.FillTransparency = 100;

        PowerPointTextBox body = slide.AddTextBox(string.Empty,global::OfficeIMO.OfficeLength.Points(22),global::OfficeIMO.OfficeLength.Points(52),global::OfficeIMO.OfficeLength.Points(118),global::OfficeIMO.OfficeLength.Points(30));
        body.FontSize = global::OfficeIMO.OfficeLength.Points(8);
        body.Color = "334155";
        body.FillTransparency = 100;
        body.SetBullets(new[] { "Text placement", "Tables and charts" });

        slide.AddImage(new MemoryStream(CreatePowerPointVisualGatePng()), OfficeImageFormat.Png, global::OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromPoints(250)), global::OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromPoints(22)), global::OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromPoints(42)), global::OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromPoints(30)));

        PowerPointTable table = slide.AddTable(2,2,global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(98),global::OfficeIMO.OfficeLength.Points(120),global::OfficeIMO.OfficeLength.Points(48));
        table.ApplyToCells(cell => cell.FontSize = global::OfficeIMO.OfficeLength.Points(8));
        table.SetColumnWidths(global::OfficeIMO.OfficeLength.Points(70),global::OfficeIMO.OfficeLength.Points(50));
        table.SetRowHeights(global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(28));
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
        value.HorizontalAlignment = OfficeIMO.PowerPoint.PowerPointTextAlignment.Center;

        var chartData = new PowerPointChartData(
            new[] { "Q1", "Q2", "Q3" },
            new[] {
                new PowerPointChartSeries("Actual", new[] { 12D, 19D, 27D }),
                new PowerPointChartSeries("Target", new[] { 14D, 21D, 25D })
            });
        PowerPointChart chart = slide.AddChartPoints(chartData, 166, 78, 132, 86);
        chart.SetTitle("Revenue");

        PowerPointAutoShape rule = slide.AddShape(OfficePresetShapeType.Line, OfficeIMO.OfficeLength.Points(154), OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(80), OfficeIMO.OfficeLength.Points(0));
        rule.OutlineColor = "1E5A96";
        rule.OutlineWidth = global::OfficeIMO.OfficeLength.Points(1.5D);

        var options = new PowerPointPdfExportOptions {
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
        return (presentation.ToPdf(options)).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
    }

    private static byte[] CreateNativePowerPointDenseLayout() {
        using var stream = new MemoryStream();
        using PowerPointDocument presentation = PowerPointDocument.Create(stream);
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(320),OfficeIMO.OfficeLength.Points(180));
        PowerPointSlide slide = presentation.AddSlide();

        PowerPointAutoShape background = slide.AddRectangle(OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(320), OfficeIMO.OfficeLength.Points(180));
        background.FillColor = "FFFDF7";
        background.OutlineColor = "FFFDF7";

        PowerPointTextBox title = slide.AddTextBox("Dense Layout Gate",global::OfficeIMO.OfficeLength.Points(18),global::OfficeIMO.OfficeLength.Points(16),global::OfficeIMO.OfficeLength.Points(170),global::OfficeIMO.OfficeLength.Points(22));
        title.FontName = "Arial";
        title.FontSize = global::OfficeIMO.OfficeLength.Points(14);
        title.Color = "1A1F2B";
        title.Bold = true;
        title.FillTransparency = 100;

        PowerPointTextBox list = slide.AddTextBox(string.Empty,global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(44),global::OfficeIMO.OfficeLength.Points(128),global::OfficeIMO.OfficeLength.Points(62));
        list.FontSize = global::OfficeIMO.OfficeLength.Points(8);
        list.Color = "334155";
        list.FillTransparency = 100;
        list.SetBullets(
            new[] { "Explicit hanging indent", "Tight but readable prefix" },
            configure: paragraph => {
                paragraph.SetLeftMargin(global::OfficeIMO.OfficeLength.Points(22));
                paragraph.SetHanging(global::OfficeIMO.OfficeLength.Points(10));
            });

        PowerPointTable table = slide.AddTable(2,2,global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(118),global::OfficeIMO.OfficeLength.Points(132),global::OfficeIMO.OfficeLength.Points(40));
        table.SetColumnWidths(global::OfficeIMO.OfficeLength.Points(76),global::OfficeIMO.OfficeLength.Points(56));
        table.SetRowHeights(global::OfficeIMO.OfficeLength.Points(18),global::OfficeIMO.OfficeLength.Points(22));
        table.GetCell(0, 0).Text = "Section";
        table.GetCell(0, 1).Text = "Status";
        table.GetCell(0, 0).Bold = true;
        table.GetCell(0, 1).Bold = true;
        table.GetCell(1, 0).Text = "Packed";
        table.GetCell(1, 1).Text = "Watch";
        table.GetCell(1, 1).HorizontalAlignment = OfficeIMO.PowerPoint.PowerPointTextAlignment.Center;

        byte[] logoBytes = File.ReadAllBytes(System.IO.Path.Combine(GetTestsProjectRoot(), "Images", "EvotecLogo.png"));
        PowerPointPicture picture = slide.AddImage(new MemoryStream(logoBytes), OfficeImageFormat.Png, global::OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromPoints(220)), global::OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromPoints(18)), global::OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromPoints(64)), global::OfficeIMO.OfficeLength.Emus(PowerPointUnits.FromPoints(30)));
        picture.AltText = "OfficeIMO visual gate logo";

        var chartData = new PowerPointChartData(
            new[] { "A", "B", "C", "D" },
            new[] {
                new PowerPointChartSeries("Current", new[] { 18D, 24D, 16D, 30D })
            });
        PowerPointChart chart = slide.AddChartPoints(chartData, 166, 82, 126, 74);
        chart.SetTitle("Flow");

        var options = new PowerPointPdfExportOptions {
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
        return (presentation.ToPdf(options)).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
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
