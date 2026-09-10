using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointParagraphFlowContractTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CondensedDescendersAreClippedOnlyAtTheTextBoxBoundary(bool rich) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        presentation.SlideSize.SetSizePoints(300, 220);
        var box = presentation.AddSlide().AddTextBoxPoints("gypsy\nplaying\nquickly", 20, 20, 240, 150);
        box.FontSize = 24; box.FontName = "Helvetica";
        box.TextMarginLeftPoints = box.TextMarginRightPoints = box.TextMarginTopPoints = box.TextMarginBottomPoints = 0;
        box.Paragraphs[0].LineSpacingMultiplier = 0.75D;
        box.Paragraphs[0].Runs[0].Bold = rich;
        byte[] bytes = presentation.ToPdfBytes();
        Assert.Contains("20 50 240 150 re W", PdfOperatorSearchText.From(bytes), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public void CondensedParagraphLeadingSurvivesNativeRendering(bool table, bool rich, bool trailingParagraph) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        presentation.SlideSize.SetSizePoints(300, 220);
        var slide = presentation.AddSlide();
        PowerPointParagraph paragraph;
        if (table) {
            var cell = slide.AddTablePoints(1, 1, 20, 20, 240, 160).GetCell(0, 0);
            cell.Text = "FIRST\nSECOND\nTHIRD"; cell.FontName = "Helvetica"; cell.FontSize = 12;
            paragraph = cell.Paragraphs[0];
            if (trailingParagraph) cell.AddParagraph("END");
        } else {
            var box = slide.AddTextBoxPoints("FIRST\nSECOND\nTHIRD", 20, 20, 240, 160);
            box.FontName = "Helvetica"; box.FontSize = 12;
            paragraph = box.Paragraphs[0];
            if (trailingParagraph) box.AddParagraph("END").Alignment = PowerPointTextAlignment.Right;
        }
        paragraph.LineSpacingMultiplier = 0.75D;
        if (rich) paragraph.Runs[0].Bold = true;
        if (!table) {
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(presentation.ToPdfBytes());
            Assert.InRange(BaselineFromTop(pdf, 1, "SECOND") - BaselineFromTop(pdf, 1, "FIRST"), 8.9D, 9.1D);
            Assert.InRange(BaselineFromTop(pdf, 1, "THIRD") - BaselineFromTop(pdf, 1, "SECOND"), 8.9D, 9.1D);
        }
        var svg = XDocument.Parse(Encoding.UTF8.GetString(slide.ExportImage(OfficeImageExportFormat.Svg).Bytes));
        double Y(string value) => double.Parse(Assert.Single(svg.Descendants(), element => element.Name.LocalName == "text" && element.Value == value).Attribute("y")!.Value, CultureInfo.InvariantCulture);
        Assert.InRange(Y("SECOND") - Y("FIRST"), 8.9D, 9.1D);
        Assert.InRange(Y("THIRD") - Y("SECOND"), 8.9D, 9.1D);
    }

    [Theory]
    [InlineData("paragraph", false)]
    [InlineData("paragraph", true)]
    [InlineData("list", false)]
    [InlineData("list", true)]
    [InlineData("master", false)]
    [InlineData("master", true)]
    [InlineData("end", false)]
    [InlineData("end", true)]
    public void BlankParagraphUsesEffectiveFontSize(string source, bool differentAlignment) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        presentation.SlideSize.SetSizePoints(300, 220);
        foreach (bool spacer in new[] { false, true }) {
            var box = presentation.AddSlide().AddTextBoxPoints("FIRST", 20, 20, 240, 160);
            box.FontName = "Helvetica"; box.FontSize = 12;
            if (spacer) {
                var empty = box.AddParagraph("");
                empty.Paragraph.Elements<A.Run>().Single().RunProperties = new A.RunProperties();
                empty.Paragraph.RemoveAllChildren<A.EndParagraphRunProperties>();
                if (source == "end") empty.Paragraph.Append(new A.EndParagraphRunProperties { FontSize = 3000 });
                else SetDefaults(box, empty.Paragraph, source, new A.DefaultRunProperties { FontSize = 3000 });
            }
            var second = box.AddParagraph("SECOND");
            if (differentAlignment) second.Alignment = PowerPointTextAlignment.Right;
        }
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(presentation.ToPdfBytes());
        Assert.InRange(BaselineFromTop(pdf, 2, "SECOND") - BaselineFromTop(pdf, 1, "SECOND"), 35D, 37D);
    }

    [Fact]
    public void BlankTableCellParagraphRetainsItsLineAndSpacing() {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        presentation.SlideSize.SetSizePoints(300, 220);
        foreach (bool spacer in new[] { false, true }) {
            var cell = presentation.AddSlide().AddTablePoints(1, 1, 20, 20, 240, 160).GetCell(0, 0);
            cell.Text = "FIRST"; cell.FontName = "Helvetica"; cell.FontSize = 12;
            if (spacer) cell.AddParagraph("").SetSpaceBeforePoints(7).SetSpaceAfterPoints(11);
            cell.AddParagraph("SECOND");
        }
        double SvgY(int index) {
            var svg = XDocument.Parse(Encoding.UTF8.GetString(presentation.Slides[index].ExportImage(OfficeImageExportFormat.Svg).Bytes));
            return double.Parse(Assert.Single(svg.Descendants(), node => node.Name.LocalName == "text" && node.Value == "SECOND").Attribute("y")!.Value, CultureInfo.InvariantCulture);
        }
        Assert.InRange(SvgY(1) - SvgY(0), 32D, 34D);
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("list")]
    [InlineData("master")]
    [InlineData("end")]
    public void BlankTableCellParagraphUsesEffectiveFontSize(string source) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        presentation.SlideSize.SetSizePoints(300, 220);
        foreach (bool spacer in new[] { false, true }) {
            var cell = presentation.AddSlide().AddTablePoints(1, 1, 20, 20, 240, 160).GetCell(0, 0);
            cell.Text = "FIRST"; cell.FontName = "Helvetica"; cell.FontSize = 12;
            if (spacer) {
                var empty = cell.AddParagraph("");
                empty.Paragraph.RemoveAllChildren<A.Run>();
                empty.Paragraph.RemoveAllChildren<A.EndParagraphRunProperties>();
                var defaults = new A.DefaultRunProperties { FontSize = 3000 };
                if (source == "end") empty.Paragraph.Append(new A.EndParagraphRunProperties { FontSize = 3000 });
                else if (source == "paragraph") empty.Paragraph.ParagraphProperties = new A.ParagraphProperties(defaults);
                else if (source == "list") cell.Cell.TextBody!.ListStyle!.Append(new A.Level1ParagraphProperties(defaults));
                else {
                    var master = cell.SlidePart!.SlideLayoutPart!.SlideMasterPart!.SlideMaster!.TextStyles!.OtherStyle!;
                    var level = master.GetFirstChild<A.Level1ParagraphProperties>() ?? master.AppendChild(new A.Level1ParagraphProperties());
                    level.RemoveAllChildren<A.DefaultRunProperties>(); level.Append(defaults);
                }
            }
            cell.AddParagraph("SECOND").Runs[0].FontSize = 12;
        }
        double SvgY(int index) {
            var svg = XDocument.Parse(Encoding.UTF8.GetString(presentation.Slides[index].ExportImage(OfficeImageExportFormat.Svg).Bytes));
            return double.Parse(Assert.Single(svg.Descendants(), node => node.Name.LocalName == "text" && node.Value == "SECOND").Attribute("y")!.Value, CultureInfo.InvariantCulture);
        }
        Assert.InRange(SvgY(1) - SvgY(0), 35D, 37D);
    }

    [Theory]
    [InlineData("list", false)]
    [InlineData("list", true)]
    [InlineData("master", false)]
    [InlineData("master", true)]
    public void SingleEffectiveRunRetainsInheritedCapitalization(string source, bool flow) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        var box = presentation.AddSlide().AddTextBoxPoints("mixed case", 20, 20, 240, 120);
        box.FontName = "Helvetica"; box.FontSize = 12;
        if (flow) box.AddParagraph("SECOND").Alignment = PowerPointTextAlignment.Right;
        SetDefaults(box, box.Paragraphs[0].Paragraph, source, new A.DefaultRunProperties { Capital = A.TextCapsValues.All });
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(presentation.ToPdfBytes());
        Assert.Contains("MIXED CASE", pdf.GetPage(1).Text, StringComparison.Ordinal);
        var svg = Encoding.UTF8.GetString(presentation.Slides[0].ExportImage(OfficeImageExportFormat.Svg).Bytes);
        Assert.Contains(">MIXED CASE</text>", svg, StringComparison.Ordinal);
    }

    private static double BaselineFromTop(UglyToad.PdfPig.PdfDocument pdf, int pageNumber, string value) {
        var page = pdf.GetPage(pageNumber);
        string text = string.Concat(page.Letters.Select(letter => letter.Value));
        return page.Height - page.Letters[text.IndexOf(value, StringComparison.Ordinal)].StartBaseLine.Y;
    }

    private static void SetDefaults(PowerPointTextBox box, A.Paragraph paragraph, string source, A.DefaultRunProperties defaults) {
        if (source == "paragraph") paragraph.ParagraphProperties = new A.ParagraphProperties(defaults);
        else if (source == "list") box.TextBody!.ListStyle!.Append(new A.Level1ParagraphProperties(defaults));
        else {
            var style = Assert.IsType<P.OtherStyle>(box.MasterTextStyle);
            var level = style.GetFirstChild<A.Level1ParagraphProperties>() ?? style.AppendChild(new A.Level1ParagraphProperties());
            level.RemoveAllChildren<A.DefaultRunProperties>(); level.Append(defaults);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void EmptyParagraphRetainsItsLineAndSpacingInPdfAndSvg(bool bullet) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        presentation.SlideSize.SetSizePoints(300, 220);
        foreach (bool spacer in new[] { false, true }) {
            var box = presentation.AddSlide().AddTextBoxPoints("FIRST", 20, 20, 240, 160);
            box.FontName = "Helvetica";
            box.FontSize = 12;
            if (spacer) {
                var empty = box.AddParagraph("");
                empty.SetSpaceBeforePoints(7);
                empty.SetSpaceAfterPoints(11);
            }
            box.AddParagraph("SECOND").Alignment = PowerPointTextAlignment.Right;
            if (bullet) box.Paragraphs[0].SetBullet('*');
        }
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(presentation.ToPdfBytes(new PowerPointToPdfOptions {
            ResourcePolicy = OfficeIMO.Pdf.PdfResourcePolicy.CreatePortableDeterministic()
        }));
        double PdfY(int pageNumber) {
            var page = pdf.GetPage(pageNumber);
            string text = string.Concat(page.Letters.Select(letter => letter.Value));
            return page.Height - page.Letters[text.IndexOf("SECOND", StringComparison.Ordinal)].StartBaseLine.Y;
        }
        Assert.InRange(PdfY(2) - PdfY(1), 32D, 34D);
        double SvgY(int index) {
            var svg = XDocument.Parse(Encoding.UTF8.GetString(presentation.Slides[index].ExportImage(OfficeImageExportFormat.Svg).Bytes));
            var second = Assert.Single(svg.Descendants(), node => node.Name.LocalName == "text" && node.Value == "SECOND");
            return double.Parse(second.Attribute("y")!.Value, CultureInfo.InvariantCulture);
        }
        Assert.InRange(SvgY(1) - SvgY(0), 32D, 34D);
    }

    [Theory]
    [InlineData("paragraph")]
    [InlineData("list")]
    [InlineData("master")]
    public void PlainParagraphUsesItsEffectiveInheritedColor(string source) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        var slide = presentation.AddSlide();
        var box = slide.AddTextBoxPoints("INHERITED", 20, 20, 240, 100);
        box.FontName = "Helvetica";
        box.FontSize = 12;
        box.AddParagraph("SECOND").Alignment = PowerPointTextAlignment.Right;
        var paragraph = box.TextBody!.Elements<A.Paragraph>().First();
        paragraph.Elements<A.Run>().Single().RunProperties = new A.RunProperties();
        var defaults = new A.DefaultRunProperties(new A.SolidFill(new A.RgbColorModelHex { Val = "C02060" }));
        if (source == "paragraph") paragraph.ParagraphProperties = new A.ParagraphProperties(defaults);
        else if (source == "list") box.TextBody.ListStyle!.Append(new A.Level1ParagraphProperties(defaults));
        else {
            var style = Assert.IsType<P.OtherStyle>(box.MasterTextStyle);
            var level = style.GetFirstChild<A.Level1ParagraphProperties>() ?? style.AppendChild(new A.Level1ParagraphProperties());
            level.RemoveAllChildren<A.DefaultRunProperties>();
            level.Append(defaults);
        }
        var svg = XDocument.Parse(Encoding.UTF8.GetString(slide.ExportImage(OfficeImageExportFormat.Svg).Bytes));
        var text = Assert.Single(svg.Descendants(), node => node.Name.LocalName == "text" && node.Value == "INHERITED");
        Assert.Equal("#c02060", text.Attribute("fill")!.Value.ToLowerInvariant());
        string operators = PdfOperatorSearchText.From(presentation.ToPdfBytes());
        Assert.Contains("0.753 0.125 0.376 rg", operators, StringComparison.Ordinal);
    }
}
