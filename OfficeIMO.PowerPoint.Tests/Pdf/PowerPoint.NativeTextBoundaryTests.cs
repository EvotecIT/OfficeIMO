using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointNativeTextBoundaryTests {
    [Fact]
    public void PartiallyVisibleTableCellParagraphsRetainTheirFittingPrefix() {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTable table = slide.AddTablePoints(1, 1, 20, 20, 140, 40);
        table.SetRowHeightsPoints(40);
        PowerPointTableCell cell = table.GetCell(0, 0);
        cell.FontSize = 12;
        cell.Text = "VISIBLE alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu TAIL";
        cell.Paragraphs[0].SetBullet('*');
        var image = slide.ExportImage(OfficeImageExportFormat.Svg);
        XDocument svg = XDocument.Parse(Encoding.UTF8.GetString(image.Bytes));
        string text = string.Concat(svg.Descendants().Where(element => element.Name.LocalName == "text").Select(element => element.Value));
        Assert.Contains("VISIBLE", text);
        Assert.DoesNotContain("TAIL", text);
        Assert.Contains(image.Diagnostics, warning => warning.Message.Contains("cell text frame boundary"));
    }

    [Theory]
    [InlineData("bullet")]
    [InlineData("alignment")]
    [InlineData("spacing")]
    public void PartiallyVisibleParagraphsRetainTheirFittingPrefix(string flow) {
        var fonts = new OfficeFontFaceCollection().Add("BoundaryFont",
            OfficeIMO.TestAssets.PdfTestFontAssets.LoadBundledOpenTypeCffFont());
        var profile = new OfficeRenderingProfile("native-overflow", fonts);
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        presentation.SlideSize.SetSizePoints(240, 140);
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTextBox box = slide.AddTextBoxPoints(
            "VISIBLE alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron TAIL", 20, 20, 120, 40);
        box.FontName = "BoundaryFont";
        box.FontSize = 12;
        if (flow == "bullet") box.Paragraphs[0].SetBullet('*');
        else {
            PowerPointParagraph next = box.AddParagraph("LATER");
            if (flow == "alignment") next.Alignment = PowerPointTextAlignment.Right;
            else box.Paragraphs[0].SetSpaceAfterPoints(4);
        }
        var result = presentation.ToPdfDocumentResult(new PowerPointToPdfOptions().UseRenderingProfile(profile));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(result.ToBytes());
        string pdfText = string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value));
        Assert.Contains("VISIBLE", pdfText);
        Assert.DoesNotContain("TAIL", pdfText);
        Assert.DoesNotContain("LATER", pdfText);
        Assert.Single(result.Warnings, warning => warning.Code == "text-box-overflow");

        var image = slide.ExportImage(OfficeImageExportFormat.Svg, new PowerPointImageExportOptions().UseRenderingProfile(profile));
        XDocument svg = XDocument.Parse(Encoding.UTF8.GetString(image.Bytes));
        string svgText = string.Concat(svg.Descendants().Where(element => element.Name.LocalName == "text").Select(element => element.Value));
        Assert.Contains("VISIBLE", svgText);
        Assert.DoesNotContain("TAIL", svgText);
        Assert.Single(image.Diagnostics, warning => warning.Code == "POWERPOINT_TEXT_OVERFLOW");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PlainAndRichTextClippingReportsBoundedLoss(bool rich) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointTextBox box = presentation.AddSlide().AddTextBoxPoints("VISIBLE alpha beta gamma delta epsilon TAIL", 20, 20, 90, 23);
        box.FontSize = 12;
        box.Paragraphs[0].Runs[0].Underline = rich;
        var result = presentation.ToPdfDocumentResult(new PowerPointToPdfOptions { ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic() });
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(result.ToBytes());
        Assert.Contains("VISIBLE", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)));
        var warning = Assert.Single(result.Warnings, warning => warning.Code == "text-box-overflow");
        Assert.Equal(PdfCore.PdfLayoutDiagnosticKind.ClippedContent, warning.LayoutDiagnostic!.Kind);
        Assert.True(warning.LayoutDiagnostic.HasBounds);
    }

    [Theory]
    [InlineData("plain")]
    [InlineData("rich")]
    [InlineData("paragraph")]
    public void ConfiguredFallbackFontsApplyToUnstyledGlyphsButPreserveAuthoredFonts(string flow) {
        using var presentation = PowerPointPresentation.Create(new MemoryStream());
        PowerPointSlide slide = presentation.AddSlide();
        PowerPointTextBox box = slide.AddTextBoxPoints("DEFAULT", 20, 20, 220, 80);
        box.FontSize = 12;
        if (flow == "rich") box.Paragraphs[0].Runs[0].Underline = true;
        if (flow == "paragraph") box.Paragraphs[0].SetBullet('*');
        PowerPointTextBox authored = slide.AddTextBoxPoints("AUTHORED", 20, 120, 220, 50);
        authored.FontName = "Courier";
        authored.FontSize = 12;
        byte[] bytes = presentation.ToPdfBytes(new PowerPointToPdfOptions {
            FontFamily = "serif", ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        });
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var letters = pdf.GetPage(1).Letters;
        string text = string.Concat(letters.Select(letter => letter.Value));
        int defaultIndex = text.IndexOf("DEFAULT", StringComparison.Ordinal);
        int authoredIndex = text.IndexOf("AUTHORED", StringComparison.Ordinal);
        Assert.True(defaultIndex >= 0 && authoredIndex >= 0, text);
        Assert.Contains("Times", letters[defaultIndex].FontName, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Courier", letters[authoredIndex].FontName, StringComparison.OrdinalIgnoreCase);
    }
}
