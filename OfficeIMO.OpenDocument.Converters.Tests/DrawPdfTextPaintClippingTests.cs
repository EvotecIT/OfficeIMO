using System;
using System.Linq;
using OfficeIMO.OpenDocument;
using OfficeIMO.OpenDocument.Odg.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class DrawPdfTextPaintClippingTests {
    [Theory]
    [InlineData("underline", false)]
    [InlineData("double", false)]
    [InlineData("wave", false)]
    [InlineData("background", false)]
    [InlineData("underline", true)]
    [InlineData("double", true)]
    [InlineData("wave", true)]
    [InlineData("background", true)]
    public void FixedFrameChecksCompletePaintAndShrinksItWhenRequested(string kind, bool shrink) {
        foreach (double height in new[] { 11D, 20D }) {
            var source = Create(height, "ABC"); var shape = source.Pages[0].Shapes[0];
            var paragraph = shape.Paragraphs[0];
            if (kind == "background") paragraph.BackgroundColor = OdfColor.Parse("#FFE080");
            else {
                paragraph.Underline = true;
                if (kind == "double") paragraph.UnderlineType = OdfTextDecorationType.Double;
                if (kind == "wave") paragraph.UnderlineStyle = OdfTextDecorationStyle.Wave;
            }
            if (shrink) shape.TextFitMode = OdfTextFitMode.ShrinkToFit;
            Verify(source, clipped: height == 11 && !shrink, "ABC");
        }
    }

    [Theory]
    [InlineData("label", 11, true)]
    [InlineData("label", 20, false)]
    [InlineData("leader", 11, true)]
    [InlineData("leader", 20, false)]
    public void FormattedLabelsAndGlyphLeadersUseTheSameCompletePaintCheck(string kind, double height, bool clipped) {
        var source = Create(height, ""); var shape = source.Pages[0].Shapes[0];
        if (kind == "label") {
            // A text box starts with an empty ordinary paragraph; this fixture
            // contains only a formatted list label, with no preceding line.
            source.GetXml("content.xml").Descendants(OdfNamespaces.Draw + "text-box").Single()
                .Element(OdfNamespaces.Text + "p")!.Remove();
            shape.AddList(true).AddItem("");
            var paragraph = shape.Paragraphs.Single(); SetCaptionFont(paragraph); paragraph.Underline = true;
            Verify(source, clipped, "1.");
        } else {
            var paragraph = shape.Paragraphs[0]; paragraph.Text = "A\tB";
            var style = source.Styles.CreateNamed("UnderlineLeader", OdfStyleFamily.Text); style.Underline = true;
            paragraph.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(80)).WithLeader(".").WithLeaderTextStyle(style) });
            Verify(source, clipped, "B");
        }
    }

    private static OdgDocument Create(double height, string text) {
        var source = OdgDocument.Create();
        var shape = source.AddPage().Shapes.AddTextBox(new OdfRect(OdfLength.Points(20), OdfLength.Points(30),
            OdfLength.Points(180), OdfLength.Points(height)), text);
        SetCaptionFont(shape.Paragraphs[0]); shape.WrapText = false;
        shape.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Top;
        shape.TextPadding = new OdfInsets(OdfLength.Points(0), OdfLength.Points(0), OdfLength.Points(0), OdfLength.Points(0));
        return source;
    }

    private static void SetCaptionFont(OdfTextParagraph paragraph) {
        paragraph.FontFamily = "Helvetica"; paragraph.FontSize = OdfLength.Points(12); paragraph.LineHeight = OdfLength.Points(4);
    }

    private static void Verify(OdgDocument source, bool clipped, string marker) {
        string[] before = Xml(source); var result = source.ToPdfDocumentResult();
        var report = Assert.IsType<OdfConversionReport>(Assert.Single(result.SourceConversionReports));
        Assert.Equal(clipped, report.Mappings.Any(m => m.Feature.EndsWith(":text-clipped", StringComparison.Ordinal)
            && m.Status == OdfConversionMappingStatus.Unsupported));
        var strict = new OdgToPdfOptions { LossPolicy = OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported };
        if (clipped) Assert.Throws<OdfConversionLossException>(() => source.ToPdfBytes(strict));
        else Assert.Contains(marker, PdfReadDocument.Open(source.ToPdfBytes(strict)).ExtractText());
        Assert.Equal(before, Xml(source));
    }
    private static string[] Xml(OdgDocument source) => new[] { source.GetXml("content.xml").ToString(), source.GetXml("styles.xml").ToString() };
}
