using System;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentOdgTextClippingTests {
    [Theory]
    [InlineData("left-margin")]
    [InlineData("right-margin")]
    [InlineData("both-margins")]
    [InlineData("tab")]
    public void UnwrappedFramesKeepFiniteParagraphAndTabConstraints(string constraint) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = page.Shapes.AddTextBox(Bounds(40, 30), "END_MARKER", "Constrained");
        Configure(shape, OdfTextAreaVerticalAlignment.Top);
        var paragraph = shape.Paragraphs[0];
        if (constraint == "tab") {
            paragraph.Text = "A\tB";
            paragraph.SetTabStops(new[] { new OdfTabStop(OdfLength.Points(60)) });
        } else {
            paragraph.MarginLeft = OdfLength.Points(constraint == "left-margin" ? 40 : constraint == "both-margins" ? 20 : 0);
            paragraph.MarginRight = OdfLength.Points(constraint == "right-margin" ? 40 : constraint == "both-margins" ? 20 : 0);
        }
        string[] before = XmlState(document); var result = page.ToDrawing();
        Assert.True(IsClipped(result.Report));
        Assert.Equal(paragraph.Text, Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()).PlainText);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        if (constraint != "tab") Assert.DoesNotContain("END_MARKER", SvgText(result.Value));
        Assert.Equal(before, XmlState(document));
    }

    [Theory]
    [InlineData("left")]
    [InlineData("center")]
    [InlineData("right")]
    public void OrdinaryUnwrappedBodyCanStillOverhangItsFiniteFrame(string alignment) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        var shape = page.Shapes.AddTextBox(Bounds(40, 30), "A long complete unwrapped caption", "Overhang");
        Configure(shape, OdfTextAreaVerticalAlignment.Top); shape.Paragraphs[0].TextAlign = alignment;
        var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        Assert.False(IsClipped(result.Report)); Assert.Contains(shape.Text, SvgText(result.Value));
    }

    [Theory]
    [InlineData("frame", 4, OdfTextAreaVerticalAlignment.Top, true)]
    [InlineData("frame", 20, OdfTextAreaVerticalAlignment.Top, false)]
    [InlineData("frame", 14, OdfTextAreaVerticalAlignment.Bottom, true)]
    [InlineData("frame", 40, OdfTextAreaVerticalAlignment.Middle, false)]
    [InlineData("rect", 4, OdfTextAreaVerticalAlignment.Top, true)]
    [InlineData("ellipse", 4, OdfTextAreaVerticalAlignment.Top, true)]
    [InlineData("line", 4, OdfTextAreaVerticalAlignment.Top, true)]
    [InlineData("connector", 4, OdfTextAreaVerticalAlignment.Top, true)]
    public void CondensedLineBoxesReportActualPlacedGlyphClipping(string kind, double height,
        OdfTextAreaVerticalAlignment vertical, bool clipped) {
        var document = OdgDocument.Create(); var page = document.AddPage();
        OdgShape shape = kind switch {
            "rect" => page.Shapes.AddRectangle(Bounds(140, height), "Body"),
            "ellipse" => page.Shapes.AddEllipse(Bounds(140, height), "Body"),
            "line" => page.Shapes.AddLine(OdfLength.Points(20), OdfLength.Points(30), OdfLength.Points(160), OdfLength.Points(30), "Body"),
            "connector" => page.Shapes.AddConnector(new OfficePoint(20, 30), new OfficePoint(160, 30)),
            _ => page.Shapes.AddTextBox(Bounds(140, height), "END_MARKER", "Body")
        };
        if (kind != "frame") shape.AddParagraph("END_MARKER");
        Configure(shape, vertical); shape.Paragraphs[0].LineHeight = OdfLength.Points(4);
        string[] before = XmlState(document); var result = page.ToDrawing();
        Assert.Equal(clipped, IsClipped(result.Report)); Assert.Contains("END_MARKER", SvgText(result.Value));
        if (clipped) Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        else Assert.False(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported).Report.HasSkippedOrUnsupported);
        Assert.Equal(before, XmlState(document));
    }

    private static void Configure(OdgShape shape, OdfTextAreaVerticalAlignment vertical) {
        shape.Paragraphs[0].FontSize = OdfLength.Points(12); shape.TextVerticalAlignment = vertical; shape.WrapText = false;
        shape.TextPadding = new OdfInsets(OdfLength.Points(0), OdfLength.Points(0), OdfLength.Points(0), OdfLength.Points(0));
    }
    private static bool IsClipped(OdfConversionReport report) => report.Mappings.Any(m =>
        m.Feature.EndsWith(":text-clipped", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
    private static OdfRect Bounds(double width, double height) => new(OdfLength.Points(20), OdfLength.Points(30), OdfLength.Points(width), OdfLength.Points(height));
    private static string SvgText(OfficeDrawing drawing) => string.Concat(XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing))
        .Descendants(XNamespace.Get("http://www.w3.org/2000/svg") + "text").Select(e => e.Value));
    private static string[] XmlState(OdfDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString() };
}
