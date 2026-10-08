using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.OpenDocument.Testing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed class OpenDocumentShapeTextGrowthTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void WidthGrowthHonorsHardBreaksAndKeepsSourceAndFixedHeight(bool flat, bool wrap) {
        var document = Create("Alpha beta gamma delta END\nShort", height: 60); var box = document.Pages[0].Shapes[0];
        box.AutoGrowWidth = true; box.WrapText = wrap;
        box.Paragraphs[0].MarginLeft = OdfLength.Points(8); box.Paragraphs[0].MarginRight = OdfLength.Points(6);
        box.Paragraphs[0].EnsureStyle().TextIndent = OdfLength.Points(4); box.Paragraphs[0].TextAlign = "right";
        using var stream = new MemoryStream();
        if (flat) document.SaveFlatXml(stream); else document.Save(stream);
        stream.Position = 0; var loaded = flat ? OdgDocument.LoadFlatXml(stream) : OdgDocument.Load(stream);
        string[] before = XmlState(loaded);
        var result = loaded.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        var paint = Paint(result); var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>());
        Assert.InRange(paint.Shape.Width, 160, 350); Assert.Equal(60, paint.Shape.Height);
        Assert.Equal(paint.Shape.Width, text.Width); Assert.Equal(paint.Shape.Height, text.Height);
        Assert.All(text.Paragraphs.SelectMany(p => p.Runs), r => Assert.Equal(12, r.FontSize));
        var svg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(result.Value));
        var lines = svg.Descendants(XNamespace.Get("http://www.w3.org/2000/svg") + "text").Select(e => e.Value).ToArray();
        Assert.Contains("Alpha beta gamma delta END", lines); Assert.Contains("Short", lines);
        Assert.Equal(40, loaded.Pages[0].Shapes[0].Bounds.Width.ToPoints()); Assert.Equal(before, XmlState(loaded));
    }

    [Fact]
    public void BothAxesGrowWidthBeforeHeightAndOnlyWrapAfterWidthCap() {
        var document = Create(string.Join(" ", Enumerable.Repeat("Alpha beta gamma delta", 4)), height: 24);
        var page = document.Pages[0]; var box = page.Shapes[0]; box.AutoGrowWidth = true; box.AutoGrowHeight = true;
        var probe = page.ToDrawing();
        Assert.True(!probe.Report.HasSkippedOrUnsupported, string.Join("\n", probe.Report.Mappings.Select(m => m.Feature + " " + m.Status + " " + m.Message)));
        var unrestricted = Paint(probe);
        Assert.True(unrestricted.Shape.Width > 400); Assert.Equal(24, unrestricted.Shape.Height);
        box.TextBoxMinimumWidth = OdfLength.Points(40); box.TextBoxMaximumWidth = OdfLength.Points(160);
        box.TextBoxMinimumHeight = OdfLength.Points(24); box.TextBoxMaximumHeight = OdfLength.Points(400);
        string[] before = XmlState(document);
        var capped = Paint(page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(160, capped.Shape.Width); Assert.InRange(capped.Shape.Height, 24.001, 399);
        Assert.Equal(before, XmlState(document));
        box.TextBoxMaximumHeight = OdfLength.Points(24);
        var clipped = page.ToDrawing(); Assert.Equal(24, Paint(clipped).Shape.Height); AssertClipped(clipped.Report);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WidthCapReportsHorizontalOrVerticalClippingWhenHeightIsFixed(bool wrap) {
        var document = Create("Alpha beta gamma delta END", height: 24); var page = document.Pages[0]; var box = page.Shapes[0];
        box.AutoGrowWidth = true; box.WrapText = wrap;
        box.TextBoxMinimumWidth = OdfLength.Points(40); box.TextBoxMaximumWidth = OdfLength.Points(60);
        string[] before = XmlState(document); var result = page.ToDrawing();
        Assert.Equal(60, Paint(result).Shape.Width); Assert.Equal(24, Paint(result).Shape.Height); AssertClipped(result.Report);
        Assert.Throws<OdfConversionLossException>(() => page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported));
        Assert.Equal(before, XmlState(document));
    }

    [Fact]
    public void WidthMeasurementUsesCompleteBodyAndCurrentFontBeforeAttachmentResolution() {
        var document = Create(OdfTextFittingTestFonts.Body, height: 40); var page = document.Pages[0]; var box = page.Shapes[0];
        box.AutoGrowWidth = true; box.Paragraphs[0].FontFamily = OdfTextFittingTestFonts.Family;
        var target = page.Shapes.AddRectangle(new OdfRect(OdfLength.Points(500), OdfLength.Points(30), OdfLength.Points(40), OdfLength.Points(40)), "Target");
        var connector = page.Shapes.AddConnector(new OfficePoint(60, 50), new OfficePoint(500, 50));
        connector.StrokeColor = OdfColor.Parse("#204060");
        connector.AttachStart(box.AddGluePoint(OdgGluePointAlignment.Right)); connector.AttachEnd(target.AddGluePoint(OdgGluePointAlignment.Left));
        string[] before = XmlState(document); double[] widths = new double[2];
        foreach (bool wide in new[] { false, true }) {
            var result = page.ToDrawing(OdfTextFittingTestFonts.Profile(wide), OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
            var text = Assert.Single(result.Value.Elements.OfType<OfficeDrawingRichText>()); widths[wide ? 1 : 0] = text.Width;
            var line = Assert.Single(result.Value.Elements.OfType<OfficeDrawingShape>(), s => s.Shape.Kind == OfficeShapeKind.Line);
            Assert.InRange(Math.Abs(line.X + line.Shape.Points[0].X - (20 + text.Width)), 0, .001);
        }
        Assert.True(widths[1] > widths[0]); Assert.Equal(before, XmlState(document));
    }

    [Fact]
    public void BothAxesCanGrowFromZeroInstanceMinima() {
        var document = Create("Body", height: 24); var box = document.Pages[0].Shapes[0];
        box.AutoGrowWidth = true; box.AutoGrowHeight = true;
        box.TextBoxMinimumWidth = OdfLength.Points(0); box.TextBoxMinimumHeight = OdfLength.Points(0);
        var result = document.Pages[0].ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        Assert.InRange(Paint(result).Shape.Width, 1, 39.999); Assert.InRange(Paint(result).Shape.Height, 1, 23.999);
    }

    [Fact]
    public void RightAlignedUnderlinedWidthGrowthContainsCompletePaint() {
        var document = Create("Alpha beta gamma delta END", height: 60); var page = document.Pages[0]; var box = page.Shapes[0];
        box.AutoGrowWidth = true; box.WrapText = false;
        box.Paragraphs[0].TextAlign = "right"; box.Paragraphs[0].Underline = true;
        string[] before = XmlState(document);
        var result = page.ToDrawing(OdfConversionLossPolicy.ThrowOnSkippedOrUnsupported);
        Assert.False(result.Report.HasSkippedOrUnsupported); Assert.True(Paint(result).Shape.Width > 40);
        Assert.Contains("Alpha beta gamma delta END", OfficeDrawingSvgExporter.ToSvg(result.Value));
        Assert.Equal(before, XmlState(document));
    }

    [Fact]
    public void UnsupportedWidthGrowthKeepsOrdinaryFixedFrameAlignment() {
        var document = Create("Alpha beta gamma delta END", height: 60); var page = document.Pages[0]; var box = page.Shapes[0];
        box.WrapText = false; box.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Middle;
        box.Paragraphs[0].TextAlign = "right"; box.Paragraphs[0].Underline = true;
        string ordinary = OfficeDrawingSvgExporter.ToSvg(page.ToDrawing().Value);
        box.AutoGrowWidth = true; string[] before = XmlState(document);
        var result = page.ToDrawing();
        Assert.Contains(result.Report.Mappings, m => m.Feature.EndsWith(":text-auto-size", StringComparison.Ordinal) &&
            m.Status == OdfConversionMappingStatus.Unsupported);
        Assert.Equal(ordinary, OfficeDrawingSvgExporter.ToSvg(result.Value));
        Assert.Equal(40, Paint(result).Shape.Width); Assert.Equal(before, XmlState(document));
    }

    private static OdgDocument Create(string text, double height) {
        var document = OdgDocument.Create(); var box = document.AddPage().Shapes.AddTextBox(new OdfRect(
            OdfLength.Points(20), OdfLength.Points(30), OdfLength.Points(40), OdfLength.Points(height)), text, "Growth");
        box.AutoGrowWidth = false; box.AutoGrowHeight = false; box.WrapText = true;
        box.TextVerticalAlignment = OdfTextAreaVerticalAlignment.Top;
        box.TextPadding = new OdfInsets(OdfLength.Points(2), OdfLength.Points(2), OdfLength.Points(2), OdfLength.Points(2));
        box.Paragraphs[0].FontSize = OdfLength.Points(12); box.FillColor = OdfColor.Parse("#e0f0ff");
        return document;
    }
    private static OfficeDrawingShape Paint(OdfConversionResult<OfficeDrawing> result) => result.Value.Elements.OfType<OfficeDrawingShape>().Single(s => s.Shape.Kind == OfficeShapeKind.Rectangle);
    private static string[] XmlState(OdfDocument document) => new[] { document.GetXml("content.xml").ToString(), document.GetXml("styles.xml").ToString() };
    private static void AssertClipped(OdfConversionReport report) => Assert.Contains(report.Mappings,
        m => m.Feature.EndsWith(":text-clipped", StringComparison.Ordinal) && m.Status == OdfConversionMappingStatus.Unsupported);
}
