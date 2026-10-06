using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsSearchablePdfTests {
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void PdfRetainsUnicodeClustersAndSpacesWithoutChangingPaint(XpsFormat format) {
        var doc = Create(format); var page = doc.Pages[0];
        var xml = page.GetMarkup(); var run = xml.Elements().Single();
        run.SetAttributeValue("UnicodeString", "fi A😀");
        run.SetAttributeValue("Indices", "(2:1)36,60;3,25;36,60;(2:1)36,60");
        page.ReplaceMarkup(xml);
        var spans = page.ToSvg().TextSpans;
        Assert.Equal(new[] { "fi ", "A", "😀" }, spans.Select(s => s.Text));
        var bytes = doc.ToPdf();
        var pdf = PdfReadDocument.Open(bytes);
        Assert.Equal("fi A😀", pdf.ExtractText().Trim());
        var expected = OfficeDrawingRasterRenderer.Render(page.ToDrawing());
        var actual = OfficeDrawingRasterRenderer.Render(PdfDocument.Load(bytes).Render.Drawing(1), 96D / 72D);
        for (int y = 0; y < expected.Height; y++) for (int x = 0; x < expected.Width; x++)
            Assert.InRange(Math.Abs(expected.GetPixel(x, y).R - actual.GetPixel(x, y).R), 0, 8);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeTransformsCarryClusterSelectionGeometry(bool sideways) {
        var doc = Create(XpsFormat.OpenXps); var page = doc.Pages[0];
        var xml = page.GetMarkup(); var run = xml.Elements().Single();
        run.SetAttributeValue("IsSideways", sideways ? "true" : "false");
        run.SetAttributeValue("UnicodeString", "A"); page.ReplaceMarkup(xml);
        var original = Assert.Single(page.ToSvg().TextSpans);
        run.Remove(); run.SetAttributeValue("RenderTransform", "1,0.2,0.3,1,5,6");
        xml.Add(new XElement(xml.Name.Namespace + "Canvas", new XAttribute("RenderTransform", "0,1,-1,0,140,10"), run));
        page.ReplaceMarkup(xml);
        var transformed = Assert.Single(page.ToSvg().TextSpans);
        var matrix = new OfficeTransform(1, .2, .3, 1, 5, 6).Then(new OfficeTransform(0, 1, -1, 0, 140, 10));
        Assert.Equal(matrix.TransformPoint(original.TopLeft), transformed.TopLeft);
        Assert.Equal(matrix.TransformPoint(original.BottomRight), transformed.BottomRight);
        var pdf = PdfReadDocument.Open(doc.ToPdf());
        Assert.Equal("A", pdf.ExtractText().Trim());
    }

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void SidewaysSelectionUsesRunCellsAndRetainsExplicitOffsets(XpsFormat format) {
        var doc = Create(format); var page = doc.Pages[0];
        var xml = page.GetMarkup(); var run = xml.Elements().Single();
        run.SetAttributeValue("IsSideways", "true");
        run.SetAttributeValue("UnicodeString", "Searchable"); page.ReplaceMarkup(xml);
        var spans = page.ToSvg().TextSpans;
        Assert.All(spans, span => {
            Assert.Equal(spans[0].TopLeft.Y, span.TopLeft.Y);
            Assert.Equal(spans[0].BottomLeft.Y, span.BottomLeft.Y);
        });
        run.SetAttributeValue("UnicodeString", "AA");
        run.SetAttributeValue("Indices", "36,100;36,100,0,25"); page.ReplaceMarkup(xml);
        spans = page.ToSvg().TextSpans;
        Assert.Equal(spans[0].TopLeft.Y - 6, spans[1].TopLeft.Y, 8);
        Assert.Equal("AA", PdfReadDocument.Open(doc.ToPdf()).ExtractText().Replace("\n", "").Trim());
    }

    [Fact]
    public void RtlAndZeroAdvanceClustersKeepNativePositions() {
        var page = Create(XpsFormat.Xps).Pages[0]; var xml = page.GetMarkup(); var run = xml.Elements().Single();
        run.SetAttributeValue("UnicodeString", "אב"); run.SetAttributeValue("Indices", "36,100;36,0,25,10");
        run.SetAttributeValue("BidiLevel", "1"); page.ReplaceMarkup(xml);
        var spans = page.ToSvg().TextSpans;
        Assert.Equal("אב", string.Concat(spans.Select(s => s.Text)));
        Assert.True(spans[0].BottomRight.X < spans[0].BottomLeft.X);
        Assert.True(spans[1].BottomRight.X < spans[1].BottomLeft.X);
        Assert.True(spans[1].TopLeft.X < spans[0].TopLeft.X);
        Assert.NotEmpty(page.Document.ToPdf());
    }

    [Fact]
    public void DecorativeVisualBrushTextDoesNotBecomeDuplicateSearchContent() {
        var doc = Create(XpsFormat.Xps); var page = doc.Pages[0]; var xml = page.GetMarkup();
        var run = xml.Elements().Single(); run.Remove(); var ns = xml.Name.Namespace;
        xml.Add(new XElement(ns + "Path", new XAttribute("Data", "M0,0H200V100H0Z"),
            new XElement(ns + "Path.Fill", new XElement(ns + "VisualBrush",
                new XAttribute("Viewbox", "0,0,200,100"), new XAttribute("Viewport", "0,0,200,100"),
                new XAttribute("ViewboxUnits", "Absolute"), new XAttribute("ViewportUnits", "Absolute"),
                new XElement(ns + "VisualBrush.Visual", run)))));
        page.ReplaceMarkup(xml);
        Assert.Empty(page.ToSvg().TextSpans);
        var pdf = PdfReadDocument.Open(doc.ToPdf()); Assert.Equal("", pdf.ExtractText().Trim());
    }

    [Fact]
    public void OffsetsMoveSelectionAndSingularTransformsRemoveIt() {
        var page = Create(XpsFormat.Xps).Pages[0]; var xml = page.GetMarkup(); var run = xml.Elements().Single();
        run.SetAttributeValue("UnicodeString", "A"); run.SetAttributeValue("Indices", "36,100"); page.ReplaceMarkup(xml);
        var original = Assert.Single(page.ToSvg().TextSpans);
        run.SetAttributeValue("Indices", "36,100,50,25"); page.ReplaceMarkup(xml);
        var shifted = Assert.Single(page.ToSvg().TextSpans);
        Assert.Equal(original.TopLeft.X + 12, shifted.TopLeft.X, 8);
        Assert.Equal(original.TopLeft.Y - 6, shifted.TopLeft.Y, 8);
        run.SetAttributeValue("RenderTransform", "0,0,0,0,0,0"); page.ReplaceMarkup(xml);
        Assert.Empty(page.ToSvg().TextSpans);
    }

    [Fact]
    public void SourceTextSurvivesClippingButGlyphOnlyRunsDoNotInventUnicode() {
        var doc = Create(XpsFormat.Xps); var page = doc.Pages[0]; var xml = page.GetMarkup(); var run = xml.Elements().Single();
        run.SetAttributeValue("Clip", "M0,0H1V1H0Z"); run.SetAttributeValue("Opacity", "0"); page.ReplaceMarkup(xml);
        Assert.Equal("Hello", PdfReadDocument.Open(doc.ToPdf()).ExtractText().Trim());
        run.Attribute("UnicodeString")!.Remove(); run.SetAttributeValue("Indices", "36,100"); page.ReplaceMarkup(xml);
        Assert.Empty(page.ToSvg().TextSpans);
        Assert.Equal("", PdfReadDocument.Open(doc.ToPdf()).ExtractText().Trim());
    }

    [Theory]
    [InlineData("  A  B  ")]
    [InlineData("    ")]
    public void WhitespaceMergingPreservesExactSourceText(string text) {
        var page = Create(XpsFormat.Xps).Pages[0]; var xml = page.GetMarkup();
        xml.Elements().Single().SetAttributeValue("UnicodeString", text); page.ReplaceMarkup(xml);
        Assert.Equal(text, string.Concat(page.ToSvg().TextSpans.Select(s => s.Text)));
        Assert.NotEmpty(page.Document.ToPdf());
    }

    [Fact]
    public void BlankPagesRetainTheirSequenceAndSizes() {
        var doc = XpsDocument.Create(); doc.AddPage(100, 200); doc.AddPage(300, 400);
        var pdf = PdfReadDocument.Open(doc.ToPdf());
        Assert.Equal(2, pdf.Pages.Count);
        Assert.Equal((75D, 150D), pdf.Pages[0].GetPageSize());
        Assert.Equal((225D, 300D), pdf.Pages[1].GetPageSize());
    }

    private static XpsDocument Create(XpsFormat format) {
        var doc = XpsDocument.Create(format);
        string font = doc.AddFont(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "RobotoFlex.ttf")));
        doc.AddPage(200, 160).AddText("Hello", font, 24, 30, 70);
        return doc;
    }
}
