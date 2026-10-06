using OfficeIMO.Drawing;
using OfficeIMO.Epub;
using OfficeIMO.Epub.Image;
using OfficeIMO.Html;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class EpubFixedLayoutInspectionTests {
    [Theory]
    [InlineData(20, false)]
    [InlineData(180, true)]
    public void RenderedChildGeometryIsComparedWithDeclaredCanvas(int top, bool overflow) {
        var book = Book(top);
        var options = new EpubImageExportOptions { ViewportWidth = 900, ViewportHeight = 1000, Margins = HtmlRenderMargins.All(25) };
        var result = Read(book).InspectFixedLayoutPage(0, options);
        Assert.Equal(200, result.ViewportWidth); Assert.Equal(200, result.ViewportHeight);
        Assert.Equal(overflow, result.HasCanvasOverflow);
        Assert.Equal(900, options.ViewportWidth); Assert.Equal(25, options.Margins.Top);
        Assert.Contains(result.Rendering.Pages.Single().Visuals, visual => visual.Height >= 80);
    }

    [Theory]
    [InlineData("width=device-width,height=200")]
    [InlineData("width=200,width=300,height=200")]
    [InlineData("width=200")]
    [InlineData("width=0,height=200")]
    [InlineData("width=40000,height=200")]
    public void InvalidOrOverBudgetViewportCannotReturnAnApparentlySuccessfulInspection(string viewport) {
        var book = Book(20); var xml = book.GetContentXml("page");
        xml.Descendants().Single(e => (string?)e.Attribute("name") == "viewport").SetAttributeValue("content", viewport);
        book.SetContentXml("page", xml);
        Assert.Throws<InvalidDataException>(() => Read(book).InspectFixedLayoutPage(0));
    }

    [Fact]
    public void TextFallbackAndReflowableChaptersCannotEstablishFixedGeometry() {
        var book = Book(20);
        var textOnly = EpubDocument.Load(new MemoryStream(book.Write().Bytes));
        Assert.Throws<NotSupportedException>(() => textOnly.InspectFixedLayoutPage(0));
        var reflow = EpubPublication.Create("Reflow", "en"); reflow.AddChapter("page", "EPUB/page.xhtml", "Page", "<p>Text</p>");
        Assert.Throws<NotSupportedException>(() => Read(reflow).InspectFixedLayoutPage(0));
    }

    [Fact]
    public void MissingResourcesRemainVisibleInTheInspection() {
        var book = Book(20); book.AddResource("image", "EPUB/image.png", "image/png", Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNg+P//HwAF/gL9HjcXBgAAAABJRU5ErkJggg=="));
        var xml = book.GetContentXml("page"); xml.Root!.Element(XName.Get("body", "http://www.w3.org/1999/xhtml"))!.Add(
            new XElement(XName.Get("img", "http://www.w3.org/1999/xhtml"), new XAttribute("src", "image.png"), new XAttribute("alt", "Pixel")));
        book.SetContentXml("page", xml);
        var source = EpubDocument.Load(new MemoryStream(book.Write().Bytes), new EpubReadOptions { IncludeRawHtml = true, IncludeResourceData = false });
        var result = source.InspectFixedLayoutPage(0);
        Assert.True(result.HasRenderingWarnings);
        Assert.Contains(result.Rendering.Diagnostics, d => d.Code == HtmlRenderDiagnosticCodes.ResourceUnavailable);
    }

    [Fact]
    public void SharedBoundsInspectionUsesTargetCanvasAndHonorsCancellation() {
        var drawing = new OfficeDrawing(300, 300).AddShape(OfficeShape.Rectangle(80, 80), 160, 160);
        Assert.False(OfficeDrawingQualityAnalyzer.Analyze(drawing).HasIssues);
        Assert.Contains(OfficeDrawingQualityAnalyzer.Analyze(drawing, 200, 200).Issues, i => i.Kind == OfficeDrawingQualityIssueKind.ElementOutsideBounds);
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeDrawingQualityAnalyzer.Analyze(drawing, double.NaN, 200));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeDrawingQualityAnalyzer.Analyze(drawing, 200, 200, cancellationToken: cancellation.Token));
        Assert.Throws<OperationCanceledException>(() => Read(Book(20)).InspectFixedLayoutPage(0, cancellationToken: cancellation.Token));
    }

    [Theory]
    [InlineData("position:relative;left:-60px", true)]
    [InlineData("transform:translateX(200px)", true)]
    public void NegativeAndTransformedOverflowRemainDetectable(string extra, bool overflow) {
        var book = Book(20); var xml = book.GetContentXml("page");
        xml.Descendants().Single(e => e.Attribute("style") != null).SetAttributeValue("style", "width:80px;height:80px;background:#124e80;" + extra);
        book.SetContentXml("page", xml);
        Assert.Equal(overflow, Read(book).InspectFixedLayoutPage(0).HasCanvasOverflow);
    }

    [Fact]
    public void AuthoredClippingAndFractionalCanvasDoNotCreateSpuriousOverflow() {
        var book = Book(180); var xml = book.GetContentXml("page");
        xml.Descendants().Single(e => (string?)e.Attribute("id") == "frame").SetAttributeValue("style", "overflow:hidden");
        book.SetContentXml("page", xml);
        Assert.False(Read(book).InspectFixedLayoutPage(0).HasCanvasOverflow);
        xml = book.GetContentXml("page");
        xml.Descendants().Single(e => (string?)e.Attribute("name") == "viewport").SetAttributeValue("content", "width=200.5,height=200.5");
        book.SetContentXml("page", xml);
        Assert.False(Read(book).InspectFixedLayoutPage(0).HasCanvasOverflow);
    }

    private static EpubPublication Book(int top) {
        var book = EpubPublication.Create("Fixed geometry", "en");
        book.AddChapter("page", "EPUB/page.xhtml", "Page", "<div id='frame'><div style='width:80px;height:80px;background:#124e80'></div></div>");
        book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(200, 200) { Regions = new[] { new EpubFixedLayoutRegion("frame", 20, top, 100, 20) } });
        return book;
    }
    private static EpubDocument Read(EpubPublication book) => EpubDocument.Load(new MemoryStream(book.Write().Bytes),
        new EpubReadOptions { IncludeRawHtml = true, IncludeResourceData = true });
}
