using OfficeIMO.Drawing;
using OfficeIMO.Epub;
using OfficeIMO.Epub.Image;
using OfficeIMO.Html;
using System.Xml.Linq;
using System.Threading;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class EpubFixedLayoutInspectionTests {
    [Fact]
    public void EmptyXhtmlSiblingsPaintLikeExplicitlyClosedElements() {
        var book = Book(40); var xml = book.GetContentXml("page"); XNamespace html = "http://www.w3.org/1999/xhtml";
        var frame = xml.Descendants().Single(e => (string?)e.Attribute("id") == "frame");
        frame.ReplaceNodes(new XElement(html + "div", new XAttribute("style", "width:80px;height:10px;background:#124e80")),
            new XElement(html + "div", new XAttribute("style", "width:40px;height:10px;background:#804e12")));
        book.SetContentXml("page", xml);
        byte[] empty = Read(book).InspectFixedLayoutPage(0).Rendering.Pages[0].CreateDrawing()
            .ExportImage(OfficeImageExportFormat.Png).Bytes;
        foreach (var child in frame.Elements()) child.Add(new XText(string.Empty));
        Assert.All(frame.Elements(), child => Assert.False(child.IsEmpty));
        book.SetContentXml("page", xml);
        byte[] closed = Read(book).InspectFixedLayoutPage(0).Rendering.Pages[0].CreateDrawing()
            .ExportImage(OfficeImageExportFormat.Png).Bytes;
        Assert.Equal(closed, empty);
    }

    [Fact]
    public void ExcessiveClippingWorkFailsInsteadOfReturningPartialDiagnostics() {
        var book = Book(20); var xml = book.GetContentXml("page"); XNamespace html = "http://www.w3.org/1999/xhtml";
        var frame = xml.Descendants().Single(e => (string?)e.Attribute("id") == "frame"); frame.RemoveNodes();
        for (int i = 0; i < 1025; i++) frame.Add(new XElement(html + "div", new XAttribute("style", "height:1px;overflow:hidden"),
            new XElement(html + "div", new XAttribute("style", "height:2px;background:#124e80"))));
        book.SetContentXml("page", xml);
        var error = Assert.Throws<NotSupportedException>(() => Read(book).InspectFixedLayoutPage(0));
        Assert.Contains("1024-clip", error.Message);
    }

    [Theory]
    [InlineData(80, 80, "overflow:hidden", true, false)]
    [InlineData(80, 10, "overflow:hidden", false, false)]
    [InlineData(80, 80, "overflow-x:clip;overflow-y:visible", false, false)]
    [InlineData(200, 10, "overflow-x:visible;overflow-y:clip", false, false)]
    [InlineData(200, 10, "overflow-x:clip;overflow-y:visible", true, false)]
    [InlineData(80, 80, "overflow:hidden;border-radius:10px", false, true)]
    [InlineData(80, 80, "overflow:hidden;transform:rotate(30deg)", true, false)]
    public void ClippingInspectionDistinguishesHiddenBoundsFromUnmeasuredPaths(int width, int height,
        string style, bool clipped, bool unmeasured) {
        var book = Book(20); var xml = book.GetContentXml("page");
        var frame = xml.Descendants().Single(e => (string?)e.Attribute("id") == "frame");
        frame.SetAttributeValue("style", style);
        frame.Elements().Single().SetAttributeValue("style", $"width:{width}px;height:{height}px;background:#124e80");
        book.SetContentXml("page", xml);
        var inspection = Read(book).InspectFixedLayoutPage(0);
        Assert.Equal(clipped, inspection.HasClippedElementBounds);
        Assert.Equal(unmeasured, inspection.ClippingDiagnostics.Any(d => d.Code == HtmlRenderDiagnosticCodes.ClipGeometryNotInspected));
        if (clipped) {
            var diagnostic = Assert.Single(inspection.ClippingDiagnostics, d => d.Code == HtmlRenderDiagnosticCodes.ClippedElementBounds);
            Assert.Equal("div#frame", diagnostic.Source);
            Assert.Equal(HtmlDiagnosticSeverity.Info, diagnostic.Severity);
        }
        if (unmeasured) Assert.True(inspection.HasRenderingWarnings);
    }

    [Theory]
    [InlineData(10, "", "", false)]
    [InlineData(80, "", "", true)]
    [InlineData(80, "overflow:hidden", "", false)]
    [InlineData(10, "transform:translateX(30px)", "", false)]
    [InlineData(10, "transform:rotate(30deg)", "", false)]
    [InlineData(80, "transform:rotate(30deg)", "", true)]
    [InlineData(10, "", "transform:translateY(30px)", true)]
    public void SelectedRegionMeasuresLocalContentWithoutLosingItsIdentity(int childHeight, string regionStyle, string childStyle, bool overflow) {
        var book = Book(20); var xml = book.GetContentXml("page");
        var frame = xml.Descendants().Single(e => (string?)e.Attribute("id") == "frame");
        frame.SetAttributeValue("style", regionStyle);
        frame.Elements().Single().SetAttributeValue("style", "width:80px;height:" + childHeight + "px;background:#124e80;" + childStyle);
        book.SetContentXml("page", xml);
        var source = Read(book); string? before = source.Chapters[0].Html;
        var report = source.InspectFixedLayoutRegions(0, new[] { "frame" });
        Assert.Equal("frame", Assert.Single(report.Regions).ElementId);
        Assert.Equal(overflow, report.Regions[0].HasOverflow);
        Assert.Equal(before, source.Chapters[0].Html);
        Assert.Empty(source.InspectFixedLayoutPage(0).Regions);
    }

    [Fact]
    public void RegionSelectionCannotSilentlySucceedWithoutGeometry() {
        var source = Read(Book(20));
        Assert.Throws<InvalidDataException>(() => source.InspectFixedLayoutRegions(0, new[] { "missing" }));
        Assert.Throws<ArgumentException>(() => source.InspectFixedLayoutRegions(0, new[] { "frame", "frame" }));
        var book = Book(20); var xml = book.GetContentXml("page");
        xml.Descendants().Single(e => (string?)e.Attribute("id") == "frame").SetAttributeValue("style", "display:none");
        book.SetContentXml("page", xml);
        Assert.Throws<NotSupportedException>(() => Read(book).InspectFixedLayoutRegions(0, new[] { "frame" }));
    }

    [Fact]
    public void RegionMarkersDoNotChangePaintingOrIncludeSiblingContent() {
        var book = Book(20); var xml = book.GetContentXml("page"); XNamespace html = "http://www.w3.org/1999/xhtml";
        xml.Root!.Element(html + "body")!.Add(new XElement(html + "div", new XAttribute("id", "other"),
            new XAttribute("style", "position:absolute;left:140px;top:100px;width:20px;height:20px;background:#804e12")));
        book.SetContentXml("page", xml); var source = Read(book);
        var plain = source.InspectFixedLayoutPage(0);
        var selected = source.InspectFixedLayoutRegions(0, new[] { "other", "frame" });
        Assert.Equal(new[] { "other", "frame" }, selected.Regions.Select(r => r.ElementId));
        Assert.False(selected.Regions[0].HasOverflow); Assert.True(selected.Regions[1].HasOverflow);
        Assert.Equal(plain.Rendering.Pages[0].CreateDrawing().ExportImage(OfficeImageExportFormat.Png).Bytes,
            selected.Rendering.Pages[0].CreateDrawing().ExportImage(OfficeImageExportFormat.Png).Bytes);
    }

    [Fact]
    public void AncestorClippingAndTransformsDoNotRedefineLocalRegionContainment() {
        var book = Book(20); var xml = book.GetContentXml("page"); XNamespace html = "http://www.w3.org/1999/xhtml";
        var frame = xml.Descendants().Single(e => (string?)e.Attribute("id") == "frame");
        frame.SetAttributeValue("style", "position:absolute;left:20px;top:20px;width:100px;height:20px");
        var wrapper = new XElement(html + "div", new XAttribute("style", "position:relative;width:100px;height:30px;overflow:hidden;transform:rotate(10deg)"));
        frame.ReplaceWith(wrapper); wrapper.Add(frame); book.SetContentXml("page", xml);
        Assert.True(Read(book).InspectFixedLayoutRegions(0, new[] { "frame" }).Regions[0].HasOverflow);
    }

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
