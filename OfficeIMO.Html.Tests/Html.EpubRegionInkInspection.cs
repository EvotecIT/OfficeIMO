using OfficeIMO.Epub;
using OfficeIMO.Epub.Image;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class EpubRegionInkInspectionTests {
    [Theory]
    [InlineData(700, "", "", false, false)]
    [InlineData(2000, "", "", true, false)]
    [InlineData(2000, "transform:rotate(20deg)", "", true, false)]
    [InlineData(700, "transform:translateY(100px)", "", false, false)]
    [InlineData(700, "", "transform:translateY(35px)", true, false)]
    [InlineData(2000, "overflow:hidden", "", false, true)]
    [InlineData(2000, "opacity:0", "", false, false)]
    public void RegionInkUsesItsOwnBoxAndRetainsDescendantEffects(int height, string regionStyle,
        string childStyle, bool overflow, bool unmeasured) {
        var report = Read(Book(height, regionStyle, childStyle)).InspectFixedLayoutRegions(0, new[] { "frame" });
        var region = Assert.Single(report.Regions);
        Assert.Equal("frame", region.ElementId);
        Assert.Equal(overflow, region.HasTextInkOverflow);
        Assert.Equal(unmeasured, region.TextInkDiagnostics.Any(d => d.Code == HtmlRenderDiagnosticCodes.TextInkNotInspected));
        if (overflow || unmeasured) Assert.True(report.HasRenderingWarnings);
        if (regionStyle.Length == 0 && childStyle.Length == 0) {
            Assert.False(report.HasTextInkOverflow);
            Assert.False(region.HasOverflow);
        }
        if (overflow) Assert.Contains(region.TextInkDiagnostics, d => d.Code == HtmlRenderDiagnosticCodes.TextInkOutsideRegion);
        Assert.DoesNotContain(region.TextInkDiagnostics, d => d.Code == HtmlRenderDiagnosticCodes.TextInkOutsideCanvas);
    }

    [Fact]
    public void RegionInkExcludesAncestorClipsAndSiblingTextWithoutChangingPaint() {
        var book = Book(2000, "", ""); var xml = book.GetContentXml("page");
        XNamespace html = "http://www.w3.org/1999/xhtml";
        var frame = xml.Descendants().Single(e => (string?)e.Attribute("id") == "frame");
        frame.SetAttributeValue("style", "position:absolute;left:50px;top:0px;width:100px;height:30px");
        var wrapper = new XElement(html + "div", new XAttribute("style", "position:relative;width:200px;height:20px;overflow:hidden;transform:rotate(10deg)"));
        frame.ReplaceWith(wrapper); wrapper.Add(frame);
        xml.Root!.Element(html + "body")!.Add(new XElement(html + "div", new XAttribute("id", "other"),
            new XAttribute("style", "position:absolute;left:150px;top:150px;width:20px;height:20px")));
        book.SetContentXml("page", xml); var source = Read(book);
        var plain = source.InspectFixedLayoutPage(0);
        var selected = source.InspectFixedLayoutRegions(0, new[] { "other", "frame" });
        Assert.False(selected.Regions[0].HasTextInkOverflow);
        Assert.Empty(selected.Regions[0].TextInkDiagnostics);
        Assert.True(selected.Regions[1].HasTextInkOverflow);
        Assert.DoesNotContain(selected.Regions[1].TextInkDiagnostics, d => d.Code == HtmlRenderDiagnosticCodes.TextInkNotInspected);
        Assert.Equal(plain.Rendering.Pages[0].CreateDrawing().ExportImage(OfficeIMO.Drawing.OfficeImageExportFormat.Png).Bytes,
            selected.Rendering.Pages[0].CreateDrawing().ExportImage(OfficeIMO.Drawing.OfficeImageExportFormat.Png).Bytes);
    }

    private static EpubPublication Book(int glyphHeight, string regionStyle, string childStyle) {
        var book = EpubPublication.Create("Region text ink", "en");
        book.AddChapter("page", "EPUB/page.xhtml", "Page", "<div id='frame'><div id='text'>A</div></div>");
        book.AddResource("font", "EPUB/ink.ttf", "font/ttf", ManagedTextShapingTestAssets.CreateFontWithTallGlyph('A', glyphHeight));
        book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(200, 200) {
            Regions = new[] { new EpubFixedLayoutRegion("frame", 50, 50, 100, 30) }
        });
        var xml = book.GetContentXml("page"); XNamespace html = "http://www.w3.org/1999/xhtml";
        xml.Root!.Element(html + "head")!.Add(new XElement(html + "style",
            "@font-face{font-family:Ink;src:url('ink.ttf')}#frame{font-family:Ink;font-size:20px;line-height:20px;" + regionStyle + "}#text{" + childStyle + "}"));
        book.SetContentXml("page", xml); return book;
    }
    private static EpubDocument Read(EpubPublication book) => EpubDocument.Load(new MemoryStream(book.Write().Bytes),
        new EpubReadOptions { IncludeRawHtml = true, IncludeResourceData = true });
}
