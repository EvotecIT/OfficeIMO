using OfficeIMO.Epub;
using OfficeIMO.Epub.Image;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using System.Xml.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class EpubTextInkInspectionTests {
    [Theory]
    [InlineData(700, "", false, false)]
    [InlineData(2000, "", true, false)]
    [InlineData(2000, "color:transparent", false, false)]
    [InlineData(2000, "opacity:0", false, false)]
    [InlineData(2000, "overflow:hidden", false, true)]
    [InlineData(700, "transform:translateY(-50px)", true, false)]
    public void TextInkReportsOverhangSeparatelyFromLayoutBoxes(int glyphHeight, string style, bool overflow, bool unmeasured) {
        var book = EpubPublication.Create("Text ink", "en");
        book.AddChapter("page", "EPUB/page.xhtml", "Page", "<div id='frame'>A</div>");
        book.AddResource("font", "EPUB/ink.ttf", "font/ttf", ManagedTextShapingTestAssets.CreateFontWithTallGlyph('A', glyphHeight));
        book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(200, 200) {
            Regions = new[] { new EpubFixedLayoutRegion("frame", 20, 0, 100, 30) }
        });
        var xml = book.GetContentXml("page"); XNamespace html = "http://www.w3.org/1999/xhtml";
        xml.Root!.Element(html + "head")!.Add(new XElement(html + "style",
            "@font-face{font-family:Ink;src:url('ink.ttf')}#frame{font-family:Ink;font-size:20px;line-height:20px;" + style + "}"));
        book.SetContentXml("page", xml);
        var source = EpubDocument.Load(new MemoryStream(book.Write().Bytes),
            new EpubReadOptions { IncludeRawHtml = true, IncludeResourceData = true });
        var report = source.InspectFixedLayoutPage(0);
        Assert.NotEmpty(report.Rendering.Fonts.Faces);
        Assert.Equal(overflow, report.HasTextInkOverflow);
        Assert.Equal(unmeasured, report.TextInkDiagnostics.Any(d => d.Code == HtmlRenderDiagnosticCodes.TextInkNotInspected));
        if (overflow || unmeasured) Assert.True(report.HasRenderingWarnings);
        if (style.Length == 0) Assert.False(report.HasCanvasOverflow);
    }
}
