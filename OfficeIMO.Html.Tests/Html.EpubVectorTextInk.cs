using OfficeIMO.Epub;
using OfficeIMO.Epub.Image;
using OfficeIMO.Html;
using OfficeIMO.TestAssets;
using Xunit;
using System.Xml.Linq;

namespace OfficeIMO.Tests;

public sealed class EpubVectorTextInkTests {
    [Theory]
    [InlineData("<rect width='30' height='30' fill='red'/>", false)]
    [InlineData("<text x='10' y='40' font-size='20' font-family='Ink'>A</text>", false)]
    [InlineData("<g transform='translate(20,10)'><text x='10' y='40' font-size='20' font-family='Ink'>A</text></g>", false)]
    [InlineData("<text x='95' y='40' font-size='20' font-family='Ink'>A</text>", true)]
    public void InlineSvgIsInspectedInsteadOfBlanketUnmeasured(string content, bool cropped) {
        var book = EpubPublication.Create("Vector text", "en");
        book.AddResource("font", "EPUB/ink.ttf", "font/ttf", ManagedTextShapingTestAssets.CreateFont('A'));
        book.AddChapter("page", "EPUB/page.xhtml", "Page", "<svg xmlns='http://www.w3.org/2000/svg' width='100' height='100' viewBox='0 0 100 100'>" + content + "</svg>");
        book.SetFixedLayoutPage("page", new EpubFixedLayoutPage(200, 200));
        var xml = book.GetContentXml("page"); XNamespace html = "http://www.w3.org/1999/xhtml";
        xml.Root!.Element(html + "head")!.Add(new XElement(html + "style", "@font-face{font-family:Ink;src:url('ink.ttf')}"));
        book.SetContentXml("page", xml);
        var source = EpubDocument.Load(new MemoryStream(book.Write().Bytes), new EpubReadOptions { IncludeRawHtml = true, IncludeResourceData = true });
        var report = source.InspectFixedLayoutPage(0);
        Assert.Contains(report.Rendering.Pages[0].Visuals, v => v is HtmlRenderDrawing);
        Assert.DoesNotContain(report.TextInkDiagnostics, d => d.Code == HtmlRenderDiagnosticCodes.TextInkNotInspected);
        Assert.False(report.HasTextInkOverflow);
        if (content.Contains("<text")) Assert.NotEmpty(report.Rendering.Fonts.Faces);
        Assert.Equal(cropped, report.TextInkDiagnostics.Any(d => d.Code == HtmlRenderDiagnosticCodes.ClippedTextInkBounds));
    }
}
