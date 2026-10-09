using OfficeIMO.Html;
using OfficeIMO.Html.Dom;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlQualifiedAttributesTests {
    [Fact]
    public void SvgXlinkAndXmlAttributesRetainTheirQualifiedNamesAcrossNormalizationAndOwnedImport() {
        const string source = "<svg><g xml:lang='ar'><use href='#modern' xlink:href='#legacy'/></g></svg>";
        var document = HtmlConversionDocument.Parse(source);
        HtmlElement original = document.Document.QuerySelector("use")!;
        Assert.Equal("#legacy", original.GetAttribute("http://www.w3.org/1999/xlink", "href"));
        Assert.Equal("#legacy", original.GetAttribute("xlink:href"));
        var normalized = document.CreateDocumentForConversion();
        HtmlElement converted = normalized.QuerySelector("use")!;
        Assert.Equal("#modern", converted.GetAttribute("href"));
        Assert.Equal("#legacy", converted.GetAttribute("http://www.w3.org/1999/xlink", "href"));
        Assert.Equal("#legacy", converted.GetAttribute("xlink:href"));
        Assert.Equal("ar", normalized.QuerySelector("g")!.GetAttribute("http://www.w3.org/XML/1998/namespace", "lang"));
        Assert.Contains("xlink:href=", normalized.DocumentElement!.OuterHtml);
        Assert.Contains("xml:lang=", normalized.DocumentElement.OuterHtml);
    }
}
