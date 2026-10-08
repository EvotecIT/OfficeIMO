using System;
using System.Linq;
using System.Threading;
using System.Xml;
using OfficeIMO.Html;
using OfficeIMO.Html.Dom;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlXmlBoundaryTests {
    private const string Start = "<html xmlns='http://www.w3.org/1999/xhtml'><head/>";

    [Theory]
    [InlineData("xlink", "")]
    [InlineData("link", "")]
    [InlineData("xlink", "xmlns:other='urn:other' other:href='https://example.test/'")]
    [InlineData("link", "xmlns:xlink='urn:other' xlink:href='https://example.test/'")]
    public void XmlSvgHyperlinksUseNamespaceAwarePolicy(string prefix, string extraAttributes) {
        string source = "\n" + Start + "<body><svg xmlns='http://www.w3.org/2000/svg' xmlns:" + prefix +
            "='http://www.w3.org/1999/xlink'><a " + extraAttributes + " " + prefix + ":href='javascript:alert(1)'><text>Open</text></a></svg></body></html>\n";
        var document = HtmlConversionDocument.ParseXhtml(source, new HtmlConversionDocumentOptions());
        var link = Assert.Single(document.ResourceManifest.Resources);
        Assert.Equal(HtmlResourceKind.Hyperlink, link.Kind);
        Assert.False(link.IsAllowed);
    }

    [Fact]
    public void XmlEmptyElementsStaySiblingsWithoutChangingHtmlParsing() {
        string source = Start + "<body><div id='a'/><div id='b'/></body></html>";
        var xml = HtmlConversionDocument.ParseXhtml(source, new HtmlConversionDocumentOptions());
        Assert.Equal(2, xml.Document.QuerySelectorAll("body > div").Count);
        Assert.Single(HtmlConversionDocument.Parse(source).Document.QuerySelectorAll("body > div"));
        Assert.Equal(2, xml.Edit(_ => { }).Document.QuerySelectorAll("body > div").Count);
    }

    [Fact]
    public void XmlRawTextAndForeignNamespacesSurviveAnalysis() {
        string source = "<html xmlns='http://www.w3.org/1999/xhtml'><head><style>" +
            "div { background-image:url('image.png?a=1&amp;b=2'); }</style></head><body><div/>" +
            "<svg xmlns='http://www.w3.org/2000/svg' xmlns:xlink='http://www.w3.org/1999/xlink'>" +
            "<linearGradient id='g'/><use xlink:href='#g'/></svg>" +
            "<math xmlns='http://www.w3.org/1998/Math/MathML'><mi>x</mi></math>" +
            "<script><![CDATA[if (a < b && c) {}]]></script></body></html>";
        var document = HtmlConversionDocument.ParseXhtml(source, new HtmlConversionDocumentOptions());
        Assert.Contains("a=1&b=2", document.Document.QuerySelector("style")!.TextContent);
        Assert.Equal("if (a < b && c) {}", document.Document.QuerySelector("script")!.TextContent);
        var gradient = document.Document.QuerySelector("svg")!.ChildNodes.OfType<HtmlElement>().First();
        Assert.Equal("linearGradient", gradient.LocalName);
        Assert.Equal("http://www.w3.org/2000/svg", gradient.NamespaceUri);
        Assert.Equal("#g", document.Document.QuerySelector("use")!.GetAttribute("http://www.w3.org/1999/xlink", "href"));
        Assert.Equal("http://www.w3.org/1998/Math/MathML", document.Document.QuerySelector("math")!.NamespaceUri);
        Assert.Contains(document.ResourceManifest.Resources, r => r.Source.Contains("a=1&b=2"));
    }

    [Fact]
    public void XmlLimitsCountActualElementDepthAndAllNodes() {
        string source = Start + "<body><div/><div/></body></html>";
        var options = new HtmlConversionDocumentOptions();
        options.Limits.MaxHtmlDepth = 3;
        Assert.Equal(2, HtmlConversionDocument.ParseXhtml(source, options).Document.QuerySelectorAll("div").Count);
        options.Limits.MaxHtmlDepth = 2;
        Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.ParseXhtml(source, options));
        options.Limits.MaxHtmlDepth = 3;
        options.Limits.MaxHtmlNodes = 4;
        Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.ParseXhtml(source, options));
        options.Limits.MaxInputCharacters = 1;
        Assert.Throws<HtmlDomLimitException>(() => HtmlConversionDocument.ParseXhtml(source, options));
        Assert.Throws<OperationCanceledException>(() => HtmlConversionDocument.ParseXhtml(source,
            new HtmlConversionDocumentOptions(), new CancellationToken(true)));
    }

    [Theory]
    [InlineData("<!DOCTYPE html [<!ENTITY injected 'payload'>]>")]
    [InlineData("<!DOCTYPE html [<!ENTITY injected SYSTEM 'file:///unavailable-epub-test'>]>")]
    public void XmlDoesNotExpandDtdEntities(string declaration) {
        Assert.Throws<XmlException>(() => HtmlConversionDocument.ParseXhtml(
            declaration + Start + "<body>&injected;</body></html>", new HtmlConversionDocumentOptions()));
    }

    [Fact]
    public void SvgArchiveResourceDiscoveryPreservesXmlCssEntities() {
        var bytes = System.Text.Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg'>" +
            "<style>@import url('paint.css?a=1&amp;b=2');</style><rect/></svg>");
        var manifest = HtmlResourcePipeline.BuildSvgArchiveManifest(bytes, new Uri("https://example.test/book.svg"),
            new HtmlResourcePipelineOptions());
        Assert.Contains(manifest.Resources, r => r.Source.Contains("a=1&b=2"));
        Assert.DoesNotContain(manifest.Resources, r => r.Source.Contains("&amp;"));
    }
}
