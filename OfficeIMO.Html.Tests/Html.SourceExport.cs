using System;
using OfficeIMO.Html;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class HtmlSourceExportTests {
    [Theory]
    [InlineData("")]
    [InlineData("<base target=\"_blank\">")]
    [InlineData("<base href=\"../assets/\" target=\"_blank\">")]
    [InlineData("<base href=\"https://assets.example.test/reports/\">")]
    public void ExportRetainsEffectiveBaseWithoutMutatingTheSource(string baseTag) {
        string html = "<!DOCTYPE html><!--source--><html><head>" + baseTag +
            "<link rel=\"stylesheet\" href=\"report.css\"><style>p { color: red; }</style></head>" +
            "<body><img src=\"marker.png\"><a href=\"details.html\">Details</a><p>Report</p></body></html>";
        var source = HtmlConversionDocument.Parse(html, new HtmlConversionDocumentOptions {
            BaseUri = new Uri("file:///reports/input%20folder/report.html")
        });
        string exported = source.ExportSourceHtml();
        var relocated = HtmlConversionDocument.Parse(exported, new HtmlConversionDocumentOptions {
            BaseUri = new Uri("file:///unrelated/output/copied.html")
        });
        Assert.Equal(source.BaseUri, relocated.BaseUri);
        Assert.Equal(new Uri(source.BaseUri!, "marker.png"), new Uri(relocated.BaseUri!, "marker.png"));
        Assert.Equal(html, source.SourceHtml);
        Assert.Equal(exported, source.ExportSourceHtml());
        Assert.StartsWith("<!DOCTYPE html>", exported, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<!--source-->", exported);
        Assert.Contains("href=\"report.css\"", exported);
        Assert.Contains("p { color: red; }", exported);
        if (baseTag.Contains("_blank")) Assert.Contains("target=\"_blank\"", exported);
    }

    [Fact]
    public void ExportWithoutABaseReturnsExactSourceText() {
        const string html = "<p data-title='one'>One &amp; two</p>";
        Assert.Equal(html, HtmlConversionDocument.Parse(html).ExportSourceHtml());
    }

    [Fact]
    public void ExportEscapesTheBaseAttributeAndPreservesAFragment() {
        var source = HtmlConversionDocument.Parse("<p id='report'>Report</p>", new HtmlConversionDocumentOptions {
            BaseUri = new Uri("https://example.test/report?x=1&y=two")
        });
        string exported = source.ExportSourceHtml();
        Assert.Contains("x=1&amp;y=two", exported);
        Assert.Equal(source.BaseUri, HtmlConversionDocument.Parse(exported).BaseUri);
        Assert.Contains("id=\"report\"", exported);
    }
}
