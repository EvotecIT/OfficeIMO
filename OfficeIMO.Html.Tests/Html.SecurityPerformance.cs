using System;
using System.Diagnostics;
using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void HtmlToWord_SharedAncestorStylesAreParsedWithinBoundedTime() {
        const int paragraphCount = 200;
        var ancestorStyle = new StringBuilder("background-color:#0000ff;");
        while (ancestorStyle.Length < 64 * 1024) {
            ancestorStyle.Append("--padding-contract:0;");
        }
        var html = new StringBuilder("<div style=\"")
            .Append(ancestorStyle)
            .Append("\">");
        for (int index = 0; index < paragraphCount; index++) {
            html.Append("<p style=\"background-color:rgba(255,0,0,0.5)\">Text</p>");
        }
        html.Append("</div>");
        HtmlToWordOptions options = HtmlToWordOptions.CreateUntrustedHtmlProfile();
        var stopwatch = Stopwatch.StartNew();

        using WordDocument document = HtmlConversionDocument.Parse(html.ToString())
            .ToWordDocumentResult(options)
            .RequireValue();
        stopwatch.Stop();

        Assert.Equal(paragraphCount, document.Paragraphs.Count);
        Assert.All(document.Paragraphs, paragraph =>
            Assert.Equal("800080", paragraph.ShadingFillColorHex));
        Assert.True(
            stopwatch.Elapsed < TimeSpan.FromSeconds(15),
            $"HTML conversion took {stopwatch.Elapsed}.");
    }
}
