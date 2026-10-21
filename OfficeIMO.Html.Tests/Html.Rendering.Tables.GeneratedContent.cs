using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("table")]
    [InlineData("inline-table")]
    public void HtmlCssTable_PrintedLinkDestinationsSurviveAnonymousCellLayoutAndSavedPdf(string display) {
        const string destination = "https://example.test/report#summary";
        string html = "<style>@page{size:240px 160px;margin:12px}body{margin:0;font:14px/20px Arial}"
            + "a{display:" + display + "}a::after{content:' (' attr(href) ')';font-size:90%;overflow-wrap:break-word}</style>"
            + "<a href='" + destination + "'>Summary</a><p>End of contents.</p>";
        var options = new HtmlToPdfOptions(CssTableOptions());
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>().ToArray();
        string joined = string.Concat(text.Select(item => item.Text));
        Assert.Contains("Summary", joined);
        Assert.Contains("(" + destination + ")", joined);
        Assert.True(joined.IndexOf("Summary", StringComparison.Ordinal) < joined.IndexOf(destination, StringComparison.Ordinal));
        Assert.Contains("End of contents.", joined);
        HtmlRenderText[] generated = text.Where(item => item.Source == "a::after").ToArray();
        Assert.NotEmpty(generated);
        Assert.All(generated, item => Assert.Equal(destination, item.LinkUri));
        Assert.DoesNotContain(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported);

        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes(options);
        PdfCore.PdfReadDocument saved = PdfCore.PdfReadDocument.Open(bytes);
        string savedText = string.Concat(saved.Pages.SelectMany(page => page.GetTextSpans()).Select(span => span.Text));
        Assert.Contains("Summary", savedText);
        Assert.Contains("(" + destination + ")", savedText);
        Assert.True(savedText.IndexOf("Summary", StringComparison.Ordinal) < savedText.IndexOf(destination, StringComparison.Ordinal));
        Assert.Contains("End of contents.", savedText);
    }

    [Fact]
    public void HtmlCssTable_GeneratedAnonymousRowsHaveDistinctSavedStructureOwners() {
        const string html = "<style>body{margin:0}table{border-spacing:0}table::before{content:'BeforeReference'}"
            + "table::after{content:'AfterReference'}</style><table><tr><td>SourceBody</td></tr></table>";
        byte[] bytes = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(CssTableOptions()));
        PdfCore.PdfReadDocument saved = PdfCore.PdfReadDocument.Open(bytes);
        Assert.Single(saved.TaggedContent!.StructureElements, item => item.StructureType == "Table");
        Assert.Equal(3, saved.TaggedContent.StructureElements.Count(item => item.StructureType == "TR"));
        Assert.Equal(3, saved.TaggedContent.StructureElements.Count(item => item.StructureType == "TD"));
        string text = string.Concat(saved.Pages.SelectMany(page => page.GetTextSpans()).Select(span => span.Text));
        Assert.Contains("BeforeReference", text);
        Assert.Contains("SourceBody", text);
        Assert.Contains("AfterReference", text);
        Assert.True(text.IndexOf("BeforeReference", StringComparison.Ordinal) < text.IndexOf("SourceBody", StringComparison.Ordinal));
        Assert.True(text.IndexOf("SourceBody", StringComparison.Ordinal) < text.IndexOf("AfterReference", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("table", "display:table-row")]
    [InlineData("table-row", "position:absolute")]
    [InlineData("table-row-group", "float:left")]
    public void HtmlCssTable_UnsupportedGeneratedFormattingKeepsTypedOmission(string display, string generatedStyle) {
        string body = display switch {
            "table-row" => "<div style='display:table'><div id='target' style='display:table-row'><div style='display:table-cell'>Retained</div></div></div>",
            "table-row-group" => "<div style='display:table'><div id='target' style='display:table-row-group'><div style='display:table-row'><div style='display:table-cell'>Retained</div></div></div></div>",
            _ => "<div id='target' style='display:table'>Retained</div>"
        };
        string html = "<style>body{margin:0}#target::after{content:'Unsupported';" + generatedStyle + "}</style>" + body;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, CssTableOptions());
        Assert.Contains("Retained", rendered.Text);
        Assert.DoesNotContain("Unsupported", rendered.Text);
        Assert.Contains(rendered.Diagnostics, item => item.Code == HtmlRenderDiagnosticCodes.TableValueUnsupported
            && item.LossKind == OfficeConversionLossKind.Omission && item.Source == "div#target::after");
    }
}
