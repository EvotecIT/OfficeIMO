using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Theory]
    [InlineData("bottom-left")]
    [InlineData("top-center")]
    [InlineData("top-left-corner")]
    [InlineData("left-middle")]
    public void HtmlRender_PageMarginTextWrapsWithoutDroppingItsLastWords(string position) {
        string html = "<style>@page { size:300px 300px; margin:60px; @" + position
            + " { content:'Report Created Today'; font:10px Arial; } } body,p { margin:0; }</style><p>BodyOnly</p>";
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged };
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText margin = Assert.Single(page.Visuals.OfType<HtmlRenderText>(), text => text.SemanticRole == "page-margin");
        Assert.True(margin.Height >= 2D * margin.LineHeight);
        Assert.Equal("Report Created Today", margin.Text);
        Assert.InRange(margin.X, 0D, page.Width - margin.Width);
        Assert.InRange(margin.Y, 0D, page.Height - margin.Height);

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(options));
        PdfCore.PdfReadDocument document = PdfCore.PdfReadDocument.Open(pdf,
            new PdfCore.PdfLoadOptions { IncludeArtifactText = true });
        string actual = string.Concat(document.ExtractText().Where(c => !char.IsWhiteSpace(c)));
        Assert.Contains("ReportCreatedToday", actual, StringComparison.Ordinal);
        PdfCore.PdfReadDocument body = PdfCore.PdfReadDocument.Open(pdf);
        Assert.Equal("BodyOnly", body.ExtractText().Trim());
    }

    [Fact]
    public void HtmlRender_PageMarginTextClipsAtTheReservedMarginHeight() {
        const string html = """
            <style>
              @page { size:300px 200px; margin:12px;
                @bottom-left { content:'Report Created Today With More Words Than Fit'; font:10px Arial; }
              }
              body,p { margin:0; }
            </style><p>Body</p>
            """;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderClipGroup clip = Assert.Single(page.Visuals.OfType<HtmlRenderClipGroup>(),
            group => group.Source == "@page @bottom-left");
        Assert.Equal(12D, clip.Height, 3);
        Assert.Equal(page.Height - 12D, clip.Y, 3);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported
            && diagnostic.Source == "@page @bottom-left" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }
}
