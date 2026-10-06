using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlPdf_PageMarginTimestampPaintsWithinItsWrappedFrame() {
        const string footer = "Data as of 21 Sep 2026 01:59 UTC · Generated 1 Oct 2026 22:41 UTC";
        string html = "<style>@page { size:A4; margin:16mm 14mm 18mm; "
            + "@bottom-left { content:'" + footer + "'; font:8pt/1.2 Calibri, Carlito, 'Segoe UI', sans-serif; } "
            + "@bottom-center {content:'Internal';font:8pt sans-serif} "
            + "@bottom-right {content:'Page 1';font:8pt sans-serif} } body,p {margin:0}</style><p>Body</p>";
        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes();
        var page = Assert.Single(PdfCore.PdfReadDocument.Open(pdf,
            new PdfCore.PdfLoadOptions { IncludeArtifactText = true }).Pages);
        var spans = page.GetTextSpans().Where(span => span.Y < 50D && span.X < 210D).ToArray();
        string actual = string.Concat(string.Concat(spans.Select(span => span.Text)).Where(c => !char.IsWhiteSpace(c)));
        Assert.Equal(string.Concat(footer.Where(c => !char.IsWhiteSpace(c))), actual);
        Assert.True(spans.Select(span => Math.Round(span.Y, 2)).Distinct().Count() >= 2,
            "The full footer must paint on multiple baselines rather than extending into its peer margin boxes.");
        Assert.All(spans, span => Assert.InRange(span.X + span.Advance, 0D, 211.6536D));
        HtmlRenderPage rendered = Assert.Single(HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged }).Pages);
        var drawingLines = rendered.CreateDrawing().Elements.OfType<OfficeDrawingText>()
            .Where(text => text.Y > 1000D && text.X < 280D).ToArray();
        Assert.True(drawingLines.Select(text => text.Y).Distinct().Count() >= 2);
        Assert.Equal(string.Concat(footer.Where(c => !char.IsWhiteSpace(c))),
            string.Concat(string.Concat(drawingLines.Select(text => text.Text)).Where(c => !char.IsWhiteSpace(c))));
    }

    [Theory]
    [InlineData("top-left")]
    [InlineData("top-center")]
    [InlineData("top-right")]
    public void HtmlRender_SolePageMarginTextUsesTheAvailableHeaderWidth(string position) {
        string html = "<style>@page { size:260px 200px; margin:24px; @" + position
            + " {content:'Generated 1 Oct 2026 at 22:41 UTC';font:10px Arial} } body,p {margin:0}</style><p>BodyOnly</p>";
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, FidelityPolicy = HtmlRenderFidelityPolicy.RequireNoLoss };
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText margin = Assert.Single(page.Visuals.OfType<HtmlRenderText>(), text => text.SemanticRole == "page-margin");
        Assert.Equal(page.Width - page.Margins.Left - page.Margins.Right, margin.Width, 3);
        Assert.Equal(margin.LineHeight, margin.Height, 3);
        string text = PdfCore.PdfReadDocument.Open(HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(options)),
            new PdfCore.PdfLoadOptions { IncludeArtifactText = true }).ExtractText();
        Assert.Contains("Generated 1 Oct 2026 at 22:41 UTC", text, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("bottom-left")]
    [InlineData("top-center")]
    [InlineData("top-left-corner")]
    [InlineData("left-middle")]
    public void HtmlRender_PageMarginTextWrapsWithoutDroppingItsLastWords(string position) {
        string html = "<style>@page { size:300px 300px; margin:60px; @" + position
            + " { content:'Report Created Today'; font:10px Arial; } "
            + (position == "bottom-left" ? "@bottom-center {content:'I';font:10px Arial} @bottom-right {content:'R';font:10px Arial}" : string.Empty)
            + (position == "top-center" ? "@top-left {content:'L';font:10px Arial} @top-right {content:'R';font:10px Arial}" : string.Empty)
            + " } body,p { margin:0; }</style><p>BodyOnly</p>";
        var options = new HtmlRenderOptions { Mode = HtmlRenderMode.Paged };
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, options);
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText margin = Assert.Single(page.Visuals.OfType<HtmlRenderText>(), text => text.Source == "@page @" + position);
        Assert.True(margin.Height >= 2D * margin.LineHeight);
        Assert.Equal("Report Created Today", margin.Text);
        Assert.InRange(margin.X, 0D, page.Width - margin.Width);
        Assert.InRange(margin.Y, 0D, page.Height - margin.Height);

        byte[] pdf = HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions(options));
        PdfCore.PdfReadDocument document = PdfCore.PdfReadDocument.Open(pdf,
            new PdfCore.PdfLoadOptions { IncludeArtifactText = true });
        var marginSpans = Assert.Single(document.Pages).GetTextSpans().Where(span =>
            span.X >= margin.X * 0.75D - 0.01D && span.X < (margin.X + margin.Width) * 0.75D - 0.01D
            && span.Y <= (page.Height - margin.Y) * 0.75D + 0.01D
            && span.Y >= (page.Height - margin.Y - margin.Height) * 0.75D - 0.01D);
        string actual = string.Concat(string.Concat(marginSpans.Select(span => span.Text)).Where(c => !char.IsWhiteSpace(c)));
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
                @bottom-center { content:'I'; font:10px Arial; }
                @bottom-right { content:'R'; font:10px Arial; }
              }
              body,p { margin:0; }
            </style><p>Body</p>
            """;
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        HtmlRenderPage page = Assert.Single(rendered.Pages);
        HtmlRenderText margin = Assert.Single(page.Visuals.OfType<HtmlRenderText>(),
            text => text.Source == "@page @bottom-left");
        Assert.Equal(12D, margin.Height, 3);
        Assert.Equal(page.Height - 12D, margin.Y, 3);
        Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported
            && diagnostic.Source == "@page @bottom-left" && diagnostic.LossKind == OfficeConversionLossKind.Omission);
    }
}
