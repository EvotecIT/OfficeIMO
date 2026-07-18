using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlGeneratedContent_RendersStyledBeforeAfterTextAndAttributes() {
        const string html = """
            <style>
              .note::before { content:"Before "; color:#123456; position:relative; left:4px; }
              .note:after { content:" " attr(data-suffix); color:#654321; }
            </style>
            <p class="note" data-suffix="After" style="margin:0">Body</p>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 240D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderText before = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Source == "p.note::before");
        HtmlRenderText body = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Body");
        HtmlRenderText after = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Source == "p.note::after");

        Assert.Equal("Before ", before.Text);
        Assert.Equal(" After", after.Text);
        Assert.Equal("generated-before", before.SemanticRole);
        Assert.Equal("generated-after", after.SemanticRole);
        Assert.Equal(OfficeColor.FromRgb(0x12, 0x34, 0x56), before.Color);
        Assert.Equal(OfficeColor.FromRgb(0x65, 0x43, 0x21), after.Color);
        Assert.Equal(4D, before.X, 3);
        Assert.True(before.X < body.X);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported);
    }

    [Fact]
    public void HtmlGeneratedContent_UsesCascadeSpecificityImportantAndLegacyPseudoSyntax() {
        const string html = """
            <style>
              #target::before { content:"Specific"; }
              .label::before { content:"Class"; }
              p::before { content:"Important" !important; }
              .label:after { content:" Legacy"; }
            </style>
            <p id="target" class="label" style="margin:0">Body</p>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 240D,
            Margins = HtmlRenderMargins.All(0D)
        });

        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Important" && text.Source == "p#target::before");
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == " Legacy" && text.Source == "p#target::after");
        Assert.DoesNotContain(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Specific" || text.Text == "Class");
    }

    [Fact]
    public void HtmlGeneratedContent_ResolvesNestedCountersAndCounterStyles() {
        const string html = """
            <style>
              body { counter-reset:section; }
              section { counter-increment:section; }
              section section { counter-reset:section; }
              section::before { content:counters(section, ".", upper-roman) " "; }
              h2 { counter-increment:item; }
              section { counter-reset:item; }
              h2::before { content:counter(item, decimal-leading-zero) ": "; }
            </style>
            <section><h2>Outer</h2><section><h2>Inner</h2></section></section>
            <section><h2>Second</h2></section>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 300D,
            Margins = HtmlRenderMargins.All(0D)
        });
        IReadOnlyList<HtmlRenderText> generated = rendered.Pages.SelectMany(page => page.Visuals)
            .OfType<HtmlRenderText>()
            .Where(text => text.SemanticRole == "generated-before")
            .ToList();

        Assert.Contains(generated, text => text.Text == "I");
        Assert.Contains(generated, text => text.Text == "I.I");
        Assert.Contains(generated, text => text.Text == "II");
        Assert.Equal(2, generated.Count(text => text.Text == "01: "));
        Assert.Contains(generated, text => text.Text == "02: ");
    }

    [Fact]
    public void HtmlGeneratedContent_RendersAroundBlockChildrenAndAtTheBodyBoundary() {
        const string html = """
            <style>
              body::before { content:"DocumentStart"; display:block; }
              article::before { content:"ArticleStart"; display:block; background:#ffeecc; }
              article::after { content:"ArticleEnd"; display:block; }
              body::after { content:"DocumentEnd"; display:block; }
            </style>
            <article><div>ChildBlock</div></article>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 240D,
            Margins = HtmlRenderMargins.All(0D)
        });
        IReadOnlyList<HtmlRenderText> text = rendered.Pages[0].Visuals.OfType<HtmlRenderText>().ToList();

        Assert.True(IndexOfText(text, "DocumentStart") < IndexOfText(text, "ArticleStart"));
        Assert.True(IndexOfText(text, "ArticleStart") < IndexOfText(text, "ChildBlock"));
        Assert.True(IndexOfText(text, "ChildBlock") < IndexOfText(text, "ArticleEnd"));
        Assert.True(IndexOfText(text, "ArticleEnd") < IndexOfText(text, "DocumentEnd"));
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape => shape.Source == "article::before" && shape.Shape.FillColor == OfficeColor.FromRgb(0xFF, 0xEE, 0xCC));
    }

    [Fact]
    public void HtmlGeneratedContent_PreservesLinkAndTableCellOwnership() {
        const string link = "https://example.test/generated";
        const string html = """
            <style>
              a::before { content:"["; }
              a::after { content:"]"; }
              a.block-link { display:block; }
              td::before { content:attr(data-label) ": "; font-weight:bold; }
            </style>
            <a href="https://example.test/generated">Linked</a>
            <a class="block-link" href="https://example.test/generated"><div>BlockLinked</div></a>
            <table><tr><td data-label="Total">42</td></tr></table>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 240D,
            Margins = HtmlRenderMargins.All(0D)
        });
        IReadOnlyList<HtmlRenderText> generatedLink = rendered.Pages[0].Visuals.OfType<HtmlRenderText>()
            .Where(text => text.Source != null && text.Source.StartsWith("a", StringComparison.Ordinal) && text.Source.Contains("::", StringComparison.Ordinal))
            .ToList();
        HtmlRenderText cellPrefix = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Source == "td::before");

        Assert.Equal(4, generatedLink.Count);
        Assert.All(generatedLink, text => Assert.Equal(link, text.LinkUri));
        Assert.Equal("Total: ", cellPrefix.Text);
        Assert.True((cellPrefix.Font.Style & OfficeFontStyle.Bold) != 0);
    }

    [Fact]
    public void HtmlGeneratedContent_FlowsThroughPngSvgAndSearchablePdf() {
        const string html = "<style>.marker::before{content:'Generated\\20';color:#123456}</style><p class='marker' style='margin:0'>BackendMarker</p>";
        var imageOptions = new HtmlRenderOptions {
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(8D)
        };

        OfficeImageExportResult png = HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Png, imageOptions);
        string svg = Encoding.UTF8.GetString(HtmlConversionDocument.Parse(html).ExportImage(OfficeImageExportFormat.Svg, imageOptions).Bytes);
        HtmlPdfSaveOptions pdfOptions = new HtmlPdfSaveOptions();
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdf(pdfOptions)).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Bytes.Take(8));
        Assert.Contains("Generated", svg, StringComparison.Ordinal);
        Assert.Contains("BackendMarker", svg, StringComparison.Ordinal);
        Assert.Contains("GeneratedBackendMarker", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlGeneratedContent_DiagnosesUnsupportedContentFunctions() {
        const string html = """
            <style>
              .image::before { content:url('data:image/png;base64,AA=='); }
              .quote::after { content:open-quote; }
              .flex::before { content:"FlexFallback"; display:flex; }
            </style>
            <p class="image">ImageFallback</p><p class="quote">QuoteFallback</p><p class="flex">FlexHost</p>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        IReadOnlyList<HtmlDiagnostic> diagnostics = rendered.Diagnostics
            .Where(diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported)
            .ToList();

        Assert.Equal(2, diagnostics.Count);
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "p.image::before" && diagnostic.Detail!.Contains("url", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(diagnostics, diagnostic => diagnostic.Source == "p.quote::after" && diagnostic.Detail!.Contains("open-quote", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Source == "p.flex::before" && text.Text == "FlexFallback");
        Assert.Single(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending && diagnostic.Source == "p.flex::before");
        Assert.DoesNotContain(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text =>
            text.Source == "p.image::before" || text.Source == "p.quote::after");
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.GeneratedContentUnsupported, out _));
    }

    [Fact]
    public void HtmlGeneratedContent_DiagnosesUnsupportedCounterDeclarationsAndStyles() {
        const string html = """
            <style>
              .declaration { --bad-counter:chapter 1 2; counter-reset:var(--bad-counter); }
              .declaration::before { content:counter(chapter) " "; }
              .style { --bad-content:counter(chapter, symbols("*")); }
              .style::before { content:var(--bad-content); }
            </style>
            <p class="declaration">DeclarationFallback</p><p class="style">StyleFallback</p>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));

        Assert.Single(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedCounterUnsupported);
        Assert.Single(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported);
        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Source == "p.declaration::before" && text.Text == "0 ");
        Assert.DoesNotContain(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Source == "p.style::before");
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.GeneratedCounterUnsupported, out _));
    }

    [Fact]
    public void HtmlGeneratedContent_UsesTheSharedLayoutDepthLimit() {
        string html = "<style>div::before{content:'x'}</style>"
            + string.Concat(Enumerable.Repeat("<div>", 8))
            + "Leaf"
            + string.Concat(Enumerable.Repeat("</div>", 8));

        HtmlDomLimitException exception = Assert.Throws<HtmlDomLimitException>(() =>
            HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions { MaxLayoutDepth = 3 }));

        Assert.Equal(HtmlRenderDiagnosticCodes.DepthLimitExceeded, exception.Code);
        Assert.Equal(nameof(HtmlRenderOptions.MaxLayoutDepth), exception.LimitSource);
        Assert.Equal(3, exception.Limit);
    }

    private static int IndexOfText(IReadOnlyList<HtmlRenderText> text, string value) {
        for (int index = 0; index < text.Count; index++) {
            if (text[index].Text == value) return index;
        }

        return -1;
    }
}
