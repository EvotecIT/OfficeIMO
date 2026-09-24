using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlGeneratedContent_EmptyInlineBlockPaintsItsBox() {
        const string html = """
            <style>body{margin:0}.badge::before{content:"";display:inline-block;width:16px;height:16px;background:#ff0000}</style>
            <span class="badge">Label</span>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 100D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape box = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderShape>(),
            shape => shape.Source == "span.badge::before" && shape.Shape.FillColor == OfficeColor.FromRgb(0xFF, 0, 0));
        HtmlRenderText label = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>(),
            text => text.Text == "Label");
        Assert.Equal(16D, box.Width, 3);
        Assert.Equal(16D, box.Height, 3);
        Assert.True(label.X >= box.X + box.Width);
    }

    [Fact]
    public void HtmlInlineAnchorWithBlockChildPreservesTextLink() {
        const string html = "<a href='https://example.com/card'><div>Linked card title</div></a>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 200D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderText title = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderText>(),
            text => text.Text == "Linked card title");
        Assert.Equal("https://example.com/card", title.LinkUri);
    }

    [Fact]
    public void HtmlGeneratedContent_EmptyBeforeReservesRatioWrapperHeight() {
        const string html = """
            <style>
              body { margin:0; }
              .ratio { width:320px; --bs-aspect-ratio:56.25%; }
              .ratio::before { content:""; display:block; padding-top:var(--bs-aspect-ratio); background:#ff0000; }
              p { margin:0; }
            </style>
            <div class="ratio"></div><p>After ratio</p>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 400D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderShape pseudo = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "div.ratio::before" && shape.Shape.FillColor == OfficeColor.FromRgb(0xFF, 0, 0));
        HtmlRenderText after = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(),
            text => text.Text == "After ratio");
        Assert.Equal(180D, pseudo.Height, 3);
        Assert.True(after.Y >= 180D, $"Following text started at {after.Y} before the ratio box ended.");
    }

    [Fact]
    public void HtmlGeneratedContent_AbsoluteAfterUsesItsPositionedHostWithoutAddingFlowHeight() {
        const string html = "<style>body,ul{margin:0;padding:0}ul{list-style:none}"
            + "li{position:relative;width:120px;height:30px;padding-right:20px;margin:0}"
            + "li::after{content:'';display:block;position:absolute;right:4px;top:10px;"
            + "width:8px;height:8px;background:#ff0000}</style>"
            + "<ul><li id='first'>First</li><li id='second'>Second</li></ul>"
            + "<p style='margin:0'>Following</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 240D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderVisual[] visuals = EnumerateRenderVisuals(rendered.Pages[0].Scene).ToArray();
        HtmlRenderShape firstArrow = Assert.Single(visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "li#first::after" && shape.Shape.FillColor == OfficeColor.Red);
        HtmlRenderShape secondArrow = Assert.Single(visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "li#second::after" && shape.Shape.FillColor == OfficeColor.Red);
        HtmlRenderText first = Assert.Single(visuals.OfType<HtmlRenderText>(), text => text.Text == "First");
        HtmlRenderText second = Assert.Single(visuals.OfType<HtmlRenderText>(), text => text.Text == "Second");
        HtmlRenderText following = Assert.Single(visuals.OfType<HtmlRenderText>(), text => text.Text == "Following");

        Assert.InRange(firstArrow.X - first.X, 125D, 130D);
        Assert.InRange(firstArrow.Y - first.Y, 9D, 11D);
        Assert.InRange(secondArrow.Y - firstArrow.Y, 29D, 31D);
        Assert.InRange(second.Y - first.Y, 29D, 31D);
        Assert.InRange(following.Y - first.Y, 59D, 61D);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Source == "li::after" && diagnostic.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported);
    }

    [Fact]
    public void HtmlGeneratedContent_AbsolutePseudosShareHostStackingBandsWithPositionedChildren() {
        const string html = "<style>body{margin:0}"
            + ".host{position:relative;width:40px;height:40px;margin:0}"
            + ".flow{width:40px;height:40px;background:#00ff00}"
            + "#before::before{content:'';display:block;position:absolute;left:0;top:0;width:40px;height:40px;background:#ff0000;z-index:1}"
            + "#positive{position:absolute;left:0;top:0;width:40px;height:40px;background:#ffff00;z-index:2}"
            + "#relative-positive{position:relative;top:-40px;width:40px;height:40px;background:#ff00ff;z-index:3}"
            + "#after::after{content:'';display:block;position:absolute;left:0;top:0;width:40px;height:40px;background:#0000ff;z-index:-1}"
            + "#relative-negative{position:relative;top:-40px;width:40px;height:40px;background:#00ffff;z-index:-2}"
            + "</style><div class='host' id='before'><div class='flow' id='first-flow'></div>"
            + "<div id='positive'></div><div id='relative-positive'></div></div>"
            + "<div class='host' id='after'><div class='flow' id='second-flow'></div><div id='relative-negative'></div></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 40D,
            ViewportHeight = 80D,
            Margins = HtmlRenderMargins.All(0D)
        });

        string[] sources = EnumerateRenderVisuals(rendered.Pages[0].Scene)
            .OfType<HtmlRenderShape>()
            .Where(shape => shape.Shape.FillColor != null)
            .Select(shape => shape.Source!)
            .ToArray();
        Assert.Contains("div#first-flow", sources);
        Assert.Contains("div#before::before", sources);
        Assert.Contains("div#positive", sources);
        Assert.Contains("div#relative-positive", sources);
        Assert.Contains("div#after::after", sources);
        Assert.Contains("div#second-flow", sources);
        Assert.Contains("div#relative-negative", sources);
        Assert.True(Array.IndexOf(sources, "div#first-flow") < Array.IndexOf(sources, "div#before::before"));
        Assert.True(Array.IndexOf(sources, "div#before::before") < Array.IndexOf(sources, "div#positive"));
        Assert.True(Array.IndexOf(sources, "div#positive") < Array.IndexOf(sources, "div#relative-positive"));
        Assert.True(Array.IndexOf(sources, "div#relative-negative") < Array.IndexOf(sources, "div#after::after"));
        Assert.True(Array.IndexOf(sources, "div#after::after") < Array.IndexOf(sources, "div#second-flow"));
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.PositioningModeUnsupported
            && (diagnostic.Source == "div#before::before" || diagnostic.Source == "div#after::after"));
    }

    [Fact]
    public void HtmlGeneratedContent_FloatAfterLinkTextKeepsGeneratedUrlOnAvailableLine() {
        string png = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(10, 10));
        string html = "<style>body,p{margin:0}p{width:440px;padding-left:120px}"
            + "img{float:left;width:100px;height:60px;margin-left:-120px}"
            + "a::after{content:' (example.test)'}</style>"
            + "<p><strong><a href='https://example.test'>Tables with one header"
            + "<img alt='' src='data:image/png;base64," + png + "'></a></strong> for rows and columns.</p>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 600D,
            Margins = HtmlRenderMargins.All(0D)
        });
        HtmlRenderVisual[] visuals = EnumerateRenderVisuals(rendered.Pages[0].Scene).ToArray();
        HtmlRenderText title = Assert.Single(visuals.OfType<HtmlRenderText>(),
            text => text.Text.Contains("Tables with one header", StringComparison.Ordinal));
        HtmlRenderText url = Assert.Single(visuals.OfType<HtmlRenderText>(),
            text => text.Source == "a::after" && text.Text.Contains("example.test", StringComparison.Ordinal));
        HtmlRenderImage image = Assert.Single(visuals.OfType<HtmlRenderImage>());

        Assert.Equal(title.Y, url.Y, 3);
        Assert.True(url.X > title.X);
        Assert.True(image.X < title.X);
    }

    [Fact]
    public void HtmlGeneratedContent_RatioWrapperPaintsPositionedImage() {
        string data = Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(10, 10));
        string html = "<style>body{margin:0}ul{display:flex;flex-wrap:wrap;list-style:none;margin:0;padding:0}"
            + "li{position:relative;overflow:hidden;width:320px;flex:none}a{width:100%}"
            + ".ratio{position:relative;width:100%;--bs-aspect-ratio:56.25%}"
            + ".ratio::before{content:\"\";display:block;padding-top:var(--bs-aspect-ratio)}"
            + ".ratio>img{position:absolute;top:0;left:0;width:100%;height:100%;object-fit:cover}</style>"
            + "<ul><li><div><a href='https://example.com/card'><div class='ratio'><img src='data:image/png;base64," + data
            + "' alt='gallery tile'></div></a></div></li></ul>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 400D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderImage image = Assert.Single(EnumerateRenderVisuals(rendered.Pages[0].Visuals).OfType<HtmlRenderImage>());
        Assert.Equal(320D, image.Width, 3);
        Assert.Equal(180D, image.Height, 3);
        Assert.Equal("https://example.com/card", image.LinkUri);
    }

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
    public void HtmlGeneratedContent_ContainerQueriesSeeTheOriginatingElementContainer() {
        const string html = """
            <style>
              .card { container-type:inline-size; width:160px; }
              .outer { container-type:inline-size; width:80px; }
              @container (width > 100px) { .card::before { content:"ContainerBefore"; color:red; } }
            </style>
            <section class="outer"><div class="card">Body</div></section>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html), new HtmlRenderOptions {
            ViewportWidth = 240D,
            Margins = HtmlRenderMargins.All(0D)
        });

        HtmlRenderText generated = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Source == "div.card::before");
        Assert.Equal("ContainerBefore", generated.Text);
        Assert.Equal(OfficeColor.Red, generated.Color);
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
        HtmlToPdfOptions pdfOptions = new HtmlToPdfOptions();
        string pdfText = string.Concat(PdfCore.PdfReadDocument.Open(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfBytes(pdfOptions)).ExtractText().Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Bytes.Take(8));
        Assert.Contains("Generated", svg, StringComparison.Ordinal);
        Assert.Contains("BackendMarker", svg, StringComparison.Ordinal);
        Assert.Contains("GeneratedBackendMarker", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(OfficeIMO.Html.HtmlConversionDocument.Parse(html).ToPdfDocumentResult(pdfOptions).Report.Warnings, warning => warning.Severity == PdfCore.PdfConversionWarningSeverity.Error);
    }

    [Fact]
    public void HtmlGeneratedContent_RendersMixedTextImagesAndQuotes() {
        const string html = """
            <style>
              .image::before { content:"[" url('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNgYAAAAAMAASsJTYQAAAAASUVORK5CYII=') "]"; }
              .quote::after { content:open-quote; }
              .flex::before { content:"FlexFallback"; display:flex; }
            </style>
            <p class="image">ImageFallback</p><p class="quote">QuoteFallback</p><p class="flex">FlexHost</p>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        HtmlRenderImage image = Assert.Single(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderImage>());
        Assert.Equal(1D, image.Width, 3);
        Assert.Equal(1D, image.Height, 3);
        Assert.Equal(new[] { "[", "]" }, rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>()
            .Where(text => text.Source == "p.image::before").Select(text => text.Text).ToArray());
        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Source == "p.quote::after" && text.Text == "\u201c");
        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Source == "p.flex::before" && text.Text == "FlexFallback");
        Assert.Single(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.FlexLayoutPending && diagnostic.Source == "p.flex::before");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported);
        Assert.True(HtmlDiagnosticCatalog.TryGet(HtmlRenderDiagnosticCodes.GeneratedContentUnsupported, out _));
    }

    [Fact]
    public void HtmlGeneratedContent_TracksNestedAuthoredQuotePairsAndNoQuoteDepthTokens() {
        const string html = """
            <style>
              body { quotes: "«" "»" "‹" "›"; }
              q::before { content:open-quote; }
              q::after { content:close-quote; }
              .silent::before { content:no-open-quote; }
              .silent::after { content:no-close-quote; }
            </style>
            <q>outer <q>inner</q></q><span class="silent">silent</span><q>again</q>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html);
        string pdfText = string.Concat(PdfCore.PdfReadDocument
            .Open(HtmlConversionDocument.Parse(html).ToPdfBytes(new HtmlToPdfOptions()))
            .ExtractText()
            .Where(character => !char.IsWhiteSpace(character)));

        Assert.Equal("«outer ‹inner›»silent«again»", rendered.Text.Replace("\r", string.Empty).Replace("\n", string.Empty));
        Assert.Contains("«outer‹inner›»silent«again»", pdfText, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported);
    }

    [Fact]
    public void HtmlGeneratedContent_DiagnosesUnsupportedCounterDeclarationsAndStyles() {
        const string html = """
            <style>
              .declaration { --bad-counter:chapter 1 2; counter-reset:var(--bad-counter); }
              .declaration::before { content:counter(chapter) " "; }
              .style { --bad-content:counter(chapter, symbols(additive "*")); }
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
    public void HtmlGeneratedContent_HonorsAuthoredOverridesOfPredefinedCounterStyles() {
        const string html = """
            <style>
              @counter-style decimal { system:cyclic; symbols:"X"; }
              body { counter-reset:item; }
              p::before { counter-increment:item; content:counter(item, decimal) " "; }
            </style>
            <p>Body</p>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));

        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(), text => text.Source == "p::before" && text.Text == "X ");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedCounterUnsupported);
    }

    [Fact]
    public void HtmlGeneratedContent_FormatsSymbolsFunctionsThroughTheSharedCounterOwner() {
        const string html = """
            <style>
              body { counter-reset:item 4; }
              .cyclic { --marker:counter(item, symbols(cyclic "①" "②" "③")) " "; }
              .numeric { --marker:counter(item, symbols(numeric "0" "1")) " "; }
              .alphabetic { --marker:counter(item, symbols(alphabetic "A" "B")) " "; }
              .symbolic { --marker:counter(item, symbols("*" "†")) " "; }
              p::before { counter-increment:item; content:var(--marker); }
            </style>
            <p class="cyclic">Cyclic</p><p class="numeric">Numeric</p>
            <p class="alphabetic">Alphabetic</p><p class="symbolic">Symbolic</p>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));
        string text = rendered.Text;

        Assert.Contains("② ", text, StringComparison.Ordinal);
        Assert.Contains("110 ", text, StringComparison.Ordinal);
        Assert.Contains("AAA ", text, StringComparison.Ordinal);
        Assert.Contains("†††† ", text, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported
            || diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedCounterUnsupported);
    }

    [Fact]
    public void HtmlGeneratedContent_FormatsNamedAdditiveStylesAndRangeFallbacks() {
        const string html = """
            <style>
              @counter-style tally {
                system:additive;
                additive-symbols:5 "V", 1 "I";
                range:1 9;
                fallback:decimal;
              }
              body { counter-reset:item 6 overflow 10; }
              .tally::before { counter-increment:item; content:counter(item, tally) " "; }
              .fallback::before { counter-increment:overflow; content:counter(overflow, tally) " "; }
            </style>
            <p class="tally">Tally</p><p class="fallback">Fallback</p>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(HtmlConversionDocument.Parse(html));

        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(),
            text => text.Source == "p.tally::before" && text.Text == "VII ");
        Assert.Contains(rendered.Pages.SelectMany(page => page.Visuals).OfType<HtmlRenderText>(),
            text => text.Source == "p.fallback::before" && text.Text == "11 ");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported);
    }

    [Fact]
    public void HtmlGeneratedContent_ResolvesTargetTextListCountersAndLeaders() {
        const string html = """
            <style>
              ol { margin:0; padding:0 0 0 24px; }
              .xref::before { content:target-text(attr(href)) " " target-counter(attr(href), list-item, upper-roman) leader(solid); }
            </style>
            <ol><li id="first">Referenced item</li><li>Other item</li></ol>
            <a class="xref" href="#first">Index</a>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            ViewportWidth = 260D,
            Margins = HtmlRenderMargins.All(0D)
        });
        IReadOnlyList<HtmlRenderText> generated = rendered.Pages[0].Visuals
            .OfType<HtmlRenderText>()
            .Where(text => text.Source == "a.xref::before")
            .ToList();
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported);
        HtmlRenderShape leader = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape =>
            shape.Source != null && shape.Source.StartsWith("a.xref::before:content-leader", StringComparison.Ordinal));

        Assert.Contains(generated, text => text.Text == "Referenced item I");
        Assert.Equal(OfficeStrokeDashStyle.Solid, leader.Shape.StrokeDashStyle);
        Assert.True(leader.Width > 40D);
        Assert.DoesNotContain("_", rendered.Text, StringComparison.Ordinal);
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported);
    }

    [Fact]
    public void HtmlGeneratedContent_ResolvesTargetPagesAfterBoundedPagedReflow() {
        const string html = """
            <style>
              @page { size:240px 90px; margin:10px; }
              body, p, h1 { margin:0; }
              .toc::before { content:target-text(url(#chapter)) leader(dotted) target-counter(url(#chapter), page, upper-roman); }
              h1 { break-before:page; font-size:14px; line-height:18px; }
            </style>
            <p class="toc">Index</p><h1 id="chapter">Chapter One</h1>
            """;

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html, new HtmlRenderOptions {
            Mode = HtmlRenderMode.Paged,
            HonorCssPageRules = true
        });
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported);
        HtmlRenderText pageCounter = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text =>
            text.Source != null && text.Source.StartsWith("p.toc::before:content-targetpage", StringComparison.Ordinal));
        HtmlRenderShape leader = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(), shape =>
            shape.Source != null && shape.Source.StartsWith("p.toc::before:content-leader", StringComparison.Ordinal));

        Assert.True(rendered.Pages.Count >= 2);
        HtmlRenderPage targetPage = Assert.Single(rendered.Pages, page =>
            page.PageNumber > 1 && page.Visuals.OfType<HtmlRenderText>().Any(text => text.Text == "Chapter One"));
        Assert.Equal(targetPage.PageNumber == 2 ? "II" : targetPage.PageNumber == 3 ? "III" : targetPage.PageNumber.ToString(), pageCounter.Text);
        Assert.Equal(OfficeStrokeDashStyle.Dot, leader.Shape.StrokeDashStyle);
        Assert.Contains(rendered.Pages[0].Visuals.OfType<HtmlRenderText>(), text => text.Text == "Chapter One" && text.Source == "p.toc::before");
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.GeneratedContentUnsupported);
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
