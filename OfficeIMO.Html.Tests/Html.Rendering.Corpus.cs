using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void StaticRendererAppliesDefaultLinkPaintWithoutOverridingAuthorStyles() {
        const string html = "<div style='color:#444'>Notice <a href='https://example.test'>Default link</a>"
            + "<a href='https://example.test' style='color:inherit;text-decoration:none'>Plain link</a>"
            + "<a href='https://example.test' style='text-decoration:none;text-decoration-line:underline'>Restored underline</a>"
            + "<a href='https://example.test' style='text-decoration-line:underline;text-decoration:none'>Removed underline</a>"
            + "<a href='https://example.test' style='text-decoration:underline dotted red'>Decorated link</a>"
            + "<a href='https://example.test'><span style='text-decoration:none'>Nested link</span></a>"
            + "<a href='https://example.test'><span style='color:red'>Nested red link</span></a>"
            + "<a href='https://example.test'><span style='display:inline-block'>Atomic link</span></a>"
            + "<a href='https://example.test' style='color:revert;text-decoration:revert'>Reverted link</a>"
            + "<a href='https://example.test' style='color:revert-layer;text-decoration:revert-layer'>Layer reverted link</a>"
            + "<a>Anchor without href</a></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { ViewportWidth = 1600D });
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateCorpusVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();

        HtmlRenderText defaultLink = Assert.Single(text, item => item.Text == "Default link");
        Assert.Equal(OfficeColor.FromRgb(0, 0, 238), defaultLink.Color);
        Assert.Equal(OfficeTextDecorationStyle.Single, defaultLink.UnderlineStyle);

        HtmlRenderText plainLink = Assert.Single(text, item => item.Text == "Plain link");
        Assert.Equal(OfficeColor.FromRgb(0x44, 0x44, 0x44), plainLink.Color);
        Assert.Equal(OfficeTextDecorationStyle.None, plainLink.UnderlineStyle);

        Assert.Equal(OfficeTextDecorationStyle.Single,
            Assert.Single(text, item => item.Text == "Restored underline").UnderlineStyle);
        Assert.Equal(OfficeTextDecorationStyle.None,
            Assert.Single(text, item => item.Text == "Removed underline").UnderlineStyle);

        HtmlRenderText decoratedLink = Assert.Single(text, item => item.Text == "Decorated link");
        Assert.Equal(OfficeTextDecorationStyle.Dotted, decoratedLink.UnderlineStyle);
        Assert.Equal(OfficeColor.Red, decoratedLink.DecorationColor);

        HtmlRenderText nestedLink = Assert.Single(text, item => item.Text == "Nested link");
        Assert.Equal(OfficeColor.FromRgb(0, 0, 238), nestedLink.Color);
        Assert.Equal(OfficeTextDecorationStyle.Single, nestedLink.UnderlineStyle);
        HtmlRenderText redLink = Assert.Single(text, item => item.Text == "Nested red link");
        Assert.Equal(OfficeColor.Red, redLink.Color);
        Assert.Equal(OfficeTextDecorationStyle.Single, redLink.UnderlineStyle);
        HtmlRenderText atomicLink = Assert.Single(text, item => item.Text == "Atomic link");
        Assert.Equal(OfficeColor.FromRgb(0, 0, 238), atomicLink.Color);
        Assert.Equal(OfficeTextDecorationStyle.None, atomicLink.UnderlineStyle);
        foreach (string label in new[] { "Reverted link", "Layer reverted link" }) {
            HtmlRenderText revertedLink = Assert.Single(text, item => item.Text == label);
            Assert.Equal(OfficeColor.FromRgb(0, 0, 238), revertedLink.Color);
            Assert.Equal(OfficeTextDecorationStyle.Single, revertedLink.UnderlineStyle);
        }

        HtmlRenderText anchor = Assert.Single(text, item => item.Text == "Anchor without href");
        Assert.Equal(OfficeColor.FromRgb(0x44, 0x44, 0x44), anchor.Color);
        Assert.Equal(OfficeTextDecorationStyle.None, anchor.UnderlineStyle);
    }

    [Theory]
    [InlineData("1.5em", 30D)]
    [InlineData("150%", 30D)]
    [InlineData("1.5", 150D)]
    public void StaticRendererInheritsComputedLineHeightAcrossFontSizeChanges(string lineHeight, double expected) {
        string html = "<div style='font-size:20px;line-height:" + lineHeight
            + "'><h1 style='font-size:100px;margin:0'>Heading</h1></div>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { ViewportWidth = 640D });
        HtmlRenderText heading = Assert.Single(rendered.Pages
            .SelectMany(page => EnumerateCorpusVisuals(page.Scene))
            .OfType<HtmlRenderText>(), item => item.Text == "Heading");

        Assert.Equal(expected, heading.LineHeight, 3);
    }

    [Fact]
    public void StaticRendererCentersOversizedGlyphPaintAroundInheritedShortLineBox() {
        static HtmlRenderText RenderHeading(string inheritedLineHeight) {
            string html = "<div style='margin-top:100px;font-size:20px;line-height:"
                + inheritedLineHeight + "'><h1 style='font-size:100px;margin:0'>Heading</h1></div>";
            HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
                new HtmlRenderOptions { ViewportWidth = 640D });
            return Assert.Single(rendered.Pages.SelectMany(page => EnumerateCorpusVisuals(page.Scene))
                .OfType<HtmlRenderText>(), item => item.Text == "Heading");
        }

        HtmlRenderText shortLine = RenderHeading("20px");
        HtmlRenderText fullLine = RenderHeading("100px");

        Assert.Equal(20D, shortLine.LineHeight, 3);
        Assert.Equal(100D, shortLine.Height, 3);
        Assert.Equal(fullLine.Y - 40D, shortLine.Y, 3);
    }

    [Fact]
    public void StaticRendererUsesSelectedFaceMetricsForOversizedTextInShortLineBoxes() {
        const string html = "<div style='font:20px/20px Example'><h1 style='font-size:100px;margin:0'>Heading</h1></div>";
        HtmlRenderText RenderHeading(HtmlRenderOptions options) => Assert.Single(
            HtmlRenderTestDriver.Render(html, options).Pages
                .SelectMany(page => EnumerateCorpusVisuals(page.Scene))
                .OfType<HtmlRenderText>(), item => item.Text == "Heading");

        HtmlRenderText fallback = RenderHeading(new HtmlRenderOptions { ViewportWidth = 640D });
        HtmlRenderText selectedFace = RenderHeading(new HtmlRenderOptions {
            ViewportWidth = 640D,
            FallbackTextFaceMetrics = (_, _, _) => new HtmlTextFaceMetrics(114D, 93D)
        });

        Assert.Equal(20D, selectedFace.LineHeight, 3);
        Assert.Equal(fallback.LayoutY, selectedFace.LayoutY, 3);
        Assert.Equal(fallback.Y - 14D, selectedFace.Y, 3);
        Assert.Equal(114D, selectedFace.Height, 3);
    }

    [Fact]
    public void StaticRendererKeepsOversizedGlyphsWhenShortLineBoxesCrossPages() {
        const string html = "<style>@page{size:300px 60px;margin:0}html,body,p{margin:0}"
            + "p{font:100px/20px Arial}</style><p>I<br>J<br>K<br>L</p>";

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged });
        HtmlRenderText[] glyphs = rendered.Pages
            .SelectMany(page => EnumerateCorpusVisuals(page.Scene))
            .OfType<HtmlRenderText>()
            .Where(item => item.Text is "I" or "J" or "K" or "L")
            .ToArray();

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Equal(new[] { "I", "J", "K", "L" }, glyphs.Select(item => item.Text).ToArray());
        Assert.DoesNotContain(rendered.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.VisualFragmentUnsupported);
    }

    [Theory]
    [InlineData("margin:0 auto", 30D)]
    [InlineData("margin-left:auto;margin-right:0", 60D)]
    [InlineData("margin-left:0;margin-right:auto", 0D)]
    public void StaticRendererResolvesNormalFlowBlockAutoMargins(string margins, double expectedX) {
        string html = "<body style='margin:0'><div id='box' style='width:240px;height:20px;background:red;" + margins + "'></div></body>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { ViewportWidth = 300D, Margins = HtmlRenderMargins.All(0D) });

        HtmlRenderShape box = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderShape>(),
            shape => shape.Source == "div#box" && shape.Shape.FillColor == OfficeColor.Red);
        Assert.Equal(expectedX, box.X, 3);
    }

    [Theory]
    [InlineData(HtmlRenderUserAgentStyleMode.Document)]
    [InlineData(HtmlRenderUserAgentStyleMode.Browser)]
    public void PagedRendererKeepsCenteredBodyMaxWidthAcrossPages(HtmlRenderUserAgentStyleMode userAgentStyles) {
        const string html = "<style>@page{size:300px 100px;margin:0}body{max-width:200px;margin:0 auto}"
            + "div{height:90px;background:red;break-after:page}</style>"
            + "<body><div id='first'></div><div id='second'></div></body>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, UserAgentStyles = userAgentStyles });

        Assert.Equal(2, rendered.Pages.Count);
        foreach (HtmlRenderPage page in rendered.Pages) {
            HtmlRenderShape box = Assert.Single(page.Visuals.OfType<HtmlRenderShape>(),
                shape => shape.Shape.FillColor == OfficeColor.Red);
            Assert.Equal(50D, box.X, 3);
            Assert.Equal(200D, box.Shape.Width, 3);
        }
    }

    [Theory]
    [InlineData(HtmlRenderUserAgentStyleMode.Document)]
    [InlineData(HtmlRenderUserAgentStyleMode.Browser)]
    public void PagedRendererRecentersBodyWhenLaterPageWidthChanges(HtmlRenderUserAgentStyleMode userAgentStyles) {
        const string html = "<style>@page{size:500px 100px;margin:0}@page:first{size:300px 100px;margin:0}"
            + "body{max-width:200px;margin:0 auto}div{height:90px;background:red}"
            + "#first{break-after:page}</style>"
            + "<body><div id='first'></div><div id='second'></div></body>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { Mode = HtmlRenderMode.Paged, UserAgentStyles = userAgentStyles });

        Assert.Equal(2, rendered.Pages.Count);
        Assert.Equal(new[] { 300D, 500D }, rendered.Pages.Select(page => page.Width).ToArray());
        for (int index = 0; index < 2; index++) {
            HtmlRenderShape box = Assert.Single(rendered.Pages[index].Visuals.OfType<HtmlRenderShape>(),
                shape => shape.Shape.FillColor == OfficeColor.Red);
            Assert.Equal(index == 0 ? 50D : 150D, box.X, 3);
            Assert.Equal(200D, box.Shape.Width, 3);
        }
        Assert.DoesNotContain(rendered.Diagnostics,
            diagnostic => diagnostic.Code == HtmlRenderDiagnosticCodes.PagePseudoGeometryPending);
    }

    [Fact]
    public void StaticRendererCentersIntrinsicBlockImageWithAutoMargins() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(40, 20));
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<body style='margin:0'><img src='" + image + "' style='display:block;margin:0 auto'></body>",
            new HtmlRenderOptions { ViewportWidth = 100D, Margins = HtmlRenderMargins.All(0D) });

        HtmlRenderImage visual = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>());
        Assert.Equal(30D, visual.X, 3);
    }

    [Fact]
    public void StaticRendererUsesSelectedPictureSourceWhenCenteringUndecodableImage() {
        string fallback = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(40, 20));
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<body style='margin:0'><picture style='display:block'><source srcset='data:image/png;base64,AQID 1x'>"
            + "<img id='picture' src='" + fallback + "' style='display:block;margin:0 auto'></picture></body>",
            new HtmlRenderOptions { ViewportWidth = 400D, Margins = HtmlRenderMargins.All(0D) });

        HtmlRenderImage image = Assert.Single(rendered.Pages[0].Visuals.OfType<HtmlRenderImage>(),
            visual => visual.Source == "img#picture");
        Assert.Equal(50D, image.X, 3);
        Assert.Equal(300D, image.Width, 3);
    }

    public static IEnumerable<object[]> HtmlRenderingRepresentativeCorpusScenarioIds => HtmlRenderingRepresentativeCorpus.All
        .Select(item => new object[] { item.Id });

    [Fact]
    public void AdvancedHeldOutRenderingCorpus_LoadsOnlyFrozenInputsAndDeclaredContracts() {
        HtmlRenderingAdvancedHeldOutCorpus corpus = HtmlRenderingAdvancedHeldOutCorpus.Load();

        Assert.Equal("officeimo-html-h4-advanced-held-out", corpus.Manifest.CorpusId);
        Assert.Equal(8, corpus.Cases.Count);
        Assert.Equal(
            new[] { "screen-full-page-v1", "print-paged-v1", "screen-snapshot-paged-v1" },
            corpus.Manifest.RenderIntents);
        Assert.Equal(new[] { "pdf", "svg", "raster" }, corpus.Manifest.OutputFamilies);
        Assert.Equal(64, corpus.ManifestSha256.Length);
        Assert.All(corpus.Cases, item => {
            Assert.NotEmpty(item.Html);
            Assert.Equal(item.Manifest.Length, item.SourceBytes.LongLength);
            Assert.All(item.Manifest.TextMarkers, marker => Assert.Contains(marker, item.Html, StringComparison.Ordinal));
        });
        HtmlRenderingAdvancedHeldOutCase legacy = corpus.Cases.Single(item => item.Id == "legacy-portal");
        Assert.Equal(1483, legacy.SourceBytes.Length);
        Assert.Contains((byte)0x97, legacy.SourceBytes);
        Assert.Contains(legacy.Html, value => value == '—');
    }

    [Fact]
    public void StaticRendererAppliesBrowserHeadingAndEmphasisDefaultsWhenFontSizeIsImplicit() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<h1>Default heading</h1><p>Body <strong>strong text</strong></p>",
            new HtmlRenderOptions { ViewportWidth = 640D });
        HtmlRenderText[] text = rendered.Pages
            .SelectMany(page => EnumerateCorpusVisuals(page.Scene))
            .OfType<HtmlRenderText>()
            .ToArray();

        HtmlRenderText heading = Assert.Single(text, item => item.Text == "Default heading");
        HtmlRenderText body = Assert.Single(text, item => item.Text.Contains("Body", StringComparison.Ordinal));
        HtmlRenderText strong = Assert.Single(text, item => item.Text.Contains("strong text", StringComparison.Ordinal));

        Assert.Equal(body.Font.Size * 2D, heading.Font.Size, 3);
        Assert.True((heading.Font.Style & OfficeFontStyle.Bold) != 0);
        Assert.True((strong.Font.Style & OfficeFontStyle.Bold) != 0);
    }

    [Fact]
    public void StaticRendererAppliesBoldElementDefaultsAfterImplicitNormalInheritance() {
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(
            "<body style='font-weight:normal'><h1>Default heading</h1><h2 style='font-weight:normal'>Plain heading</h2>"
            + "<p><strong>Default emphasis</strong><b>Default bold</b><strong style='font-weight:normal'>Plain emphasis</strong></p></body>",
            new HtmlRenderOptions { ViewportWidth = 640D });
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateCorpusVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();

        foreach (string value in new[] { "Default heading", "Default emphasis", "Default bold" }) {
            HtmlRenderText visual = Assert.Single(text, item => item.Text == value);
            Assert.True((visual.Font.Style & OfficeFontStyle.Bold) != 0);
        }
        foreach (string value in new[] { "Plain heading", "Plain emphasis" }) {
            HtmlRenderText visual = Assert.Single(text, item => item.Text == value);
            Assert.True((visual.Font.Style & OfficeFontStyle.Bold) == 0);
        }
    }

    [Fact]
    public void PositionedTextRetainsRequestedNumericFaceAcrossInlineAndControlText() {
        const string html = "<style>body{font-weight:300}h1{font-weight:100}"
            + "input::placeholder{font-style:italic}</style>"
            + "<h1>Light heading</h1><p>Light body</p>"
            + "<input placeholder='Light prompt'>";
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(html,
            new HtmlRenderOptions { ViewportWidth = 640D });
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateCorpusVisuals(page.Scene))
            .OfType<HtmlRenderText>().ToArray();

        Assert.Equal(100, Assert.Single(text, item => item.Text == "Light heading").FontDescriptor.Weight);
        Assert.Equal(300, Assert.Single(text, item => item.Text == "Light body").FontDescriptor.Weight);
        HtmlRenderText prompt = Assert.Single(text, item => item.Text == "Light prompt");
        Assert.Equal(300, prompt.FontDescriptor.Weight);
        Assert.Equal(OfficeFontSlant.Italic, prompt.FontDescriptor.Slant);
    }

    [Fact]
    public void HtmlRenderingRepresentativeCorpus_CoversEveryPublishedMarketScenario() {
        Assert.Equal(
            HtmlMarketScenarioCatalog.All.Select(item => item.Id),
            HtmlRenderingRepresentativeCorpus.All.Select(item => item.Id));
    }

    [Fact]
    public void HtmlRenderingRepresentativeCorpus_DashboardHeadingAndIncidentRemainFullyVisible() {
        HtmlRenderingCorpusCase scenario = HtmlRenderingRepresentativeCorpus.All.Single(item => item.Id == "dashboard-print");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(scenario.Html, scenario.CreateOptions());
        HtmlRenderText[] text = rendered.Pages.SelectMany(page => EnumerateCorpusVisuals(page.Scene))
            .OfType<HtmlRenderText>()
            .ToArray();
        HtmlRenderText[] heading = text
            .Where(fragment => fragment.Text.Contains("Documents", StringComparison.Ordinal)
                || fragment.Text.Contains("processed", StringComparison.Ordinal))
            .ToArray();
        HtmlRenderText[] incident = text
            .Where(fragment => fragment.Text.Contains("Open incident", StringComparison.Ordinal)
                || fragment.Text.Contains("remapping", StringComparison.Ordinal))
            .ToArray();

        Assert.NotEmpty(heading);
        Assert.NotEmpty(incident);
        Assert.Single(heading.Select(fragment => Math.Round(fragment.Y, 3)).Distinct());
        Assert.Single(incident.Select(fragment => Math.Round(fragment.Y, 3)).Distinct());
    }

    [Fact]
    public void HtmlRenderingRepresentativeCorpus_StaticStandardsGridUsesTwoAuthoredColumns() {
        HtmlRenderingCorpusCase scenario = HtmlRenderingRepresentativeCorpus.All.Single(item => item.Id == "static-standards-showcase");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(scenario.Html, scenario.CreateOptions());
        HtmlRenderText firstRow = Assert.Single(
            rendered.Pages.SelectMany(page => EnumerateCorpusVisuals(page.Scene)).OfType<HtmlRenderText>(),
            text => text.Text == "Inherited row A");
        HtmlRenderText badge = Assert.Single(
            rendered.Pages.SelectMany(page => EnumerateCorpusVisuals(page.Scene)).OfType<HtmlRenderText>(),
            text => text.Text == "Clipped vector badge");
        HtmlRenderText evidence = Assert.Single(
            rendered.Pages.SelectMany(page => EnumerateCorpusVisuals(page.Scene)).OfType<HtmlRenderText>(),
            text => text.Text == "Named page evidence");

        Assert.True(firstRow.X < badge.X);
        Assert.InRange(Math.Abs(badge.X - evidence.X), 0D, 2D);
    }

    [Fact]
    public void HtmlRenderingRepresentativeCorpus_StaticStandardsRunningHeaderPaintsOnEveryRasterPage() {
        HtmlRenderingCorpusCase scenario = HtmlRenderingRepresentativeCorpus.All.Single(item => item.Id == "static-standards-showcase");
        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(scenario.Html, scenario.CreateOptions());

        Assert.Equal(2, rendered.Pages.Count);
        foreach (HtmlRenderPage page in rendered.Pages) {
            OfficeDrawing drawing = page.CreateDrawing();
            OfficeDrawingText header = Assert.Single(
                drawing.Elements.OfType<OfficeDrawingText>(),
                text => text.Text.Contains("Managed static standards", StringComparison.Ordinal));
            OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(drawing, 1D, OfficeColor.White);
            int coloredPixels = 0;
            int left = Math.Max(0, (int)Math.Floor(header.X - 5D));
            int top = Math.Max(0, (int)Math.Floor(header.Y - 5D));
            int right = Math.Min(image.Width - 1, (int)Math.Ceiling(header.X + header.Width + 5D));
            int bottom = Math.Min(image.Height - 1, (int)Math.Ceiling(header.Y + header.Height + 5D));
            for (int y = top; y <= bottom; y++) {
                for (int x = left; x <= right; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.B > pixel.R + 20 && pixel.B > pixel.G + 5) coloredPixels++;
                }
            }

            Assert.True(coloredPixels > 20, $"Page {page.PageNumber} running header produced only {coloredPixels} blue raster pixels.");
        }
    }

    [Theory]
    [MemberData(nameof(HtmlRenderingRepresentativeCorpusScenarioIds))]
    public void HtmlRenderingRepresentativeCorpus_ProvesSharedSceneImageAndSearchablePdf(string scenarioId) {
        HtmlRenderingCorpusCase scenario = HtmlRenderingRepresentativeCorpus.All.Single(item => item.Id == scenarioId);
        HtmlRenderOptions options = scenario.CreateOptions();

        HtmlRenderDocument rendered = HtmlRenderTestDriver.Render(scenario.Html, options);

        Assert.Equal(scenario.Mode, rendered.Mode);
        Assert.Equal(scenario.ExpectedPageCount, rendered.Pages.Count);
        Assert.All(rendered.Pages, page => {
            Assert.Equal(scenario.ExpectedSurfaceWidth, page.Width, 3);
            Assert.True(page.Height > 0D);
            Assert.True(
                page.Visuals.Count >= scenario.MinimumVisualCount,
                scenario.Id + " page " + page.PageNumber + " produced " + page.Visuals.Count + " visuals; expected at least " + scenario.MinimumVisualCount + ".");
        });
        Assert.True(rendered.Headings.Count >= scenario.MinimumHeadingCount);
        string logicalText = NormalizeCorpusWhitespace(rendered.Text);
        foreach (string marker in scenario.TextMarkers) Assert.Contains(NormalizeCorpusWhitespace(marker), logicalText, StringComparison.Ordinal);
        foreach (string code in scenario.DiagnosticCodes) {
            Assert.Contains(rendered.Diagnostics, diagnostic => diagnostic.Code == code);
        }
        foreach (string code in scenario.ForbiddenDiagnosticCodes) {
            Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Code == code);
        }
        Assert.DoesNotContain(rendered.Diagnostics, diagnostic => diagnostic.Severity == HtmlDiagnosticSeverity.Error);
        HtmlRenderVisual[] visuals = rendered.Pages.SelectMany(page => EnumerateCorpusVisuals(page.Scene)).ToArray();
        foreach (string source in scenario.RequiredVisualSources) {
            Assert.Contains(visuals, visual => string.Equals(visual.Source, source, StringComparison.Ordinal));
        }
        if (scenario.LinkUri != null) {
            Assert.Contains(visuals, visual => visual.LinkUri == scenario.LinkUri);
        }

        OfficeDrawing firstPage = rendered.Pages[0].CreateDrawing();
        byte[] png = OfficeDrawingRasterRenderer.ToPng(firstPage, 0.5D, OfficeColor.White);
        string svg = OfficeDrawingSvgExporter.ToSvg(firstPage, 0.5D);
        Assert.True(png.Length > 100);
        Assert.Equal(new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }, png.Take(8).ToArray());
        Assert.Contains("<svg", svg, StringComparison.Ordinal);
        string svgText = System.Net.WebUtility.HtmlDecode(
            System.Text.RegularExpressions.Regex.Replace(svg, "<[^>]+>", string.Empty));
        foreach (string word in NormalizeCorpusWhitespace(scenario.TextMarkers[0]).Split(' ')) {
            Assert.True(
                svg.Contains(word, StringComparison.Ordinal) || svgText.Contains(word, StringComparison.Ordinal),
                $"Expected SVG paint text to retain '{word}' either in one text node or across positioned grapheme nodes.");
        }

        HtmlToPdfOptions pdfOptions = new HtmlToPdfOptions();
        pdfOptions = new HtmlToPdfOptions(options);
        byte[] pdf = OfficeIMO.Html.HtmlConversionDocument.Parse(scenario.Html).ToPdfBytes(pdfOptions);
        PdfCore.PdfDocumentInfo pdfInfo = PdfCore.PdfInspector.Inspect(pdf);
        string pdfText = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Equal(scenario.ExpectedPageCount, pdfInfo.PageCount);
        string normalizedPdfText = NormalizeCorpusWhitespace(pdfText);
        foreach (string marker in scenario.TextMarkers) {
            foreach (string searchableToken in NormalizeCorpusWhitespace(marker).Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                Assert.Contains(searchableToken, normalizedPdfText, StringComparison.Ordinal);
            }
        }
        if (scenario.LinkUri != null) Assert.Contains(scenario.LinkUri, pdfInfo.LinkUris);

        WriteStaticStandardsReviewArtifacts(scenario, rendered, pdf);
    }

    private static void WriteStaticStandardsReviewArtifacts(
        HtmlRenderingCorpusCase scenario,
        HtmlRenderDocument rendered,
        byte[] pdf) {
        string? directory = Environment.GetEnvironmentVariable("OFFICEIMO_HTML_STANDARDS_ARTIFACT_DIR");
        if (scenario.Id != "static-standards-showcase" || string.IsNullOrWhiteSpace(directory)) return;

        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, "static-standards.pdf"), pdf);
        for (int pageIndex = 0; pageIndex < rendered.Pages.Count; pageIndex++) {
            OfficeDrawing pageDrawing = rendered.Pages[pageIndex].CreateDrawing();
            byte[] pagePng = OfficeDrawingRasterRenderer.ToPng(
                pageDrawing,
                1D,
                OfficeColor.White);
            File.WriteAllBytes(
                Path.Combine(directory, "static-standards-page-" + (pageIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + ".png"),
                pagePng);
            File.WriteAllText(
                Path.Combine(directory, "static-standards-page-" + (pageIndex + 1).ToString(System.Globalization.CultureInfo.InvariantCulture) + ".svg"),
                OfficeDrawingSvgExporter.ToSvg(pageDrawing, 1D));
        }
    }

    private static string NormalizeCorpusWhitespace(string value) {
        var result = new System.Text.StringBuilder(value.Length);
        bool pendingSpace = false;
        foreach (char character in value) {
            if (char.IsWhiteSpace(character)) {
                pendingSpace = result.Length > 0;
                continue;
            }
            if (pendingSpace) result.Append(' ');
            result.Append(character);
            pendingSpace = false;
        }
        return result.ToString();
    }

    private static IEnumerable<HtmlRenderVisual> EnumerateCorpusVisuals(IEnumerable<HtmlRenderVisual> visuals) {
        foreach (HtmlRenderVisual visual in visuals) {
            yield return visual;
            IEnumerable<HtmlRenderVisual>? children = visual switch {
                HtmlRenderClipGroup clip => clip.Visuals,
                HtmlRenderPathClipGroup pathClip => pathClip.Visuals,
                HtmlRenderEffectGroup effect => effect.Visuals,
                HtmlRenderSemanticGroup semantic => semantic.Visuals,
                HtmlRenderLogicalTextGroup logical => logical.Visuals,
                _ => null
            };
            if (children == null) continue;
            foreach (HtmlRenderVisual child in EnumerateCorpusVisuals(children)) yield return child;
        }
    }
}
