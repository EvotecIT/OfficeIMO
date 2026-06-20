using AngleSharp.Dom;
using OfficeIMO.Html;
using OfficeIMO.Word.Html;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void HtmlEnginePlatform_ConnectsProfilesIrStylesResourcesScoringDiagnosticsAndGallery() {
        const string sourceHtml = """
            <!doctype html>
            <html lang="en">
            <head>
                <base href="https://example.test/reports/">
                <title>Market Report</title>
                <link rel="stylesheet" href="/assets/report.css">
                <link rel="preload" as="script" href="javascript:alert(2)">
                <link rel="preload" as="image" href="/images/preload.png">
                <style>
                    @import "file:///secret/theme.css";
                    @import url('file:///secret/print.css');
                    @font-face { font-family: Demo; src: url('file:///secret/font.woff2'); }
                    body { color: #222; font-family: Aptos; }
                    article.report { color: #123456; }
                    .warning { color: #aa0000; }
                    p { color: #0000aa; }
                    p.force { color: #0000aa !important; }
                    a[href*=','] { font-weight: 700; }
                    .hero { background-image: url('/images/background.png'); }
                    table.financials td { color: #654321; border: 1px solid #444; padding: 4pt; }
                </style>
            </head>
            <body>
                <article class="report" dir="ltr">
                    <h1 style="font-weight: 700">Market Report</h1>
                    <p>Prepared for OfficeIMO HTML engine validation.</p>
                    <p class="warning">Specificity should keep this warning red.</p>
                    <p class="warning force">Important rules should still win.</p>
                    <a href="https://example.test/a,b">comma link</a>
                    <a href="javascript:alert(1)">unsafe link</a>
                    <picture>
                        <source srcset="/images/chart-large.png 2x, file:///secret/chart.png 3x">
                        <source data-srcset="/images/chart-lazy.png 2x">
                        <img src="/images/chart.png" alt="Revenue chart">
                    </picture>
                    <video poster="/media/poster.png">
                        <source src="/media/demo.mp4" type="video/mp4">
                        <source data-src="/media/lazy-demo.mp4" type="video/mp4">
                    </video>
                    <div class="hero"></div>
                    <object data="file:///secret/report.pdf"></object>
                    <script src="javascript:alert(1)"></script>
                    <table class="financials">
                        <tr><th>Metric</th><th>Value</th></tr>
                        <tr><td>Revenue</td><td>$42,000</td></tr>
                    </table>
                    <label>Approved <input type="checkbox" checked></label>
                </article>
            </body>
            </html>
            """;

        const string roundTripHtml = """
            <html>
            <body>
                <article>
                    <h1>Market Report</h1>
                    <p>Prepared for OfficeIMO HTML engine validation.</p>
                    <img src="https://example.test/reports/images/chart.png" alt="Revenue chart">
                    <table><tr><th>Metric</th><th>Value</th></tr><tr><td>Revenue</td><td>$42,000</td></tr></table>
                    <input type="checkbox" checked>
                </article>
            </body>
            </html>
            """;

        HtmlLogicalDocument logical = HtmlLogicalDocumentBuilder.FromHtml(sourceHtml);
        Assert.Contains("tables", logical.Capabilities);
        Assert.Contains("images", logical.Capabilities);
        Assert.Contains("forms", logical.Capabilities);
        Assert.True(logical.Count(HtmlLogicalNodeKind.Table) >= 1);
        Assert.True(logical.Count(HtmlLogicalNodeKind.FormControl) >= 1);

        var parsed = HtmlDocumentParser.ParseDocument(sourceHtml);
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(parsed);
        IElement heading = parsed.QuerySelector("h1")!;
        IElement tableCell = parsed.QuerySelector("td")!;
        IElement warning = parsed.QuerySelector("p.warning")!;
        IElement force = parsed.QuerySelector("p.force")!;
        IElement commaLink = parsed.QuerySelector("a[href*=',']")!;
        Assert.Equal("rgba(18, 52, 86, 1)", styles[heading].GetValue("color"));
        Assert.Equal("700", styles[heading].GetValue("font-weight"));
        Assert.Equal("rgba(101, 67, 33, 1)", styles[tableCell].GetValue("color"));
        Assert.Equal("rgba(170, 0, 0, 1)", styles[warning].GetValue("color"));
        Assert.Equal("rgba(0, 0, 170, 1)", styles[force].GetValue("color"));
        Assert.Equal("700", styles[commaLink].GetValue("font-weight"));

        var resourceManifest = HtmlResourcePipeline.BuildManifest(parsed, new HtmlResourcePipelineOptions {
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
        });
        Assert.True(resourceManifest.AllowedCount >= 3);
        Assert.True(resourceManifest.BlockedCount >= 2);
        Assert.Contains(resourceManifest.Resources, resource => resource.Kind == HtmlResourceKind.Stylesheet && resource.IsAllowed);
        Assert.Contains(resourceManifest.Resources, resource => resource.Kind == HtmlResourceKind.Script && resource.AttributeName == "href" && resource.DiagnosticCode == "ScriptResourceRejectedByPolicy");
        Assert.Contains(resourceManifest.Resources, resource => resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "href" && resource.IsAllowed);
        Assert.Contains(resourceManifest.Resources, resource => resource.ElementName == "video" && resource.AttributeName == "poster" && resource.Kind == HtmlResourceKind.Image && resource.IsAllowed);
        Assert.Contains(resourceManifest.Resources, resource => resource.ElementName == "source" && resource.Kind == HtmlResourceKind.Media && resource.IsAllowed);
        Assert.Contains(resourceManifest.Resources, resource => resource.ElementName == "source" && resource.AttributeName == "data-src" && resource.Kind == HtmlResourceKind.Media && resource.IsAllowed);
        Assert.Contains(resourceManifest.Resources, resource => resource.ElementName == "source" && resource.AttributeName == "data-srcset" && resource.Kind == HtmlResourceKind.Image && resource.IsAllowed);
        Assert.Contains(resourceManifest.Resources, resource => resource.ElementName == "object" && resource.AttributeName == "data" && resource.DiagnosticCode == "HtmlResourceRejectedByPolicy");
        Assert.Contains(resourceManifest.Resources, resource => resource.ElementName == "style" && resource.AttributeName == "css-import" && resource.DiagnosticCode == "StylesheetResourceRejectedByPolicy");
        Assert.DoesNotContain(resourceManifest.Resources, resource => resource.ElementName == "style" && resource.AttributeName == "css-url" && resource.Source == "file:///secret/print.css");
        Assert.Contains(resourceManifest.Resources, resource => resource.ElementName == "style" && resource.AttributeName == "css-url" && resource.Kind == HtmlResourceKind.Font && resource.DiagnosticCode == "FontResourceRejectedByPolicy");
        Assert.Contains(resourceManifest.Resources, resource => resource.ElementName == "style" && resource.AttributeName == "css-url" && resource.Kind == HtmlResourceKind.Image && resource.IsAllowed);
        Assert.DoesNotContain(resourceManifest.Resources, resource => resource.ElementName == "base");
        Assert.Contains(resourceManifest.Resources, resource => resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(resourceManifest.Resources, resource => resource.DiagnosticCode == "HyperlinkRejectedByPolicy");
        Assert.Contains(resourceManifest.Resources, resource => resource.DiagnosticCode == "ScriptResourceRejectedByPolicy");
        Assert.Contains(resourceManifest.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == "HyperlinkRejectedByPolicy");

        HtmlRoundTripScore score = HtmlRoundTripScorer.Compare(sourceHtml, roundTripHtml);
        Assert.InRange(score.Score, 0.55D, 1.00D);
        Assert.Equal(1D, score.Metrics["headings"], 3);
        Assert.Equal(1D, score.Metrics["tables"], 3);

        Assert.True(HtmlDiagnosticCatalog.TryGet("ImageResourceRejectedByPolicy", out HtmlDiagnosticDefinition imageDefinition));
        Assert.Equal("ResourcePolicy", imageDefinition.Category);
        Assert.Contains("data URI", imageDefinition.Remediation, StringComparison.OrdinalIgnoreCase);
        Assert.True(HtmlDiagnosticCatalog.TryGet("ScriptResourceRejectedByPolicy", out HtmlDiagnosticDefinition scriptDefinition));
        Assert.Equal("ResourcePolicy", scriptDefinition.Category);

        IReadOnlyList<HtmlMarketScenario> marketScenarios = HtmlMarketScenarioCatalog.All;
        Assert.Contains(marketScenarios, scenario => scenario.Id == "invoice" && scenario.Profile == HtmlConversionProfile.Document);
        Assert.Contains(marketScenarios, scenario => scenario.Id == "dashboard-print" && scenario.Profile == HtmlConversionProfile.HighFidelityPrint);

        HtmlConversionProfileContract profile = HtmlConversionProfileContracts.Get(HtmlConversionProfile.Document);
        Assert.Contains("form-controls", profile.SupportedHtml);
        Assert.Contains("blocked resource reporting", profile.ResourceGuarantees);

        var galleryResult = new HtmlCapabilityGalleryResult(new HtmlCapabilityGalleryScenario(
            "market-report",
            "Market Report",
            "HTML Engine",
            "Exercises shared OfficeIMO HTML engine contracts."));
        galleryResult.AddArtifact(new HtmlCapabilityGalleryArtifact("source", "input-html", "market-report.input.html", "text/html", sourceHtml.Length, new string('0', 64)));
        galleryResult.Diagnostics.Add("OfficeIMO.Tests", "HtmlCommentSkipped", "Comment skipped for manifest catalog coverage.", HtmlDiagnosticSeverity.Info);
        var manifest = new HtmlCapabilityGalleryManifest(galleryResult, HtmlConversionProfile.Document, score, resourceManifest);
        string manifestMarkdown = HtmlCapabilityGalleryManifestWriter.ToMarkdown(manifest);
        Assert.Contains("Profile: Document", manifestMarkdown);
        Assert.Contains("Round Trip Score", manifestMarkdown);
        Assert.Contains("ImageResourceRejectedByPolicy", manifestMarkdown);
        Assert.Contains("[ContentSimplification]", manifestMarkdown);

        HtmlToWordOptions untrusted = HtmlToWordOptions.CreateUntrustedHtmlProfile();
        HtmlToWordOptions trusted = HtmlToWordOptions.CreateTrustedDocumentProfile();
        Assert.Equal(HtmlConversionProfile.Semantic, untrusted.ConversionProfile);
        Assert.Equal(HtmlConversionProfile.Document, trusted.Clone().ConversionProfile);
    }

    [Fact]
    public void HtmlEnginePlatform_LogicalBuilderCountsOnlyRetainedNodes() {
        HtmlLogicalDocument logical = HtmlLogicalDocumentBuilder.FromHtml("<main> \r\n <p>Hello</p> \t </main>");

        Assert.Equal(0, logical.Count(HtmlLogicalNodeKind.Unknown));
        Assert.True(logical.Count(HtmlLogicalNodeKind.Text) >= 1);
        Assert.Equal(4, logical.GetCounts().Values.Sum());

        HtmlLogicalDocument scriptBody = HtmlLogicalDocumentBuilder.FromHtml("<main><p>Hello</p><script>console.log('internal')</script><style>.x{display:none}</style></main>");
        Assert.Equal(0, scriptBody.Count(HtmlLogicalNodeKind.Metadata));
        Assert.Equal(4, scriptBody.GetCounts().Values.Sum());
    }

    [Fact]
    public void HtmlEnginePlatform_RoundTripScorerUsesLogicalTextAndTailWindows() {
        HtmlLogicalDocument source = HtmlLogicalDocumentBuilder.FromHtml("<main><p>Alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi tail-one</p></main>");
        HtmlLogicalDocument target = HtmlLogicalDocumentBuilder.FromHtml("<main><p>Alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi tail-two</p></main>");

        HtmlRoundTripScore logicalScore = HtmlRoundTripScorer.Compare(source, target);
        Assert.InRange(logicalScore.Metrics["text"], 0D, 0.99D);

        HtmlRoundTripScore htmlScore = HtmlRoundTripScorer.Compare(
            "<main><p>0123456789 0123456789 0123456789 identical-prefix trailing-alpha</p></main>",
            "<main><p>0123456789 0123456789 0123456789 identical-prefix trailing-beta</p></main>");
        Assert.InRange(htmlScore.Metrics["text"], 0D, 0.99D);

        HtmlRoundTripScore appendedTextScore = HtmlRoundTripScorer.Compare(
            "<main><p>Alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu</p></main>",
            "<main><p>Alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu inserted target-only content</p></main>");
        Assert.InRange(appendedTextScore.Metrics["text"], 0D, 0.99D);

        HtmlRoundTripScore nonVisibleTextScore = HtmlRoundTripScorer.Compare(
            "<main><p>Visible text stays.</p><script>console.log('not document text')</script><style>.x{color:red}</style></main>",
            "<main><p>Visible text stays.</p></main>");
        Assert.Equal(1D, nonVisibleTextScore.Metrics["text"], 3);

        HtmlRoundTripScore listScore = HtmlRoundTripScorer.Compare(
            "<main><ul><li>One</li><li>Two</li></ul></main>",
            "<main><div><p>One</p><p>Two</p></div></main>");
        Assert.InRange(listScore.Metrics["lists"], 0D, 0.99D);
        Assert.InRange(listScore.Metrics["list-items"], 0D, 0.99D);

        HtmlRoundTripScore paragraphScore = HtmlRoundTripScorer.Compare(
            "<main><p>Block text</p></main>",
            "<main><div>Block text</div></main>");
        Assert.InRange(paragraphScore.Metrics["paragraphs"], 0D, 0.99D);

        HtmlRoundTripScore formStateScore = HtmlRoundTripScorer.Compare(
            "<main><input type=\"checkbox\" name=\"approval\" value=\"approved\" checked></main>",
            "<main><input type=\"checkbox\" name=\"approval\" value=\"rejected\"></main>");
        Assert.Equal(1D, formStateScore.Metrics["forms"], 3);
        Assert.InRange(formStateScore.Metrics["form-state"], 0D, 0.99D);

        HtmlRoundTripScore textOnlyLossScore = HtmlRoundTripScorer.Compare(
            "<main><p>Critical retained text</p></main>",
            "<main></main>");
        Assert.DoesNotContain("images", textOnlyLossScore.Metrics.Keys);
        Assert.DoesNotContain("links", textOnlyLossScore.Metrics.Keys);
        Assert.InRange(textOnlyLossScore.Score, 0D, 0.50D);
    }

    [Fact]
    public void HtmlEnginePlatform_ComputedStylesHonorUniversalSpecificityAndInlineCssSyntax() {
        var classNames = new List<string>();
        for (int i = 1; i <= 101; i++) {
            classNames.Add("c" + i.ToString("000"));
        }

        string manyClassSelector = "." + string.Join(".", classNames);
        string manyClassAttribute = string.Join(" ", classNames);
        string html = $$"""
            <main>
                <style>
                    p { color: #aa0000; }
                    * { color: #0000aa; }
                    :not(p) { color: #aa0000; }
                    .x { color: #0000aa; }
                    :where(strong) { font-weight: 900; }
                    strong { font-weight: 400; }
                    @media all { em.media { text-transform: uppercase; } }
                    @media not screen { em.media { text-transform: lowercase; } }
                    @media screen and (max-width: 0px) { em.media { text-transform: lowercase; } }
                    @supports not (color: red) { em.media { text-transform: lowercase; } }
                    span.reset { color: #ff0000; border-color: #ff0000; }
                    span.reset-later { color: initial; }
                    span { color: #0000ff; }
                    #specificity-id { outline-color: #00ff00; }
                    {{manyClassSelector}} { outline-color: #ff0000; }
                    span.brand { font-family: Corporate !important; }
                </style>
                <div style="color: #123456"><p style="background-image: url('data:image/svg+xml;utf8,<svg></svg>'); font-family: 'A;B'; color: inherit;">Hello</p></div>
                <span class="x">Pseudo specificity</span>
                <strong>Where specificity</strong>
                <em class="media">Media rule</em>
                <div style="color: #123456"><span class="reset" style="color: initial; border-color: unset;">Reset</span></div>
                <span class="reset-later">Reset wins lower specificity</span>
                <span class="brand" style="font-family: 'Brand!important'">Important string</span>
                <span id="specificity-id" class="{{manyClassAttribute}}">Specificity tuple</span>
                <style media="not all">span.inactive { text-decoration-line: underline; }</style>
                <style media="speech">span.speech { text-decoration-line: underline; }</style>
                <span class="inactive">Inactive media</span>
                <span class="speech">Speech media</span>
            </main>
            """;

        var parsed = HtmlDocumentParser.ParseDocument(html);
        IReadOnlyDictionary<IElement, HtmlComputedStyle> styles = HtmlComputedStyleEngine.Compute(parsed);
        IElement paragraph = parsed.QuerySelector("p")!;
        IElement pseudo = parsed.QuerySelector("span.x")!;
        IElement where = parsed.QuerySelector("strong")!;
        IElement media = parsed.QuerySelector("em.media")!;
        IElement reset = parsed.QuerySelector("span.reset")!;
        IElement resetLater = parsed.QuerySelector("span.reset-later")!;
        IElement brand = parsed.QuerySelector("span.brand")!;
        IElement specificityId = parsed.QuerySelector("#specificity-id")!;
        IElement inactive = parsed.QuerySelector("span.inactive")!;
        IElement speech = parsed.QuerySelector("span.speech")!;

        Assert.Equal("#123456", styles[paragraph].GetValue("color"));
        Assert.Contains("data:image/svg+xml;utf8", styles[paragraph].GetValue("background-image"));
        Assert.Equal("'A;B'", styles[paragraph].GetValue("font-family"));
        Assert.Equal("rgba(0, 0, 170, 1)", styles[pseudo].GetValue("color"));
        Assert.Equal("400", styles[where].GetValue("font-weight"));
        Assert.Equal("uppercase", styles[media].GetValue("text-transform"));
        Assert.Equal(string.Empty, styles[reset].GetValue("color"));
        Assert.Equal(string.Empty, styles[reset].GetValue("border-color"));
        Assert.Equal(string.Empty, styles[resetLater].GetValue("color"));
        Assert.Equal("Corporate", styles[brand].GetValue("font-family"));
        Assert.Equal("rgba(0, 255, 0, 1)", styles[specificityId].GetValue("outline-color"));
        Assert.Equal(string.Empty, styles[inactive].GetValue("text-decoration-line"));
        Assert.Equal(string.Empty, styles[speech].GetValue("text-decoration-line"));

        HtmlRoundTripScore printHiddenTextScore = HtmlRoundTripScorer.Compare(
            "<main><style media=\"print\">.screen{display:none}</style><span class=\"screen\">Visible on screen</span></main>",
            "<main></main>");
        Assert.InRange(printHiddenTextScore.Metrics["text"], 0D, 0.99D);
    }

    [Fact]
    public void HtmlEnginePlatform_ResourcePipelineAvoidsMetadataDuplicatesAndFontMisclassification() {
        const string html = """
            <html>
            <head>
                <base href="file:///secret/">
                <link rel="modulepreload" href="https://example.test/app.js">
                <link rel="icon" href="https://example.test/favicon.png">
                <link rel="preload" as="image" imagesrcset="file:///secret/preload.png 1x, https://example.test/images/preload-large.png 2x">
                <style>
                    /* @import url('file:///secret/commented.css'); */
                    @import url('file:///secret/theme.css');
                    @import "https://example.test/themes/dark mode.css";
                    @import url('https://example.test/images/shared.png');
                    @importurl(file:///secret/import-token.css);
                    :root { --hero: url(file:///secret/custom-property.png); }
                    .late { color: red; } @import url(file:///secret/late.css);
                    .comment-url { background-image: url('https://example.test/images/a/*v*/b.png'); }
                    .hero { background-image: url('https://example.test/images/bg.png'); }
                    .reuse { background-image: url('https://example.test/images/shared.png'); }
                    .logo::before { content: url('file:///secret/logo.png'); }
                    .label::before { content: "@import url(file:///secret/content.css)"; }
                    .label::before { content: "url(file:///secret/label.png)"; }
                </style>
                <meta http-equiv="refresh" content="0; url=file:///secret/refresh.html">
            </head>
            <body>
                <script src="mailto:ops@example.test"></script>
                <img src="mailto:ops@example.test">
                <svg><image href="https://example.test/images/vector.png" /><image xlink:href="file:///secret/vector-xlink.png" /></svg>
                <video poster="file:///secret/poster.png" data-src="https://example.test/media/movie.mp4"></video>
                <embed src="file:///secret/embed.pdf">
                <form action="file:///secret/upload"></form>
                <input type="image" src="file:///secret/submit.png">
                <input type="text" src="file:///secret/input-metadata">
                <div data="file:///secret/metadata"></div>
                <iframe src="file:///secret/ignored-frame.html" srcdoc="<img src=&quot;file:///secret/srcdoc.png&quot;><style>.nested { content: url(file:///secret/srcdoc-content.png); }</style>"></iframe>
                <div class="hero"></div>
            </body>
            </html>
            """;

        var manifest = HtmlResourcePipeline.BuildManifest(html, new HtmlResourcePipelineOptions {
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
        });

        Assert.DoesNotContain(manifest.Resources, resource => resource.ElementName == "base");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/commented.css");
        Assert.Single(manifest.Resources, resource => resource.Source == "file:///secret/theme.css");
        Assert.Single(manifest.Resources, resource => resource.Source == "https://example.test/themes/dark mode.css");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/shared.png" && resource.Kind == HtmlResourceKind.Stylesheet);
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/shared.png" && resource.Kind == HtmlResourceKind.Image);
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/content.css");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/label.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/custom-property.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/late.css");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/import-token.css");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/metadata");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/refresh.html" && resource.ElementName == "meta" && resource.Kind == HtmlResourceKind.Hyperlink && resource.AttributeName == "content" && resource.DiagnosticCode == "HyperlinkRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/a/*v*/b.png" && resource.Kind == HtmlResourceKind.Image);
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/app.js" && resource.Kind == HtmlResourceKind.Script);
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/favicon.png" && resource.Kind == HtmlResourceKind.Image);
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/preload.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "imagesrcset" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/preload-large.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "imagesrcset");
        Assert.Contains(manifest.Resources, resource => resource.Source == "mailto:ops@example.test" && resource.Kind == HtmlResourceKind.Script && !resource.IsAllowed && resource.DiagnosticCode == "ScriptResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "mailto:ops@example.test" && resource.Kind == HtmlResourceKind.Image && !resource.IsAllowed && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/vector.png" && resource.Kind == HtmlResourceKind.Image && resource.ElementName == "image");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/vector-xlink.png" && resource.Kind == HtmlResourceKind.Image && resource.ElementName == "image" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/poster.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "poster" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/media/movie.mp4" && resource.Kind == HtmlResourceKind.Media && resource.AttributeName == "data-src");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/embed.pdf" && resource.Kind == HtmlResourceKind.Other && resource.ElementName == "embed" && resource.AttributeName == "src");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/upload" && resource.Kind == HtmlResourceKind.Hyperlink && resource.AttributeName == "action");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/submit.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "src" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/input-metadata");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/logo.png" && resource.Kind == HtmlResourceKind.Image && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/ignored-frame.html");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/srcdoc.png" && resource.Kind == HtmlResourceKind.Image && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/srcdoc-content.png" && resource.Kind == HtmlResourceKind.Image && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/bg.png" && resource.Kind == HtmlResourceKind.Image && resource.IsAllowed);
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "https://example.test/images/bg.png" && resource.Kind == HtmlResourceKind.Font);
    }

    [Fact]
    public void HtmlEnginePlatform_RoundTripScorerComparesSemanticSignaturesAndVisibleText() {
        HtmlLogicalDocument textLeaf = HtmlLogicalDocumentBuilder.FromHtml("<main><custom>Total</custom></main>");
        Assert.Equal(1, textLeaf.Count(HtmlLogicalNodeKind.Text));
        Assert.Empty(textLeaf.Root.Children[0].Children[0].Children);
        HtmlLogicalDocument pictureWrapper = HtmlLogicalDocumentBuilder.FromHtml("<main><picture><source srcset=\"wide.png\"><img src=\"small.png\"></picture></main>");
        Assert.Equal(2, pictureWrapper.Count(HtmlLogicalNodeKind.Image));
        HtmlLogicalDocument mediaWrapper = HtmlLogicalDocumentBuilder.FromHtml("<main><video src=\"movie.mp4\"><source src=\"movie-hd.mp4\"><track src=\"captions.vtt\"></video></main>");
        Assert.Equal(3, mediaWrapper.Count(HtmlLogicalNodeKind.Media));
        Assert.Contains("media", mediaWrapper.Capabilities);
        HtmlLogicalDocument tableCaption = HtmlLogicalDocumentBuilder.FromHtml("<main><table><caption>Revenue</caption><tr><td>Total</td></tr></table></main>");
        Assert.Equal(1, tableCaption.Count(HtmlLogicalNodeKind.TableCaption));
        HtmlLogicalDocument inlineSvg = HtmlLogicalDocumentBuilder.FromHtml("<main><svg><image href=\"chart-a.png\"></image></svg></main>");
        Assert.Equal(2, inlineSvg.Count(HtmlLogicalNodeKind.Image));

        HtmlRoundTripScore repeatedTextScore = HtmlRoundTripScorer.Compare(
            "<main><p>" + new string('a', 100) + "</p></main>",
            "<main><p>" + new string('a', 32) + "</p></main>");
        Assert.InRange(repeatedTextScore.Metrics["text"], 0D, 0.99D);

        HtmlRoundTripScore hiddenTextScore = HtmlRoundTripScorer.Compare(
            "<main><p>Visible <span hidden>draft</span><span aria-hidden=\"true\">internal</span></p></main>",
            "<main><p>Visible</p></main>");
        Assert.Equal(1D, hiddenTextScore.Metrics["text"], 3);

        HtmlRoundTripScore stylesheetHiddenTextScore = HtmlRoundTripScorer.Compare(
            "<main><style>.draft{display:none}.private{visibility:hidden}</style><p>Visible <span class=\"draft\">draft</span><span class=\"private\">internal</span></p></main>",
            "<main><p>Visible</p></main>");
        Assert.Equal(1D, stylesheetHiddenTextScore.Metrics["text"], 3);

        HtmlRoundTripScore pictureSourceScore = HtmlRoundTripScorer.Compare(
            "<main><picture><source srcset=\"wide.png\"><img src=\"small.png\"></picture></main>",
            "<main><img src=\"small.png\"></main>");
        Assert.InRange(pictureSourceScore.Metrics["images"], 0D, 0.99D);
        Assert.InRange(pictureSourceScore.Metrics["image-sources"], 0D, 0.99D);

        HtmlRoundTripScore mediaScore = HtmlRoundTripScorer.Compare(
            "<main><video src=\"movie.mp4\"><source src=\"movie-hd.mp4\"></video></main>",
            "<main></main>");
        Assert.InRange(mediaScore.Metrics["media"], 0D, 0.99D);
        Assert.InRange(mediaScore.Metrics["media-sources"], 0D, 0.99D);

        HtmlRoundTripScore headingLevelScore = HtmlRoundTripScorer.Compare(
            "<main><h1>Title</h1></main>",
            "<main><h6>Title</h6></main>");
        Assert.Equal(1D, headingLevelScore.Metrics["headings"], 3);
        Assert.InRange(headingLevelScore.Metrics["heading-levels"], 0D, 0.99D);

        HtmlRoundTripScore tableShapeScore = HtmlRoundTripScorer.Compare(
            "<main><table><tr><td>A</td><td>B</td></tr><tr><td>C</td><td>D</td></tr></table></main>",
            "<main><table><tr><td>A</td><td>B</td><td>C</td><td>D</td></tr></table></main>");
        Assert.Equal(1D, tableShapeScore.Metrics["tables"], 3);
        Assert.InRange(tableShapeScore.Metrics["table-rows"], 0D, 0.99D);
        Assert.InRange(tableShapeScore.Metrics["table-grid"], 0D, 0.99D);
        HtmlRoundTripScore tableSpanScore = HtmlRoundTripScorer.Compare(
            "<main><table><tr><td colspan=\"2\">A</td></tr></table></main>",
            "<main><table><tr><td colspan=\"3\">A</td></tr></table></main>");
        Assert.Equal(1D, tableSpanScore.Metrics["tables"], 3);
        Assert.Equal(1D, tableSpanScore.Metrics["table-rows"], 3);
        Assert.Equal(1D, tableSpanScore.Metrics["table-cells"], 3);
        Assert.InRange(tableSpanScore.Metrics["table-grid"], 0D, 0.99D);

        HtmlRoundTripScore tableHeaderScore = HtmlRoundTripScorer.Compare(
            "<main><table><tr><th>Metric</th></tr></table></main>",
            "<main><table><tr><td>Metric</td></tr></table></main>");
        Assert.Equal(1D, tableHeaderScore.Metrics["tables"], 3);
        Assert.Equal(1D, tableHeaderScore.Metrics["table-cells"], 3);
        Assert.InRange(tableHeaderScore.Metrics["table-grid"], 0D, 0.99D);

        HtmlRoundTripScore nestedTableScore = HtmlRoundTripScorer.Compare(
            "<main><table><tr><td><table><tr><td>A</td></tr></table></td></tr></table></main>",
            "<main><table><tr><td><table><tr><td>A</td><td>B</td></tr></table></td></tr></table></main>");
        Assert.InRange(nestedTableScore.Metrics["table-grid"], 0.49D, 0.51D);

        HtmlRoundTripScore tableCaptionScore = HtmlRoundTripScorer.Compare(
            "<main><table><caption>Revenue</caption><tr><td>Total</td></tr></table></main>",
            "<main><p>Revenue</p><table><tr><td>Total</td></tr></table></main>");
        Assert.Equal(1D, tableCaptionScore.Metrics["tables"], 3);
        Assert.InRange(tableCaptionScore.Metrics["table-captions"], 0D, 0.99D);

        HtmlRoundTripScore listKindScore = HtmlRoundTripScorer.Compare(
            "<main><ol><li>Step</li></ol></main>",
            "<main><ul><li>Step</li></ul></main>");
        Assert.Equal(1D, listKindScore.Metrics["lists"], 3);
        Assert.Equal(1D, listKindScore.Metrics["list-items"], 3);
        Assert.InRange(listKindScore.Metrics["list-kinds"], 0D, 0.99D);

        HtmlRoundTripScore formTargetScore = HtmlRoundTripScorer.Compare(
            "<main><form action=\"/save\" method=\"post\"><input name=\"x\"></form></main>",
            "<main><form action=\"/delete\" method=\"post\"><input name=\"x\"></form></main>");
        Assert.Equal(1D, formTargetScore.Metrics["forms"], 3);
        Assert.InRange(formTargetScore.Metrics["form-state"], 0D, 0.99D);

        HtmlRoundTripScore formAssociationScore = HtmlRoundTripScorer.Compare(
            "<main><form action=\"/save\"><input name=\"x\"></form><form action=\"/delete\"><input name=\"y\"></form></main>",
            "<main><form action=\"/save\"><input name=\"y\"></form><form action=\"/delete\"><input name=\"x\"></form></main>");
        Assert.Equal(1D, formAssociationScore.Metrics["forms"], 3);
        Assert.InRange(formAssociationScore.Metrics["form-state"], 0D, 0.99D);

        HtmlRoundTripScore linkScore = HtmlRoundTripScorer.Compare(
            "<main><a href=\"https://example.test/a\">same text</a></main>",
            "<main><a href=\"https://example.test/b\">same text</a></main>");
        Assert.Equal(1D, linkScore.Metrics["links"], 3);
        Assert.InRange(linkScore.Metrics["link-targets"], 0D, 0.99D);

        HtmlRoundTripScore imageScore = HtmlRoundTripScorer.Compare(
            "<main><img src=\"https://example.test/a.png\" alt=\"Chart A\"></main>",
            "<main><img src=\"https://example.test/b.png\" alt=\"Chart B\"></main>");
        Assert.Equal(1D, imageScore.Metrics["images"], 3);
        Assert.InRange(imageScore.Metrics["image-sources"], 0D, 0.99D);

        HtmlRoundTripScore svgImageScore = HtmlRoundTripScorer.Compare(
            "<main><svg><image href=\"https://example.test/a.svg\"></image></svg></main>",
            "<main><svg><image href=\"https://example.test/b.svg\"></image></svg></main>");
        Assert.Equal(1D, svgImageScore.Metrics["images"], 3);
        Assert.InRange(svgImageScore.Metrics["image-sources"], 0D, 0.99D);

        HtmlRoundTripScore figureScore = HtmlRoundTripScorer.Compare(
            "<main><figure><img src=\"https://example.test/chart.png\" alt=\"Chart\"><figcaption>Revenue</figcaption></figure></main>",
            "<main><div><img src=\"https://example.test/chart.png\" alt=\"Chart\"><p>Revenue</p></div></main>");
        Assert.InRange(figureScore.Metrics["figures"], 0D, 0.99D);
        Assert.InRange(figureScore.Metrics["figure-signatures"], 0D, 0.99D);
    }
}
