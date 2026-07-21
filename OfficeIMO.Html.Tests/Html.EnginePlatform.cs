using AngleSharp.Dom;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Html;
using OfficeIMO.Word.Html;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Html {
    [Fact]
    public void HtmlSupportMatrix_IsGeneratedDeterministicallyFromProfilesAndDiagnostics() {
        string first = HtmlSupportMatrixWriter.ToMarkdown();
        string second = HtmlSupportMatrixWriter.ToMarkdown();

        Assert.Equal(first, second);
        Assert.DoesNotContain("\r", first, StringComparison.Ordinal);
        Assert.Contains("generated from `HtmlConversionProfileContracts`, `HtmlTargetCapabilityContracts`, and `HtmlDiagnosticCatalog`", first, StringComparison.Ordinal);
        foreach (HtmlConversionProfileContract contract in HtmlConversionProfileContracts.All) {
            Assert.Contains("### " + contract.Name, first, StringComparison.Ordinal);
        }
        Assert.Equal(Enum.GetValues<HtmlConversionTarget>().Length, HtmlTargetCapabilityContracts.All.Count);
        foreach (HtmlTargetCapabilityContract contract in HtmlTargetCapabilityContracts.All) {
            Assert.Contains("| " + contract.Target + " |", first, StringComparison.Ordinal);
            foreach (HtmlSemanticFeature feature in Enum.GetValues<HtmlSemanticFeature>()) {
                Assert.True(Enum.IsDefined(contract.GetSupport(feature)));
            }
        }
        foreach (HtmlDiagnosticDefinition definition in HtmlDiagnosticCatalog.Ordered) {
            Assert.Contains("`" + definition.Code + "`", first, StringComparison.Ordinal);
        }

        Assert.Equal(HtmlDiagnosticCatalog.All.Count, HtmlDiagnosticCatalog.Ordered.Select(definition => definition.Code).Distinct(StringComparer.OrdinalIgnoreCase).Count());
        Assert.Equal(
            HtmlDiagnosticCatalog.Ordered.Select(definition => definition.Category + "\0" + definition.Code),
            HtmlDiagnosticCatalog.Ordered.Select(definition => definition.Category + "\0" + definition.Code).OrderBy(value => value, StringComparer.Ordinal));
    }

    [Fact]
    public void HtmlSupportMatrix_CheckedInArtifactMatchesExecutableContracts() {
        string outputPath = Path.Combine(FindRepositoryRoot(), "Docs", "officeimo.html-support-matrix.md");
        if (string.Equals(Environment.GetEnvironmentVariable("OFFICEIMO_UPDATE_HTML_SUPPORT_MATRIX"), "1", StringComparison.Ordinal)) {
            HtmlSupportMatrixWriter.WriteMarkdown(outputPath);
        }

        Assert.True(File.Exists(outputPath),
            "Generated HTML support matrix is missing. Run Build/Export-HtmlSupportMatrix.ps1.");
        string actual = File.ReadAllText(outputPath).Replace("\r\n", "\n");
        Assert.Equal(HtmlSupportMatrixWriter.ToMarkdown(), actual);
    }

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
        Assert.DoesNotContain(resourceManifest.Resources, resource => resource.ElementName == "source" && resource.AttributeName == "data-src" && resource.Source == "https://example.test/reports/media/lazy-demo.mp4");
        Assert.Contains(resourceManifest.Resources, resource => resource.ElementName == "source" && resource.AttributeName == "srcset" && resource.Kind == HtmlResourceKind.Image && resource.IsAllowed);
        Assert.DoesNotContain(resourceManifest.Resources, resource => resource.ElementName == "source" && resource.AttributeName == "data-srcset" && resource.Source == "https://example.test/reports/images/chart-lazy.png");
        Assert.Contains(resourceManifest.Resources, resource => resource.ElementName == "object" && resource.AttributeName == "data" && resource.DiagnosticCode == "HtmlResourceRejectedByPolicy");
        Assert.Contains(resourceManifest.Resources, resource => resource.ElementName == "style" && resource.AttributeName == "css-import" && resource.DiagnosticCode == "StylesheetResourceRejectedByPolicy");
        Assert.DoesNotContain(resourceManifest.Resources, resource => resource.ElementName == "style" && resource.AttributeName == "css-url" && resource.Source == "file:///secret/print.css");
        Assert.Contains(resourceManifest.Resources, resource => resource.ElementName == "style" && resource.AttributeName == "css-url" && resource.Kind == HtmlResourceKind.Font && resource.DiagnosticCode == "FontResourceRejectedByPolicy");
        Assert.Contains(resourceManifest.Resources, resource => resource.ElementName == "style" && resource.AttributeName == "css-url" && resource.Kind == HtmlResourceKind.Image && resource.IsAllowed);
        Assert.DoesNotContain(resourceManifest.Resources, resource => resource.ElementName == "base");
        Assert.Contains(resourceManifest.Resources, resource => resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(resourceManifest.Resources, resource => resource.DiagnosticCode == "HyperlinkRejectedByPolicy");
        Assert.Contains(resourceManifest.Resources, resource => resource.DiagnosticCode == "ScriptResourceRejectedByPolicy");
        Assert.Contains(resourceManifest.Diagnostics, diagnostic => diagnostic.Code == "HyperlinkRejectedByPolicy");

        HtmlRoundTripScore score = HtmlRoundTripScorer.Compare(sourceHtml, roundTripHtml);
        Assert.InRange(score.Score, 0.10D, 0.55D);
        Assert.Equal(1D, score.Metrics["headings"], 3);
        Assert.Equal(1D, score.Metrics["tables"], 3);
        Assert.InRange(score.Dimensions["styles"], 0D, 0.50D);
        Assert.InRange(score.Dimensions["resources"], 0D, 0.50D);

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
        HtmlConversionProfileContract positionedReview = HtmlConversionProfileContracts.Get(HtmlConversionProfile.PositionedReview);
        Assert.Equal(4, HtmlConversionProfileContracts.All.Count);
        Assert.Equal("Positioned Review", positionedReview.Name);
        Assert.Contains("page wrappers", positionedReview.SupportedHtml);
        Assert.Contains("absolute positioning", positionedReview.SupportedCss);
        Assert.Contains("no-editable-reconstruction boundary", positionedReview.DiagnosticGuarantees);

        var galleryResult = new HtmlCapabilityGalleryResult(new HtmlCapabilityGalleryScenario(
            "market-report",
            "Market Report",
            "HTML Engine",
            "Exercises shared OfficeIMO HTML engine contracts."));
        galleryResult.AddArtifact(new HtmlCapabilityGalleryArtifact("source", "input-html", "market-report.input.html", "text/html", sourceHtml.Length, new string('0', 64)));
        galleryResult.Diagnostics.Add("OfficeIMO.Tests", "HtmlCommentSkipped", "Comment skipped for manifest catalog coverage.", HtmlDiagnosticSeverity.Info);
        var expectations = new[] {
            new HtmlCapabilityGalleryExpectation("headings", HtmlCapabilityGalleryExpectationOutcome.Preserved, "roundtrip HTML contains h1"),
            new HtmlCapabilityGalleryExpectation("blocked resources", HtmlCapabilityGalleryExpectationOutcome.Blocked, "resource manifest reports rejected data URI image"),
            new HtmlCapabilityGalleryExpectation("comments", HtmlCapabilityGalleryExpectationOutcome.Reported, "HtmlCommentSkipped diagnostic is present")
        };
        var manifest = new HtmlCapabilityGalleryManifest(galleryResult, HtmlConversionProfile.Document, score, resourceManifest, expectations);
        string manifestMarkdown = HtmlCapabilityGalleryManifestWriter.ToMarkdown(manifest);
        string manifestJson = HtmlCapabilityGalleryManifestJsonWriter.ToJson(manifest);
        using JsonDocument manifestJsonDocument = JsonDocument.Parse(manifestJson);
        JsonElement manifestJsonRoot = manifestJsonDocument.RootElement;
        Assert.Contains("Profile: Document", manifestMarkdown);
        Assert.Contains("Profile Contract", manifestMarkdown);
        Assert.Contains("Supported HTML: semantic sections", manifestMarkdown);
        Assert.Contains("Diagnostic guarantees: diagnostic catalog lookup", manifestMarkdown);
        Assert.Contains("Roundtrip Expectations", manifestMarkdown);
        Assert.Contains("Preserved: headings => roundtrip HTML contains h1", manifestMarkdown);
        Assert.Contains("Blocked: blocked resources => resource manifest reports rejected data URI image", manifestMarkdown);
        Assert.Contains("Round Trip Score", manifestMarkdown);
        Assert.Contains("ImageResourceRejectedByPolicy", manifestMarkdown);
        Assert.Contains("[ContentSimplification]", manifestMarkdown);
        Assert.Equal("officeimo.html.capability-gallery", manifestJsonRoot.GetProperty("schemaId").GetString());
        Assert.Equal("1.0", manifestJsonRoot.GetProperty("schemaVersion").GetString());
        Assert.Equal("market-report", manifestJsonRoot.GetProperty("scenario").GetProperty("id").GetString());
        Assert.Equal("Document", manifestJsonRoot.GetProperty("profile").GetProperty("id").GetString());
        Assert.Equal(3, manifestJsonRoot.GetProperty("expectations").GetArrayLength());
        Assert.Equal("source", manifestJsonRoot.GetProperty("artifacts")[0].GetProperty("id").GetString());
        Assert.True(manifestJsonRoot.GetProperty("roundTripScore").GetProperty("score").GetDouble() >= 0D);
        Assert.True(manifestJsonRoot.GetProperty("resources").GetProperty("blockedCount").GetInt32() > 0);
        Assert.Contains("ImageResourceRejectedByPolicy", manifestJson);
        Assert.Contains("\"origin\": \"resource\"", manifestJson);

        HtmlToWordOptions untrusted = HtmlToWordOptions.CreateUntrustedHtmlProfile();
        HtmlToWordOptions trusted = HtmlToWordOptions.CreateTrustedDocumentProfile();
        Assert.Equal(ImageProcessingMode.EmbedDataUriOnly, untrusted.ImageProcessing);
        Assert.True(trusted.Clone().AllowDocumentStylesheetLinks);
    }

    [Fact]
    public void HtmlEnginePlatform_DeclaresOfficeHtmlLaneContracts() {
        Assert.Equal(7, OfficeHtmlConversionProfileContracts.All.Count);

        OfficeHtmlConversionProfileContract wordSemantic = OfficeHtmlConversionProfileContracts.Get(OfficeHtmlConversionProfile.WordSemanticDocument);
        Assert.Equal("Word", wordSemantic.SourceFormat);
        Assert.Equal(HtmlConversionProfile.Semantic, wordSemantic.SharedProfile);
        Assert.Equal("none", wordSemantic.VisualPrimitiveOwner);
        Assert.Contains("footnotes", wordSemantic.SupportedHtml);
        Assert.Contains("layout simplification diagnostics", wordSemantic.DiagnosticGuarantees);

        OfficeHtmlConversionProfileContract wordRoundTrip = OfficeHtmlConversionProfileContracts.Get(OfficeHtmlConversionProfile.WordDocumentRoundTrip);
        Assert.Equal(HtmlConversionProfile.Document, wordRoundTrip.SharedProfile);
        Assert.Contains("form controls", wordRoundTrip.SupportedHtml);
        Assert.Contains("OpenXML validation proof", wordRoundTrip.DiagnosticGuarantees);

        OfficeHtmlConversionProfileContract wordPrint = OfficeHtmlConversionProfileContracts.Get(OfficeHtmlConversionProfile.WordPrintReview);
        Assert.Equal(HtmlConversionProfile.HighFidelityPrint, wordPrint.SharedProfile);
        Assert.Equal("OfficeIMO.Pdf", wordPrint.VisualPrimitiveOwner);
        Assert.Contains("section wrappers", wordPrint.SupportedHtml);
        Assert.Contains("print-fidelity boundary", wordPrint.DiagnosticGuarantees);

        OfficeHtmlConversionProfileContract excelSemantic = OfficeHtmlConversionProfileContracts.Get(OfficeHtmlConversionProfile.ExcelSemanticTables);
        Assert.Equal("Excel", excelSemantic.SourceFormat);
        Assert.Equal(HtmlConversionProfile.Semantic, excelSemantic.SharedProfile);
        Assert.Equal("none", excelSemantic.VisualPrimitiveOwner);
        Assert.Contains("tables", excelSemantic.SupportedHtml);
        Assert.Contains("formula/display-value diagnostics", excelSemantic.DiagnosticGuarantees);

        OfficeHtmlConversionProfileContract excelVisual = OfficeHtmlConversionProfileContracts.Get(OfficeHtmlConversionProfile.ExcelVisualReview);
        Assert.Equal(HtmlConversionProfile.PositionedReview, excelVisual.SharedProfile);
        Assert.Equal("OfficeIMO.Drawing", excelVisual.VisualPrimitiveOwner);
        Assert.Contains("positioned cell regions", excelVisual.SupportedHtml);
        Assert.Contains("Drawing-owned image and shape rendering", excelVisual.ResourceGuarantees);

        OfficeHtmlConversionProfileContract powerPointSemantic = OfficeHtmlConversionProfileContracts.Get(OfficeHtmlConversionProfile.PowerPointSemanticSlides);
        Assert.Equal("PowerPoint", powerPointSemantic.SourceFormat);
        Assert.Equal(HtmlConversionProfile.Semantic, powerPointSemantic.SharedProfile);
        Assert.Contains("speaker notes", powerPointSemantic.SupportedHtml);
        Assert.Contains("reading-order diagnostics", powerPointSemantic.DiagnosticGuarantees);

        OfficeHtmlConversionProfileContract powerPointVisual = OfficeHtmlConversionProfileContracts.Get(OfficeHtmlConversionProfile.PowerPointVisualReview);
        Assert.Equal(HtmlConversionProfile.PositionedReview, powerPointVisual.SharedProfile);
        Assert.Equal("OfficeIMO.Drawing", powerPointVisual.VisualPrimitiveOwner);
        Assert.Contains("positioned text frames", powerPointVisual.SupportedHtml);
        Assert.Contains("Drawing-owned slide rendering", powerPointVisual.ResourceGuarantees);

        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeHtmlConversionProfileContracts.Get((OfficeHtmlConversionProfile)99));
    }

    [Fact]
    public void HtmlEnginePlatform_BuildsCanonicalConversionDocumentAndNormalizedHtmlForAdapters() {
        const string html = """
            <!doctype html>
            <html>
            <head>
                <base href="https://example.test/reports/">
                <title>Canonical Contract</title>
                <link rel="stylesheet" href="/assets/report.css">
                <style>
                    body { color: #123456; font-family: Aptos; }
                    .secret { display: none; }
                </style>
            </head>
            <body onclick="alert(1)">
                <main>
                    <h1>Canonical Contract</h1>
                    <p class="lead">Visible text</p>
                    <style>.raw > .child { content: "a & b"; }</style>
                    <p><span>Hello </span><span>world</span></p>
                    <p class="secret">Internal draft</p>
                    <img src="chart.png" srcset="chart.png 1x, file:///secret/chart.png 2x" alt="Chart">
                    <noscript><p>Enable billing exports manually</p><img src="fallback.png" alt="Fallback"></noscript>
                    <object data="file:///secret/object.pdf"></object>
                    <a href="javascript:alert(1)">Unsafe link</a>
                    <a href="">Current document</a>
                    <div style="background-image:url(file:///secret/inline.png); color: red"></div>
                    <svg viewBox="0 0 100 40" preserveAspectRatio="xMidYMid meet"><image href="vector.svg"></image></svg>
                    <iframe srcdoc="<a href='javascript:alert(1)' onclick='nested()'><img src='file:///secret/nested.png'></a>"></iframe>
                    <form action=""><input type="checkbox" checked><button formaction="">Submit</button></form>
                    <form action="submit"><input type="checkbox" checked></form>
                </main>
            </body>
            </html>
            """;

        HtmlConversionDocument conversion = OfficeIMO.Html.HtmlConversionDocument.Parse(html, new HtmlConversionDocumentOptions {
            Profile = HtmlConversionProfile.Document,
            BaseUri = new Uri("https://example.test/reports/"),
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
        });

        Assert.Equal(HtmlConversionProfile.Document, conversion.ProfileContract.Profile);
        Assert.Equal(HtmlInputTrust.Untrusted, conversion.Trust);
        Assert.Contains("forms", conversion.LogicalDocument.Capabilities);
        Assert.Contains("images", conversion.LogicalDocument.Capabilities);
        Assert.Contains("font-family", conversion.StyleSummary.PropertyNames);
        Assert.Contains("Aptos", conversion.StyleSummary.FontFamilies);
        Assert.True(conversion.StyleSummary.HiddenElementCount >= 1);

        HtmlResourceDependencySummary imageSummary = conversion.ResourcePlan.GetSummary(HtmlResourceKind.Image);
        Assert.True(imageSummary.AllowedCount >= 1);
        Assert.True(imageSummary.BlockedCount >= 1);
        Assert.True(conversion.ResourcePlan.HasBlockedResources);

        Assert.Contains("https://example.test/reports/chart.png", conversion.NormalizedHtml);
        Assert.Contains("https://example.test/reports/chart.png", conversion.HtmlForConversion);
        Assert.Contains("Enable billing exports manually", conversion.HtmlForConversion);
        Assert.Contains("https://example.test/reports/fallback.png", conversion.HtmlForConversion);
        Assert.Contains("href=\"https://example.test/reports/\"", conversion.NormalizedHtml);
        Assert.Contains("action=\"https://example.test/reports/\"", conversion.NormalizedHtml);
        Assert.Contains("formaction=\"https://example.test/reports/\"", conversion.NormalizedHtml);
        Assert.Contains("https://example.test/reports/submit", conversion.HtmlForConversion);
        Assert.Contains("<head>", conversion.HtmlForConversion);
        Assert.Contains("body { color: #123456; font-family: Aptos; }", conversion.HtmlForConversion);
        Assert.Contains(".raw > .child", conversion.NormalizedHtml);
        Assert.DoesNotContain(".raw &gt; .child", conversion.NormalizedHtml);
        Assert.Contains("Hello </span><span>world", conversion.NormalizedHtml);
        Assert.Contains("checked", conversion.NormalizedHtml);
        Assert.Contains("viewBox", conversion.NormalizedHtml);
        Assert.Contains("preserveAspectRatio", conversion.NormalizedHtml);
        Assert.Contains("srcdoc=", conversion.NormalizedHtml);
        Assert.DoesNotContain("javascript:", conversion.NormalizedHtml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("onclick", conversion.NormalizedHtml, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("file:///secret", conversion.NormalizedHtml, StringComparison.OrdinalIgnoreCase);

        string markdown = conversion.ToMarkdown();
        Assert.Contains("# Canonical Contract", markdown);
        Assert.DoesNotContain("file:///secret", markdown, StringComparison.OrdinalIgnoreCase);
        HtmlToWordOptions sharedWordDefaults = WordHtmlConverterExtensions.CreateWordOptionsForSharedDocument();
        Assert.Equal(ImageProcessingMode.EmbedDataUriOnly, sharedWordDefaults.ImageProcessing);
        Assert.False(sharedWordDefaults.AllowDocumentStylesheetLinks);
        HtmlToWordOptions trustedWordDefaults = WordHtmlConverterExtensions.CreateWordOptionsForSharedDocument(
            HtmlInputTrust.Trusted);
        Assert.Equal(ImageProcessingMode.Embed, trustedWordDefaults.ImageProcessing);
        Assert.True(trustedWordDefaults.AllowDocumentStylesheetLinks);
        using var wordDocument = WordHtmlConverterExtensions.ToWordDocument(conversion);
        Assert.NotNull(wordDocument);

        HtmlConversionDocument nullPolicyConversion = OfficeIMO.Html.HtmlConversionDocument.Parse(
            "<main><img src=\"chart.png\" alt=\"Chart\"></main>",
            new HtmlConversionDocumentOptions {
                BaseUri = new Uri("https://example.test/reports/"),
                UrlPolicy = null!
            });
        Assert.Contains("https://example.test/reports/chart.png", nullPolicyConversion.NormalizedHtml);

        HtmlConversionDocument adapterOnlyConversion = OfficeIMO.Html.HtmlConversionDocument.Parse(
            "<main onclick=\"alert(1)\"><img src=\"file:///secret/chart.png\"><a href=\"javascript:alert(1)\">Unsafe</a></main>",
            new HtmlConversionDocumentOptions {
                IncludeNormalizedHtml = false,
                UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
            });
        Assert.Equal(string.Empty, adapterOnlyConversion.NormalizedHtml);
        Assert.DoesNotContain("onclick", adapterOnlyConversion.HtmlForConversion, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("javascript:", adapterOnlyConversion.HtmlForConversion, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("file:///secret", adapterOnlyConversion.HtmlForConversion, StringComparison.OrdinalIgnoreCase);

        HtmlConversionDocument reviewOnlyNormalization = OfficeIMO.Html.HtmlConversionDocument.Parse(
            "<html><head><style>.keep { color: red; }</style></head><body><p class=\"keep\">Styled</p></body></html>",
            new HtmlConversionDocumentOptions {
                NormalizationOptions = new HtmlNormalizationOptions {
                    PreserveStyleElements = false
                }
            });
        Assert.DoesNotContain("<style", reviewOnlyNormalization.NormalizedHtml, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<style", reviewOnlyNormalization.HtmlForConversion, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlEnginePlatform_NormalizationOptionsCanBeReusedAcrossDocuments() {
        var sharedOptions = new HtmlConversionDocumentOptions {
            IncludeNormalizedHtml = true,
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
            NormalizationOptions = new HtmlNormalizationOptions()
        };

        HtmlConversionDocument first = OfficeIMO.Html.HtmlConversionDocument.Parse(
            "<html><head><base href=\"https://first.example.test/assets/\"></head><body><img src=\"chart.png\"></body></html>",
            sharedOptions);
        HtmlConversionDocument second = OfficeIMO.Html.HtmlConversionDocument.Parse(
            "<html><head><base href=\"https://second.example.test/assets/\"></head><body><img src=\"chart.png\"></body></html>",
            sharedOptions);
        HtmlConversionDocument relativeBase = OfficeIMO.Html.HtmlConversionDocument.Parse(
            "<html><head><base href=\"assets/\"></head><body><img src=\"chart.png\"></body></html>",
            new HtmlConversionDocumentOptions {
                BaseUri = new Uri("https://example.test/root/page.html"),
                UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
                UseBodyContentsOnly = false
            });

        Assert.Contains("https://first.example.test/assets/chart.png", first.NormalizedHtml);
        Assert.Contains("https://second.example.test/assets/chart.png", second.NormalizedHtml);
        Assert.Contains("https://example.test/root/assets/chart.png", relativeBase.NormalizedHtml);
        Assert.Contains("<base href=\"https://example.test/root/assets/\">", relativeBase.NormalizedHtml);
        Assert.DoesNotContain("assets/assets", relativeBase.NormalizedHtml);
        Assert.Contains("<base href=\"https://example.test/root/assets/\">", relativeBase.HtmlForConversion);
        Assert.Contains("https://example.test/root/assets/chart.png", relativeBase.HtmlForConversion);
        Assert.DoesNotContain("https://example.test/root/chart.png", relativeBase.HtmlForConversion);
        Assert.DoesNotContain("assets/assets", relativeBase.HtmlForConversion);
        Assert.Null(sharedOptions.NormalizationOptions.BaseUri);
    }

    [Fact]
    public void HtmlEnginePlatform_HighFidelityPrintUsesPrintMediaContext() {
        const string html = """
            <html>
            <head>
                <link rel="stylesheet" media="screen" href="file:///secret/print-screen.css">
                <link rel="stylesheet" media="print/*c*/ and (color)" href="file:///secret/commented-print-link.css">
                <link rel="stylesheet" media="print" href="file:///secret/print-only-link.css">
                <style media="screen">
                    .total { color: #ff0000; }
                    .print-only { display: none; }
                </style>
                <style media="print">
                    .print-only { display: block; }
                </style>
                <style>
                    @media print { @supports (color: red) { .total { background-image: url(file:///secret/print-total.png); color: #123456; } } }
                    @media print and (future-feature) { .total { background-image: url(file:///secret/print-future.png); color: #ff00ff; } }
                    @media screen { .total { background-image: url(file:///secret/screen-total.png); } }
                </style>
                <style type="text/plain" media="print">
                    .total { color: #ff0000; }
                </style>
            </head>
            <body>
                <main>
                    <p class="total">Total</p>
                    <p class="print-only">Print target retained</p>
                    <picture>
                        <source media="screen" srcset="https://example.test/images/screen-chart.png 1x">
                        <source media="print/*c*/ and (color)" type="image/avif" srcset="https://example.test/images/ignored-print-chart.avif 1x">
                        <source media="print" srcset="https://example.test/images/print-chart.png 1x">
                        <img src="https://example.test/images/fallback-chart.png" alt="Chart">
                    </picture>
                </main>
            </body>
            </html>
            """;

        HtmlConversionDocument screen = OfficeIMO.Html.HtmlConversionDocument.Parse(html, new HtmlConversionDocumentOptions {
            Profile = HtmlConversionProfile.Semantic,
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
        });
        HtmlConversionDocument print = OfficeIMO.Html.HtmlConversionDocument.Parse(html, new HtmlConversionDocumentOptions {
            Profile = HtmlConversionProfile.HighFidelityPrint,
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
        });

        Assert.DoesNotContain(screen.ResourceManifest.Resources, resource => resource.Source == "file:///secret/print-total.png");
        Assert.DoesNotContain(screen.ResourceManifest.Resources, resource => resource.Source == "file:///secret/print-only-link.css");
        Assert.Contains(print.ResourceManifest.Resources, resource => resource.Source == "file:///secret/print-total.png" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(print.ResourceManifest.Resources, resource => resource.Source == "file:///secret/commented-print-link.css" && resource.DiagnosticCode == "StylesheetResourceRejectedByPolicy");
        Assert.Contains(print.ResourceManifest.Resources, resource => resource.Source == "file:///secret/print-only-link.css" && resource.DiagnosticCode == "StylesheetResourceRejectedByPolicy");
        Assert.DoesNotContain(print.ResourceManifest.Resources, resource => resource.Source == "file:///secret/print-future.png");
        Assert.Contains("background-image", print.StyleSummary.PropertyNames);

        string screenWordHtml = screen.CreateDocumentForConversion().DocumentElement?.OuterHtml ?? string.Empty;
        string printWordHtml = print.CreateDocumentForConversion().DocumentElement?.OuterHtml ?? string.Empty;
        string screenRetargetedToPrint = screen.CreateDocumentForConversion(HtmlCssMediaContext.Print).DocumentElement?.OuterHtml ?? string.Empty;
        Assert.Contains("https://example.test/images/screen-chart.png", screenWordHtml);
        Assert.DoesNotContain("https://example.test/images/print-chart.png", screenWordHtml);
        Assert.DoesNotContain("https://example.test/images/screen-chart.png", printWordHtml);
        Assert.DoesNotContain("https://example.test/images/ignored-print-chart.avif", printWordHtml);
        Assert.Contains("https://example.test/images/print-chart.png", printWordHtml);
        Assert.DoesNotContain(".print-only { display: none; }", screenRetargetedToPrint, StringComparison.Ordinal);
        Assert.Contains(".print-only { display: block; }", screenRetargetedToPrint, StringComparison.Ordinal);
        string printMarkdown = print.ToMarkdown();
        Assert.DoesNotContain("https://example.test/images/screen-chart.png", printMarkdown);
        Assert.DoesNotContain("https://example.test/images/ignored-print-chart.avif", printMarkdown);
        Assert.Contains("https://example.test/images/print-chart.png", printMarkdown);

        using var screenWordDocument = WordHtmlConverterExtensions.ToWordDocument(screen);
        var screenRun = screenWordDocument.Paragraphs.Single(paragraph => paragraph.Text.Contains("Total", StringComparison.Ordinal)).GetRuns().First();
        Assert.Equal("FF0000", screenRun.ColorHex);

        using var printWordDocument = WordHtmlConverterExtensions.ToWordDocument(print);
        var printRun = printWordDocument.Paragraphs.Single(paragraph => paragraph.Text.Contains("Total", StringComparison.Ordinal)).GetRuns().First();
        Assert.Equal("123456", printRun.ColorHex);

        byte[] printPdf = print.ToPdf();
        Assert.NotEmpty(printPdf);
        string retargetedPdfText = OfficeIMO.Pdf.PdfReadDocument.Open(screen.ToPdf()).ExtractText();
        Assert.Contains("Print target retained", retargetedPdfText, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlEnginePlatform_SharedManifestMatchesFilteredAdapterResources() {
        string blockedSrcSet = string.Join(", ", Enumerable.Range(1, 32).Select(index => $"file:///secret/blocked-{index}.png {index}w"));
        string html = $$"""
            <html>
            <head>
                <style>
                    .invalid-var { --hero: initial; background-image: var(--hero, url(file:///secret/invalid-var-fallback.png)); }
                    .escaped-property { background-\69mage: url(file:///secret/escaped-property.png); }
                </style>
            </head>
            <body>
                <main>
                    <picture><source srcset="mailto:ops@example.test"><img src="https://example.test/images/mailto-fallback.png" alt="Mail fallback"></picture>
                    <picture><source type="image/avif" srcset="https://example.test/images/ignored.avif 1x"><img src="https://example.test/images/adapter-fallback.png" alt="Type fallback"></picture>
                    <picture><source type="image/apng" srcset="https://example.test/images/ignored.apng 1x"><img src="https://example.test/images/apng-fallback.png" alt="APNG fallback"></picture>
                    <img alt="Late candidate" srcset="{{blockedSrcSet}}, https://example.test/images/late-candidate.png 33w">
                    <video><source src="mailto:media@example.test" type="video/mp4"><source src="https://example.test/media/fallback.mp4" type="video/mp4"></video>
                    <div class="invalid-var"></div>
                    <div class="escaped-property"></div>
                </main>
            </body>
            </html>
            """;

        HtmlConversionDocument document = OfficeIMO.Html.HtmlConversionDocument.Parse(html, new HtmlConversionDocumentOptions {
            Profile = HtmlConversionProfile.Semantic,
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
        });

        Assert.Contains(document.ResourceManifest.Resources, resource => resource.Source == "https://example.test/images/mailto-fallback.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "src");
        Assert.Contains(document.ResourceManifest.Resources, resource => resource.Source == "mailto:ops@example.test" && resource.Kind == HtmlResourceKind.Image && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.DoesNotContain(document.ResourceManifest.Resources, resource => resource.Source == "https://example.test/images/ignored.avif");
        Assert.Contains(document.ResourceManifest.Resources, resource => resource.Source == "https://example.test/images/adapter-fallback.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "src");
        Assert.DoesNotContain(document.ResourceManifest.Resources, resource => resource.Source == "https://example.test/images/ignored.apng");
        Assert.Contains(document.ResourceManifest.Resources, resource => resource.Source == "https://example.test/images/apng-fallback.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "src");
        Assert.Contains(document.ResourceManifest.Resources, resource => resource.Source == "https://example.test/images/late-candidate.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "srcset");
        Assert.Contains(document.ResourceManifest.Resources, resource => resource.Source == "mailto:media@example.test" && resource.Kind == HtmlResourceKind.Media && resource.DiagnosticCode == "MediaResourceRejectedByPolicy");
        Assert.Contains(document.ResourceManifest.Resources, resource => resource.Source == "https://example.test/media/fallback.mp4" && resource.Kind == HtmlResourceKind.Media && resource.AttributeName == "src");
        Assert.Contains(document.ResourceManifest.Resources, resource => resource.Source == "file:///secret/invalid-var-fallback.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(document.ResourceManifest.Resources, resource => resource.Source == "file:///secret/escaped-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");

        string markdown = document.ToMarkdown();
        Assert.DoesNotContain("https://example.test/images/ignored.avif", markdown);
        Assert.Contains("https://example.test/images/adapter-fallback.png", markdown);
        Assert.DoesNotContain("https://example.test/images/ignored.apng", markdown);
        Assert.Contains("https://example.test/images/apng-fallback.png", markdown);

        byte[] pdf = document.ToPdf();
        Assert.NotEmpty(pdf);

        string filteredFragment = HtmlActiveMediaFilter.Filter(
            "<picture><source media=\"print\" srcset=\"print.png\"><img src=\"screen.png\"></picture>",
            HtmlCssMediaContext.Screen);
        Assert.DoesNotContain("<html", filteredFragment, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("print.png", filteredFragment);
        Assert.Contains("screen.png", filteredFragment);
    }

    [Fact]
    public void HtmlEnginePlatform_ResourceManifestHandlesCiteSrcSetAndCustomPropertyCascadeEdges() {
        const string html = """
            <html>
            <head>
                <style>
                    :root { --escaped: url(file:///secret/escaped-var.png); }
                    .escaped { background-image: \76ar(--escaped); }
                    .card { --hero: url(file:///secret/stylesheet-var.png); background-image: var(--hero); }
                    .pseudo::before { --hero: url(file:///secret/pseudo-var.png); }
                    .pseudo { background-image: var(--hero, url(https://example.test/fallback.png)); }
                    .importantInline { --hero: url(file:///secret/important-inline.png) !important; background-image: var(--hero); }
                </style>
            </head>
            <body>
                <picture><source srcset="   "><img src="file:///secret/fallback-image.png" alt="Fallback"></picture>
                <blockquote cite="file:///secret/source.html">Quoted material</blockquote>
                <div class="escaped"></div>
                <div class="card" style="--hero: url(https://example.test/inline-var.png)"></div>
                <div class="pseudo"></div>
                <div class="importantInline" style="--hero: url(https://example.test/inline-overridden.png)"></div>
            </body>
            </html>
            """;

        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(
            html,
            new HtmlResourcePipelineOptions {
                UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
            });

        Assert.Contains(manifest.Resources, resource =>
            resource.ElementName == "img"
            && resource.AttributeName == "src"
            && resource.Source == "file:///secret/fallback-image.png"
            && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource =>
            resource.ElementName == "blockquote"
            && resource.AttributeName == "cite"
            && resource.Source == "file:///secret/source.html"
            && resource.DiagnosticCode == "HyperlinkRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource =>
            resource.ElementName == "style"
            && resource.AttributeName == "css-var-url"
            && resource.Source == "file:///secret/escaped-var.png"
            && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource =>
            resource.ElementName == "style"
            && resource.AttributeName == "css-var-url"
            && resource.Source == "https://example.test/inline-var.png"
            && resource.IsAllowed);
        Assert.Contains(manifest.Resources, resource =>
            resource.ElementName == "style"
            && resource.AttributeName == "css-url"
            && resource.Source == "https://example.test/fallback.png"
            && resource.IsAllowed);
        Assert.Contains(manifest.Resources, resource =>
            resource.ElementName == "style"
            && resource.AttributeName == "css-var-url"
            && resource.Source == "file:///secret/important-inline.png"
            && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/stylesheet-var.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/pseudo-var.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "https://example.test/inline-overridden.png");
    }

    [Fact]
    public void HtmlEnginePlatform_ResourceManifestHandlesVarFallbackInheritanceAndPseudoSelectors() {
        const string html = """
            <html>
            <head>
                <style>
                    :root {
                        --ok: url(https://example.test/ok.png);
                        --resolved: var(--ok, url(file:///secret/resolved-fallback.png));
                        --broken: var(--missing, url(file:///secret/alias-fallback.png));
                    }
                    .resolved { background-image: var(--resolved); }
                    .broken { background-image: var(--broken); }
                    .outer { --hero: var(--missing); background-image: var(--hero, url(file:///secret/outer-fallback.png)); }
                    .parent { --hero: url(https://example.test/inherited.png); }
                    .child { --hero: inherit; background-image: var(--hero); }
                    .unicode { --zażółć: url(file:///secret/unicode.png); background-image: var(--zażółć); }
                    .escapedName { --escaped-\6e ame: url(file:///secret/escaped-name.png); background-image: var(--escaped-\6e ame); }
                    a:hover { background-image: url(file:///secret/hover.png); }
                    ::before { content: url(file:///secret/bare-before.png); }
                </style>
            </head>
            <body>
                <div class="resolved"></div>
                <div class="broken"></div>
                <div class="outer"></div>
                <div class="parent"><div class="child"></div></div>
                <div class="unicode"></div>
                <div class="escapedName"></div>
                <a href="https://example.test/">Hover</a>
            </body>
            </html>
            """;

        HtmlResourceManifest manifest = HtmlResourcePipeline.BuildManifest(
            html,
            new HtmlResourcePipelineOptions {
                UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
            });

        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/ok.png" && resource.IsAllowed);
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/alias-fallback.png" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/outer-fallback.png" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/inherited.png" && resource.IsAllowed);
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/unicode.png" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/escaped-name.png" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/hover.png" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/bare-before.png" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/resolved-fallback.png");
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
        HtmlRoundTripScore retainedDocumentScore = HtmlRoundTripScorer.Compare(
            HtmlConversionDocument.Parse("<main><p>0123456789 0123456789 0123456789 identical-prefix trailing-alpha</p></main>"),
            HtmlConversionDocument.Parse("<main><p>0123456789 0123456789 0123456789 identical-prefix trailing-beta</p></main>"));
        Assert.Equal(htmlScore.Score, retainedDocumentScore.Score, 12);

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
                    @media not print and (color) { em.media { white-space: pre-wrap; } }
                    @media projection { em.media { font-style: italic; } }
                    @media not projection { em.media { font-size: 18px; } }
                    @media screen and (color) { em.media { direction: rtl; } }
                    @media screen and (max-width: 0px) { em.media { text-transform: lowercase; } }
                    @supports (not-a-real-prop: value) { em.media { text-transform: lowercase; } }
                    @supports not (color: red) { em.media { text-transform: lowercase; } }
                    @supports (display: block) or (not-a-real-prop: value) { em.media { text-decoration-line: underline; } }
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
        Assert.Equal("underline", styles[media].GetValue("text-decoration-line"));
        Assert.Equal("pre-wrap", styles[media].GetValue("white-space"));
        Assert.Equal(string.Empty, styles[media].GetValue("font-style"));
        Assert.Equal("18px", styles[media].GetValue("font-size"));
        Assert.Equal("rtl", styles[media].GetValue("direction"));
        Assert.Equal(string.Empty, styles[reset].GetValue("color"));
        Assert.Equal(string.Empty, styles[reset].GetValue("border-color"));
        Assert.Equal(string.Empty, styles[resetLater].GetValue("color"));
        Assert.Equal("Corporate", styles[brand].GetValue("font-family"));
        Assert.Equal("rgba(0, 255, 0, 1)", styles[specificityId].GetValue("outline-color"));
        Assert.Equal(string.Empty, styles[inactive].GetValue("text-decoration-line"));
        Assert.Equal(string.Empty, styles[speech].GetValue("text-decoration-line"));
        Assert.True(HtmlComputedStyleEngine.IsApplicableSupports("(--officeimo-probe: 0)"));

        HtmlRoundTripScore printHiddenTextScore = HtmlRoundTripScorer.Compare(
            "<main><style media=\"print\">.screen{display:none}</style><span class=\"screen\">Visible on screen</span></main>",
            "<main></main>");
        Assert.InRange(printHiddenTextScore.Metrics["text"], 0D, 0.99D);

        HtmlRoundTripScore notPrintHiddenTextScore = HtmlRoundTripScorer.Compare(
            "<main><style media=\"not print\">.draft{display:none}</style><p>Visible <span class=\"draft\">draft</span></p></main>",
            "<main><p>Visible</p></main>");
        Assert.Equal(1D, notPrintHiddenTextScore.Metrics["text"], 3);

        HtmlRoundTripScore notUnsupportedSupportsScore = HtmlRoundTripScorer.Compare(
            "<main><style>@supports not (not-a-real-prop: value){.draft{display:none}}</style><p>Visible <span class=\"draft\">draft</span></p></main>",
            "<main><p>Visible</p></main>");
        Assert.Equal(1D, notUnsupportedSupportsScore.Metrics["text"], 3);

        HtmlRoundTripScore supportsOrScore = HtmlRoundTripScorer.Compare(
            "<main><style>@supports (display:block) or (not-a-real-prop:value){.draft{display:none}}</style><p>Visible <span class=\"draft\">draft</span></p></main>",
            "<main><p>Visible</p></main>");
        Assert.Equal(1D, supportsOrScore.Metrics["text"], 3);

        HtmlRoundTripScore invalidSupportsValueScore = HtmlRoundTripScorer.Compare(
            "<main><style>@supports (display:not-a-real-value){.draft{display:none}}</style><p>Visible <span class=\"draft\">draft</span></p></main>",
            "<main><p>Visible</p></main>");
        Assert.InRange(invalidSupportsValueScore.Metrics["text"], 0D, 0.99D);

        HtmlRoundTripScore mediaListFallbackScore = HtmlRoundTripScorer.Compare(
            "<main><style>@media screen and (max-width:0px), screen {.draft{display:none}}</style><p>Visible <span class=\"draft\">draft</span></p></main>",
            "<main><p>Visible</p></main>");
        Assert.Equal(1D, mediaListFallbackScore.Metrics["text"], 3);

        HtmlRoundTripScore nonCssStyleScore = HtmlRoundTripScorer.Compare(
            "<main><style type=\"text/plain\">.screen{display:none}</style><span class=\"screen\">Visible</span></main>",
            "<main></main>");
        Assert.InRange(nonCssStyleScore.Metrics["text"], 0D, 0.99D);

        HtmlRoundTripScore inlineCommentHiddenTextScore = HtmlRoundTripScorer.Compare(
            "<main><span style=\"display:/*x*/none\">hidden</span></main>",
            "<main></main>");
        Assert.Equal(1D, inlineCommentHiddenTextScore.Metrics["text"], 3);

        HtmlRoundTripScore invalidInlineCssScore = HtmlRoundTripScorer.Compare(
            "<main><style>.draft{display:none}</style><span class=\"draft\" style=\"display:bogus\">hidden</span></main>",
            "<main></main>");
        Assert.Equal(1D, invalidInlineCssScore.Metrics["text"], 3);

        HtmlRoundTripScore spacedImportantScore = HtmlRoundTripScorer.Compare(
            "<main><style>.x{display:none!important}</style><span class=\"x\" style=\"display:inline ! important\">Visible</span></main>",
            "<main></main>");
        Assert.InRange(spacedImportantScore.Metrics["text"], 0D, 0.99D);
    }

    [Fact]
    public void HtmlEnginePlatform_ResourcePipelineAvoidsMetadataDuplicatesAndFontMisclassification() {
        const string html = """
            <html>
            <head>
                <base href="file:///secret/">
                <link rel="modulepreload" href="https://example.test/app.js">
                <link rel="icon" href="https://example.test/favicon.png">
                <link rel="notstylesheet" href="file:///secret/notstylesheet.css">
                <link rel="preload" as="image" imagesrcset="file:///secret/preload.png 1x, https://example.test/images/preload-large.png 2x">
                <link rel="preload" as="image" media="(max-width: 0px)" href="file:///secret/inactive-preload.png" imagesrcset="file:///secret/inactive-preload-2x.png 2x">
                <link rel="stylesheet" href="https://example.test/app.css" imagesrcset="file:///secret/stylesheet-image.png 1x">
                <style>
                    /* @import url('file:///secret/commented.css'); */
                    @import url('file:///secret/theme.css');
                    @import url(file:///secret/print-import.css) print;
                    @import url(file:///secret/screen-import.css) screen;
                    @import/*comment*/url('file:///secret/comment-import.css');
                    @import "https://example.test/themes/dark mode.css";
                    @import url('https://example.test/images/shared.png');
                    @import \75rl(file:///secret/escaped-import.css);
                    @importurl(file:///secret/import-token.css);
                    :root { --hero: url(file:///secret/custom-property.png); --used-hero: url(file:///secret/custom-property-used.png); }
                    :root { --alias-target: url(file:///secret/alias-target.png); --alias-hero: var(--alias-target); }
                    :root { --cross-block-hero: url(file:///secret/cross-block-property.png); --important-hero: url(file:///secret/important-property.png) !important; }
                    .unused { --card-hero: url(file:///secret/unused-custom-property.png); }
                    .theme { --theme-hero: url(file:///secret/inherited-custom-property.png); }
                    .theme-dom { --dom-hero: url(file:///secret/dom-inherited-property.png); }
                    .card-dom { background-image: var(--dom-hero); }
                    .card { --inline-card-hero: url(file:///secret/inline-card-property.png); }
                    .theme-late { --theme-late-hero: url(file:///secret/nearer-custom-property.png); }
                    :root { --theme-late-hero: url(https://example.test/images/later-root-property.png); }
                    .theme-late .card { background-image: var(--theme-late-hero); }
                    .nearest-inner { --nearest-hero: url(https://example.test/images/nearest-property.png); }
                    .nearest-outer { --nearest-hero: url(file:///secret/nearest-outer-property.png); }
                    .nearest-card { background-image: var(--nearest-hero); }
                    .theme-important { --rank-hero: url(https://example.test/images/rank-local.png); }
                    :root { --rank-hero: url(file:///secret/root-important-rank.png) !important; }
                    .theme-important .card { background-image: var(--rank-hero); }
                    :root { --spaced-important-hero: url(file:///secret/spaced-important-property.png) ! important; }
                    :root { --spaced-important-hero: url(https://example.test/images/spaced-important-property.png); }
                    .spaced-important { background-image: var(--spaced-important-hero); }
                    :is(.split-theme,.unused-theme) { --split-hero: url(file:///secret/split-selector-property.png); }
                    .split-theme .card { background-image: var(--split-hero); }
                    #specific-card { --specific-hero: url(file:///secret/specific-property.png); }
                    .specific-card { --specific-hero: url(https://example.test/images/specific-property.png); }
                    .specific-card { background-image: var(--specific-hero); }
                    .where-theme { --where-hero: url(https://example.test/images/where-ok.png); }
                    :where(.where-theme) { --where-hero: url(file:///secret/where-zero.png); }
                    .where-theme .card { background-image: var(--where-hero); }
                    :is(.is-theme) { --is-hero: url(file:///secret/is-functional.png); }
                    .is-theme { --is-hero: url(https://example.test/images/is-ok.png); }
                    .is-theme .card { background-image: var(--is-hero); }
                    .theme-fallback { --fallback-hero: url(https://example.test/images/inherited-fallback.png); }
                    .card-fallback { background-image: var(--fallback-hero, url(file:///secret/inherited-fallback.png)); }
                    .theme-element-fallback { --element-fallback-hero: url(https://example.test/images/element-fallback.png); }
                    .card-element-fallback { background-image: var(--element-fallback-hero, url(file:///secret/element-fallback.png)); }
                    @media screen { :root { --media-hero: url(file:///secret/grouped-custom-property.png); } .media-hero { background-image: var(--media-hero); } }
                    @media screen { @media print { .nested-inactive-media { background-image: url(file:///secret/nested-print-in-screen.png); } } }
                    :root { --inactive-use-hero: url(file:///secret/inactive-var-use.png); }
                    @media print { :root { --print-hero: url(file:///secret/print-custom-property.png); } .print-only { background-image: url(file:///secret/print-only.png); } .print-card { background-image: var(--print-hero); } .inactive-var-use { background-image: var(--inactive-use-hero); } }
                    @media not screen and (max-width: 0px) { .negated-active-feature { background-image: url(file:///secret/negated-active-feature.png); } }
                    @media projection { .projection-only { background-image: url(file:///secret/projection-only.png); } }
                    @media not projection { .not-projection { background-image: url(file:///secret/not-projection.png); } }
                    @supports (background-image:url(file:///secret/supports-condition.png)) { .ok { color: red; } }
                    @supports (not-a-real-prop:value) { .supports-inactive { background-image: url(file:///secret/supports-inactive.png); } }
                    @supports (--officeimo-probe: 0) { .custom-supports { background-image: url(file:///secret/custom-supports.png); } }
                    .late { color: red; } @import url(file:///secret/late.css);
                    .comment-url { background-image: url('https://example.test/images/a/*v*/b.png'); }
                    .hero { background-image: var(--used-hero, url(file:///secret/unused-fallback.png)), url('https://example.test/images/bg.png'); }
                    .defined-non-url { --defined-hero: none; background-image: var(--defined-hero, url(file:///secret/defined-fallback.png)); }
                    :root { --important-hero: url(https://example.test/images/non-important-property.png); }
                    .cross-block { background-image: var(--cross-block-hero); }
                    .important { background-image: var(--important-hero); }
                    .theme .card { background-image: var(--theme-hero); }
                    .cascaded { --cascaded-hero: url(file:///secret/old-cascaded-property.png); }
                    .cascaded { --cascaded-hero: url(https://example.test/images/cascaded-property.png); }
                    .cascaded .card { background-image: var(--cascaded-hero); }
                    .compound { --same-element-hero: url(file:///secret/same-element-property.png); }
                    .compound.highlight { background-image: var(--same-element-hero); }
                    .multi { --multi-hero: url(file:///secret/multi-property-a.png), url(https://example.test/images/multi-property-b.png); }
                    .multi .card { background-image: var(--multi-hero); }
                    :root { --notvar-hero: url(file:///secret/notvar-custom-property.png); }
                    .notvar { background-image: notvar(--notvar-hero); }
                    .escaped { background-image: url(\66 ile:///secret/escaped.png); }
                    .escaped-function { background-image: \75rl(file:///secret/escaped-function.png); }
                    .invalid-escape { background-image: url(\ffffff.png); }
                    .not-url { background-image: noturl(file:///secret/not-url-function.png); }
                    .unsupported-url { background-notimage: url(file:///secret/unsupported-property.png); }
                    .unsupported-image-set { background-notimage: image-set("file:///secret/unsupported-image-set.png" 1x); }
                    .unused-resource { background-image: url(file:///secret/unused-rule.png); }
                    @media (max-width: 0px) { .zero-media { background-image: url(file:///secret/zero-media.png); } }
                    @media(max-width:0px){.zero-media-minified{background-image:url(file:///secret/zero-media-minified.png)}}
                    @media (max-width: 0em) { .zero-media-em { background-image: url(file:///secret/zero-media-em.png); } }
                    @media (not-a-real-feature) { .unknown-feature { background-image: url(file:///secret/unknown-media-feature.png); } }
                    .masked { mask: url(file:///secret/mask.svg); }
                    .filtered { filter: url(file:///secret/filter.svg#f); clip-path: url(file:///secret/clip.svg#c); }
                    .fragment-only { filter: url(#blur); clip-path: url("#shape"); background-image: var(--fragment-hero); }
                    :root { --fragment-hero: url(#internal-image); }
                    .image-set { background-image: image-set("file:///secret/hero.png" 1x, url(https://example.test/images/hero-2x.png) 2x, "https://example.test/images/hero.avif" type("image/avif")); }
                    .escaped-image-set { background-image: \69mage-set("file:///secret/escaped-image-set.png" 1x); }
                    .escaped-type-image-set { background-image: image-set("https://example.test/images/typed.avif" \74ype("image/avif")); }
                    .commented-type-image-set { background-image: image-set("https://example.test/images/commented-type.avif" type(/*c*/"image/avif")); }
                    .fragment-image-set { background-image: image-set("#inline-sprite" 1x); }
                    .not-image-set { background-image: notimage-set("file:///secret/not-image-set.png" 1x); }
                    .reuse { background-image: url('https://example.test/images/shared.png'); }
                    .logo::before { content: url('file:///secret/logo.png'); }
                    .label::before { content: "@import url(file:///secret/content.css)"; }
                    .label::before { content: "url(file:///secret/label.png)"; }
                    .alias { background-image: var(--alias-hero); }
                    .inline-ancestor-card { background-image: var(--inline-ancestor-hero); }
                    .inline-self-card { background-image: var(--inline-self-hero); }
                    .inline-override-card { --inline-override-hero: url(https://example.test/images/inline-override.png); background-image: var(--inline-override-hero); }
                    a:hover { --hover-hero: url(file:///secret/hover-var.png); background-image: var(--hover-hero); }
                </style>
                <style media="print">.inactive-media-style { background-image: url(file:///secret/inactive-media-style.png); }</style>
                <style type="text/plain">.plain { background-image: url(file:///secret/plain-style.png); }</style>
                <meta http-equiv="refresh" content="0; url=file:///secret/refresh.html">
                <meta http-equiv="refresh" content="0; url='https://example.test/report;v=1.html'">
                <meta http-equiv="refresh" content="0; noturl=file:///secret/not-refresh.html">
            </head>
            <body background="file:///secret/body-bg.png">
                <script src="mailto:ops@example.test"></script>
                <script type="application/json" src="file:///secret/script-data.json"></script>
                <svg><script href="file:///secret/svg-script.js" /></svg>
                <img src="mailto:ops@example.test">
                <svg><defs><symbol id="icon" /></defs><image href="https://example.test/images/vector.png" /><image xlink:href="file:///secret/vector-xlink.png" /><feImage href="file:///secret/filter-image.png" /><use href="#icon" /><use href="file:///secret/symbols.svg#icon" /></svg>
                <img data-lazy-src="file:///secret/lazy-img.png" data-original-srcset="file:///secret/original-img.png 1x, https://example.test/images/original-img.png 2x">
                <video poster="file:///secret/poster.png" data-src="https://example.test/media/movie.mp4"></video>
                <video src="https://example.test/media/selected.mp4"><source src="file:///secret/ignored-child-source.mp4" type="video/mp4"></video>
                <video><source src="file:///secret/unsupported-child-source.mp4" type="video/not-real"></video>
                <video><source src="file:///secret/selected-child-source.mp4" type="video/mp4"></video>
                <video><source src="https://example.test/media/first-playable.mp4" type="video/mp4"><source src="file:///secret/unselected-playable.mp4" type="video/mp4"></video>
                <video><source src="file:///secret/blocked-video.mp4" type="video/mp4"><source src="https://example.test/media/policy-fallback.mp4" type="video/mp4"></video>
                <embed data="file:///secret/embed-data.pdf" src="file:///secret/embed.pdf">
                <form action="file:///secret/upload"></form>
                <button formaction="file:///secret/delete">Delete</button>
                <button type="bogus" formaction="file:///secret/bogus-delete">Delete anyway</button>
                <input type="image" src="file:///secret/submit.png" srcset="file:///secret/inert-submit-srcset.png 2x">
                <input type=" image " src="file:///secret/spaced-submit.png">
                <input type="text" src="file:///secret/input-metadata">
                <picture><source src="file:///secret/ignored-picture-source.png" srcset="https://example.test/images/picture-source.png 1x"><img src="https://example.test/images/picture-fallback.png"></picture>
                <picture><source srcset="https://example.test/images/selected-picture.png 1x"><source srcset="file:///secret/unselected-picture.png 1x"><img src="file:///secret/selected-fallback.png"></picture>
                <picture><source media="print" srcset="file:///secret/print-picture-source.png 1x"><img src="https://example.test/images/screen-picture.png"></picture>
                <picture><source media="(max-width: 0px)" srcset="file:///secret/zero-picture-source.png 1x"><img src="https://example.test/images/zero-picture-fallback.png"></picture>
                <picture><source type="image/not-real" srcset="file:///secret/unsupported-picture-type.png 1x"><img src="https://example.test/images/typed-fallback.png"></picture>
                <picture><source data-srcset="https://example.test/images/lazy-only.png 1x"><img src="file:///secret/lazy-fallback.png"></picture>
                <picture><source src="file:///secret/source-url-only.png"><img src="https://example.test/images/source-url-fallback.png"></picture>
                <picture><source srcset="file:///secret/policy-source.png 1x"><img src="https://example.test/images/policy-fallback.png"></picture>
                <picture><source srcset="file:///secret/blocked-wide.png 1x"><source srcset="https://example.test/images/next-wide.png 1x"><img src="https://example.test/images/unused-wide-fallback.png"></picture>
                <picture><source srcset="file:///secret/mixed-picture.png 1x, https://example.test/images/mixed-picture.png 2x"><source srcset="https://example.test/images/unused-mixed-next.png 1x"><img src="https://example.test/images/unused-mixed-fallback.png"></picture>
                <picture><source type="image/avif" srcset="https://example.test/images/word-skipped.avif 1x"><img src="https://example.test/images/word-fallback.png"></picture>
                <div data="file:///secret/metadata"></div>
                <x-card href="file:///secret/custom-href.png" src="file:///secret/custom-src.png"></x-card>
                <div background="file:///secret/not-legacy-background.png"></div>
                <div style="@import url(file:///secret/inline-import.css); background-image: url(https://example.test/images/inline.png)"></div>
                <iframe src="file:///secret/ignored-frame.html" srcdoc="<img src=&quot;file:///secret/srcdoc.png&quot;><style>.nested { content: url(file:///secret/srcdoc-content.png); }</style><div class=&quot;nested&quot;></div>"></iframe>
                <iframe src="file:///secret/empty-srcdoc-frame.html" srcdoc=""></iframe>
                <div class="theme"><div class="card"></div></div>
                <div class="theme-dom"><div class="card-dom"></div></div>
                <div class="theme-late"><div class="card"></div></div>
                <div class="nearest-outer"><div class="nearest-inner"><div class="nearest-card"></div></div></div>
                <div class="theme-important"><div class="card"></div></div>
                <div class="spaced-important"></div>
                <div class="split-theme"><div class="card"></div></div>
                <div id="specific-card" class="specific-card"></div>
                <div class="where-theme"><div class="card"></div></div>
                <div class="is-theme"><div class="card"></div></div>
                <div class="theme-fallback"><div class="card-fallback"></div></div>
                <div class="theme-element-fallback"><div class="card-element-fallback"></div></div>
                <div class="card-element-fallback"></div>
                <div class="cascaded"><div class="card"></div></div>
                <div class="cross-block"></div>
                <div class="important"></div>
                <div class="compound highlight"></div>
                <div class="multi"><div class="card"></div></div>
                <div class="hero"></div>
                <div class="alias"></div>
                <div class="zero-media"></div>
                <div class="zero-media-minified"></div>
                <div class="zero-media-em"></div>
                <div class="unknown-feature"></div>
                <div class="media-hero"></div>
                <div class="inactive-var-use"></div>
                <div class="negated-active-feature"></div>
                <div class="projection-only"></div>
                <div class="not-projection"></div>
                <div class="supports-inactive"></div>
                <div class="custom-supports"></div>
                <div class="comment-url"></div>
                <div class="escaped"></div>
                <div class="escaped-function"></div>
                <div class="nested-inactive-media"></div>
                <div class="masked"></div>
                <div class="filtered"></div>
                <div class="fragment-only"></div>
                <div class="image-set"></div>
                <div class="escaped-image-set"></div>
                <div class="escaped-type-image-set"></div>
                <div class="commented-type-image-set"></div>
                <div class="fragment-image-set"></div>
                <div class="reuse"></div>
                <div class="logo"></div>
                <div class="card" style="background-image: var(--inline-card-hero)"></div>
                <div style="--inline-inherited: url(file:///secret/inline-inherited-property.png)"><span style="background-image: var(--inline-inherited)"></span></div>
                <div style="--inline-ancestor-hero: url(file:///secret/inline-ancestor-stylesheet-property.png)"><div class="inline-ancestor-card"></div></div>
                <div class="inline-self-card" style="--inline-self-hero: url(file:///secret/inline-self-stylesheet-property.png)"></div>
                <div style="--inline-override-hero: url(file:///secret/inline-override-ancestor.png)"><div class="inline-override-card"></div></div>
                <a href="https://example.test/hover">Hover</a>
            </body>
            </html>
            """;

        var manifest = HtmlResourcePipeline.BuildManifest(html, new HtmlResourcePipelineOptions {
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
        });

        Assert.DoesNotContain(manifest.Resources, resource => resource.ElementName == "base");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/commented.css");
        Assert.Single(manifest.Resources, resource => resource.Source == "file:///secret/theme.css");
        Assert.Single(manifest.Resources, resource => resource.Source == "file:///secret/comment-import.css");
        Assert.Single(manifest.Resources, resource => resource.Source == "file:///secret/escaped-import.css");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/print-import.css");
        Assert.Single(manifest.Resources, resource => resource.Source == "file:///secret/screen-import.css");
        Assert.Single(manifest.Resources, resource => resource.Source == "https://example.test/themes/dark mode.css");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/shared.png" && resource.Kind == HtmlResourceKind.Stylesheet);
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/shared.png" && resource.Kind == HtmlResourceKind.Image);
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/content.css");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/label.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/custom-property.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/late.css");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/import-token.css");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/inline-import.css");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/notstylesheet.css" && resource.Kind == HtmlResourceKind.Stylesheet);
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/stylesheet-image.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/metadata");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/custom-href.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/custom-src.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/not-legacy-background.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/unused-custom-property.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/unused-rule.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/supports-condition.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/supports-inactive.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/print-only.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/print-custom-property.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/inactive-var-use.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/projection-only.png");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/not-projection.png" && resource.Kind == HtmlResourceKind.Image && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/inactive-media-style.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/inactive-preload.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/inactive-preload-2x.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/zero-media.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/zero-media-minified.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/zero-media-em.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/unknown-media-feature.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/nested-print-in-screen.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/zero-picture-source.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/plain-style.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/unused-fallback.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/defined-fallback.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/not-url-function.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/old-cascaded-property.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/root-important-rank.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/nearest-outer-property.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/where-zero.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/is-functional.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/inline-override-ancestor.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "https://example.test/images/spaced-important-property.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/inherited-fallback.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/not-image-set.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/notvar-custom-property.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/unsupported-property.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/unsupported-image-set.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/embed-data.pdf");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/script-data.json");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/not-refresh.html");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/inert-submit-srcset.png");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/ignored-picture-source.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "src" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/unselected-picture.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/selected-fallback.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/print-picture-source.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/unsupported-picture-type.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "#icon");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/ignored-child-source.mp4");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/unsupported-child-source.mp4");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "https://example.test/images/non-important-property.png");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/custom-property-used.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/cross-block-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/important-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/inherited-custom-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/dom-inherited-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/nearer-custom-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/nearest-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/spaced-important-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "https://example.test/images/later-root-property.png");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/rank-local.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/split-selector-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/specific-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "https://example.test/images/specific-property.png");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/where-ok.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/is-ok.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/inherited-fallback.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/element-fallback.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/element-fallback.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/cascaded-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/same-element-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/multi-property-a.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/multi-property-b.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/grouped-custom-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/alias-target.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/inline-inherited-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "style-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/inline-card-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "style-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/inline-ancestor-stylesheet-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/inline-self-stylesheet-property.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/inline-override.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/hover-var.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-var-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/escaped.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/escaped-function.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/custom-supports.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/hero.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-image-set" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/escaped-image-set.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-image-set" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/typed.avif" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-image-set");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/commented-type.avif" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-image-set");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "image/avif");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/hero-2x.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-url");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/hero.avif" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-image-set");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/refresh.html" && resource.ElementName == "meta" && resource.Kind == HtmlResourceKind.Hyperlink && resource.AttributeName == "content" && resource.DiagnosticCode == "HyperlinkRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/report;v=1.html" && resource.ElementName == "meta" && resource.Kind == HtmlResourceKind.Hyperlink && resource.AttributeName == "content");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/a/*v*/b.png" && resource.Kind == HtmlResourceKind.Image);
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/inline.png" && resource.Kind == HtmlResourceKind.Image);
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/app.js" && resource.Kind == HtmlResourceKind.Script);
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/favicon.png" && resource.Kind == HtmlResourceKind.Image);
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/preload.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "imagesrcset" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/spaced-submit.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "src" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/preload-large.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "imagesrcset");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/picture-source.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "srcset");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/selected-picture.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "srcset");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/lazy-only.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "data-srcset");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/source-url-only.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "src" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/source-url-fallback.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "src");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/policy-source.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "srcset" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/policy-fallback.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "src");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/blocked-wide.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "srcset" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/next-wide.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "srcset");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "https://example.test/images/unused-wide-fallback.png");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/mixed-picture.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "srcset" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/mixed-picture.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "srcset");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "https://example.test/images/unused-mixed-next.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "https://example.test/images/unused-mixed-fallback.png");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "https://example.test/images/word-skipped.avif");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/word-fallback.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "src");
        Assert.Contains(manifest.Resources, resource => resource.Source == "mailto:ops@example.test" && resource.Kind == HtmlResourceKind.Script && !resource.IsAllowed && resource.DiagnosticCode == "ScriptResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/svg-script.js" && resource.Kind == HtmlResourceKind.Script && resource.AttributeName == "href" && resource.DiagnosticCode == "ScriptResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "mailto:ops@example.test" && resource.Kind == HtmlResourceKind.Image && !resource.IsAllowed && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/vector.png" && resource.Kind == HtmlResourceKind.Image && resource.ElementName == "image");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/vector-xlink.png" && resource.Kind == HtmlResourceKind.Image && resource.ElementName == "image" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/filter-image.png" && resource.Kind == HtmlResourceKind.Image && resource.ElementName == "feimage" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/symbols.svg#icon" && resource.Kind == HtmlResourceKind.Image && resource.ElementName == "use" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/lazy-img.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "data-lazy-src" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/original-img.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "data-original-srcset" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/original-img.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "data-original-srcset");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/poster.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "poster" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/media/movie.mp4" && resource.Kind == HtmlResourceKind.Media && resource.AttributeName == "data-src");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/selected-child-source.mp4" && resource.Kind == HtmlResourceKind.Media && resource.AttributeName == "src" && resource.DiagnosticCode == "MediaResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/media/first-playable.mp4" && resource.Kind == HtmlResourceKind.Media && resource.AttributeName == "src");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/unselected-playable.mp4");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/blocked-video.mp4" && resource.Kind == HtmlResourceKind.Media && resource.AttributeName == "src" && resource.DiagnosticCode == "MediaResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/media/policy-fallback.mp4" && resource.Kind == HtmlResourceKind.Media && resource.AttributeName == "src");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/negated-active-feature.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "css-url" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/embed.pdf" && resource.Kind == HtmlResourceKind.Other && resource.ElementName == "embed" && resource.AttributeName == "src");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/body-bg.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "background" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/upload" && resource.Kind == HtmlResourceKind.Hyperlink && resource.AttributeName == "action");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/delete" && resource.Kind == HtmlResourceKind.Hyperlink && resource.AttributeName == "formaction");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/bogus-delete" && resource.Kind == HtmlResourceKind.Hyperlink && resource.AttributeName == "formaction");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/submit.png" && resource.Kind == HtmlResourceKind.Image && resource.AttributeName == "src" && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/lazy-fallback.png");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/mask.svg" && resource.Kind == HtmlResourceKind.Image && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/filter.svg#f" && resource.Kind == HtmlResourceKind.Image && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/clip.svg#c" && resource.Kind == HtmlResourceKind.Image && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "#blur");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "#shape");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "#internal-image");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "#inline-sprite");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/input-metadata");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/logo.png" && resource.Kind == HtmlResourceKind.Image && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/ignored-frame.html");
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "file:///secret/empty-srcdoc-frame.html");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/srcdoc.png" && resource.Kind == HtmlResourceKind.Image && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "file:///secret/srcdoc-content.png" && resource.Kind == HtmlResourceKind.Image && resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(manifest.Resources, resource => resource.Source == "https://example.test/images/bg.png" && resource.Kind == HtmlResourceKind.Image && resource.IsAllowed);
        Assert.DoesNotContain(manifest.Resources, resource => resource.Source == "https://example.test/images/bg.png" && resource.Kind == HtmlResourceKind.Font);

        string nestedSrcDoc = "<img src=\"file:///secret/srcdoc-too-deep.png\">";
        for (int i = 0; i < 9; i++) {
            nestedSrcDoc = "<iframe srcdoc=\"" + System.Net.WebUtility.HtmlEncode(nestedSrcDoc) + "\"></iframe>";
        }

        HtmlResourceManifest nestedManifest = HtmlResourcePipeline.BuildManifest(nestedSrcDoc, new HtmlResourcePipelineOptions {
            UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
        });
        Assert.DoesNotContain(nestedManifest.Resources, resource => resource.Source == "file:///secret/srcdoc-too-deep.png");
    }

    [Fact]
    public void HtmlEnginePlatform_NormalizerPreservesCssStringsAndSvgElementCasing() {
        string normalized = HtmlNormalizer.Normalize(
            "<main><style>@import \"file:///secret/theme.css\";@import url('file:///secret/import-url.css');@import/*x*/\"file:///secret/comment-import.css\";.label::before{content:\"url(file:///secret/literal.png)\"}.escaped{background-image:url(\\66 ile:///secret/escaped.png)}.escaped-fn{background-image:\\75rl(file:///secret/escaped-function.png)}.real{background-image:url(file:///secret/bg.png);background-image:image-set(\"file:///secret/hero.png\" 1x);background-image:\\69mage-set(\"file:///secret/escaped-hero.png\" 1x);background-image:image-set(\"https://example.test/images/typed.avif\" \\74ype(\"image/avif\"));background-image:image-set(\"https://example.test/images/commented-type.avif\" type(/*c*/\"image/avif\"))}</style><pre>line 1\n  line 2</pre><textarea>  draft\n text</textarea><input name=\"code\" value=\"  A  \"><svg viewBox=\"0 0 10 10\"><defs><linearGradient id=\"g\"></linearGradient><clipPath id=\"c\"></clipPath></defs></svg></main>",
            new HtmlNormalizationOptions {
                UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
            });

        Assert.Contains("content:\"url(file:///secret/literal.png)\"", normalized);
        Assert.Contains("background-image:url(\"\")", normalized);
        Assert.DoesNotContain("file:///secret/theme.css", normalized);
        Assert.DoesNotContain("file:///secret/import-url.css", normalized);
        Assert.DoesNotContain("file:///secret/comment-import.css", normalized);
        Assert.DoesNotContain("file:///secret/escaped.png", normalized);
        Assert.DoesNotContain("file:///secret/escaped-function.png", normalized);
        Assert.DoesNotContain("file:///secret/hero.png", normalized);
        Assert.DoesNotContain("file:///secret/escaped-hero.png", normalized);
        Assert.Contains("https://example.test/images/commented-type.avif", normalized);
        Assert.Contains("image/avif", normalized);
        Assert.Contains("line 1\n  line 2", normalized);
        Assert.Contains("  draft\n text", normalized);
        Assert.Contains("value=\"  A  \"", normalized);
        Assert.Contains("<linearGradient", normalized);
        Assert.Contains("</linearGradient>", normalized);
        Assert.Contains("<clipPath", normalized);
        Assert.DoesNotContain("<lineargradient", normalized, StringComparison.Ordinal);
        Assert.DoesNotContain("<clippath", normalized, StringComparison.Ordinal);

        string policyNormalized = HtmlNormalizer.Normalize(
            "<main><a href=\"mailto:ops@example.test\">Mail</a><img src=\"mailto:ops@example.test\"><a href=\"\">Current</a><form action=\"\"></form></main>",
            new HtmlNormalizationOptions {
                BaseUri = new Uri("file:///secret/page.html"),
                UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
            });
        Assert.Contains("href=\"mailto:ops@example.test\"", policyNormalized);
        Assert.DoesNotContain("src=\"mailto:ops@example.test\"", policyNormalized);
        Assert.DoesNotContain("file:///secret/page.html", policyNormalized);

        string lazyPolicyNormalized = HtmlNormalizer.Normalize(
            "<main><img data-lazy-src=\"file:///secret/lazy.png\" data-original-src=\"https://example.test/images/original.png\" data-lazy-srcset=\"file:///secret/lazy-1x.png 1x, https://example.test/images/lazy-2x.png 2x\"></main>",
            new HtmlNormalizationOptions {
                UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
            });
        Assert.DoesNotContain("file:///secret", lazyPolicyNormalized);
        Assert.Contains("data-original-src=\"https://example.test/images/original.png\"", lazyPolicyNormalized);
        Assert.Contains("data-lazy-srcset=\"https://example.test/images/lazy-2x.png 2x\"", lazyPolicyNormalized);

        string metadataNormalized = HtmlNormalizer.Normalize(
            "<main><div data-src=\"invoice-42\" data-original=\"batch-7\" data-lazy-srcset=\"not-a-resource\"></div><img data-src=\"file:///secret/lazy.png\" data-original=\"https://example.test/images/original.png\"></main>",
            new HtmlNormalizationOptions {
                BaseUri = new Uri("https://example.test/root/"),
                UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile()
            });
        Assert.Contains("data-src=\"invoice-42\"", metadataNormalized);
        Assert.Contains("data-original=\"batch-7\"", metadataNormalized);
        Assert.Contains("data-lazy-srcset=\"not-a-resource\"", metadataNormalized);
        Assert.DoesNotContain("file:///secret/lazy.png", metadataNormalized);
        Assert.Contains("data-original=\"https://example.test/images/original.png\"", metadataNormalized);

        string refreshNormalized = HtmlNormalizer.Normalize(
            "<html><head><meta http-equiv=\"refresh\" content=\"0; url=javascript:alert(1)\"><meta http-equiv=\"refresh\" content=\"5; url='next.html'\"></head><body></body></html>",
            new HtmlNormalizationOptions {
                BaseUri = new Uri("https://example.test/root/page.html"),
                UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
                UseBodyContentsOnly = false
            });
        Assert.Contains("content=\"0\"", refreshNormalized);
        Assert.Contains("content=\"5; url=https://example.test/root/next.html\"", refreshNormalized);
        Assert.DoesNotContain("javascript:", refreshNormalized, StringComparison.OrdinalIgnoreCase);
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
        HtmlLogicalDocument imageMap = HtmlLogicalDocumentBuilder.FromHtml("<main><map name=\"m\"><area href=\"https://example.test/a\" alt=\"A\"></map></main>");
        Assert.Equal(1, imageMap.Count(HtmlLogicalNodeKind.Link));
        HtmlLogicalDocument definitionList = HtmlLogicalDocumentBuilder.FromHtml("<main><dl><dt>Term</dt><dd>Definition</dd></dl></main>");
        Assert.Equal(1, definitionList.Count(HtmlLogicalNodeKind.List));
        Assert.Equal(2, definitionList.Count(HtmlLogicalNodeKind.ListItem));
        HtmlLogicalDocument noscriptFallback = HtmlLogicalDocumentBuilder.FromHtml("<main><noscript><p>Enable billing exports manually</p></noscript></main>");
        Assert.Equal(1, noscriptFallback.Count(HtmlLogicalNodeKind.Paragraph));
        HtmlRoundTripScore noscriptFallbackScore = HtmlRoundTripScorer.Compare(
            "<main><noscript><p>Enable billing exports manually</p></noscript></main>",
            "<main></main>");
        Assert.InRange(noscriptFallbackScore.Metrics["text"], 0D, 0.99D);

        HtmlRoundTripScore repeatedTextScore = HtmlRoundTripScorer.Compare(
            "<main><p>" + new string('a', 100) + "</p></main>",
            "<main><p>" + new string('a', 32) + "</p></main>");
        Assert.InRange(repeatedTextScore.Metrics["text"], 0D, 0.99D);

        HtmlRoundTripScore hiddenTextScore = HtmlRoundTripScorer.Compare(
            "<main><p>Visible <span hidden>draft</span><span aria-hidden=\"true\">internal</span></p></main>",
            "<main><p>Visible</p></main>");
        Assert.Equal(1D, hiddenTextScore.Metrics["text"], 3);
        Assert.Equal(1D, hiddenTextScore.Metrics["nodes"], 3);
        Assert.Equal(1D, hiddenTextScore.Metrics["paragraphs"], 3);

        HtmlRoundTripScore formDefaultScore = HtmlRoundTripScorer.Compare(
            "<main><form><input name=\"q\"></form></main>",
            "<main><form method=\"get\" enctype=\"application/x-www-form-urlencoded\"><input name=\"q\"></form></main>");
        Assert.Equal(1D, formDefaultScore.Metrics["form-state"], 3);

        HtmlRoundTripScore emptyActionScore = HtmlRoundTripScorer.Compare(
            "<main><form action=\"\"><input name=\"q\"></form></main>",
            "<main><form><input name=\"q\"></form></main>");
        Assert.Equal(1D, emptyActionScore.Metrics["form-state"], 3);

        HtmlRoundTripScore stylesheetHiddenTextScore = HtmlRoundTripScorer.Compare(
            "<main><style>.draft{display:none}.private{visibility:hidden}</style><p>Visible <span class=\"draft\">draft</span><span class=\"private\">internal</span></p></main>",
            "<main><p>Visible</p></main>");
        Assert.Equal(1D, stylesheetHiddenTextScore.Metrics["text"], 3);

        HtmlRoundTripScore hiddenParagraphScore = HtmlRoundTripScorer.Compare(
            "<main><style>.draft{display:none}</style><p>Visible</p><p class=\"draft\">Internal</p></main>",
            "<main><p>Visible</p></main>");
        Assert.Equal(1D, hiddenParagraphScore.Metrics["nodes"], 3);
        Assert.Equal(1D, hiddenParagraphScore.Metrics["paragraphs"], 3);

        HtmlRoundTripScore hiddenControlOwnerScore = HtmlRoundTripScorer.Compare(
            "<main><style>.draft{display:none}</style><form id=\"f\"><input class=\"draft\" form=\"missing\" name=\"q\"></form></main>",
            "<main><form id=\"f\"></form></main>");
        Assert.Equal(1D, hiddenControlOwnerScore.Metrics["form-state"], 3);

        HtmlRoundTripScore visibilityHiddenParagraphScore = HtmlRoundTripScorer.Compare(
            "<main><style>.private{visibility:hidden}</style><p>Visible</p><p class=\"private\">Internal</p></main>",
            "<main><p>Visible</p></main>");
        Assert.Equal(1D, visibilityHiddenParagraphScore.Metrics["nodes"], 3);
        Assert.Equal(1D, visibilityHiddenParagraphScore.Metrics["paragraphs"], 3);

        HtmlRoundTripScore inlineCascadeVisibleTextScore = HtmlRoundTripScorer.Compare(
            "<main><span style=\"display:none; display:inline\">Visible</span></main>",
            "<main></main>");
        Assert.InRange(inlineCascadeVisibleTextScore.Metrics["text"], 0D, 0.99D);

        HtmlRoundTripScore visibilityDescendantTextScore = HtmlRoundTripScorer.Compare(
            "<main><div style=\"visibility:hidden\"><span style=\"visibility:visible\">Visible</span></div></main>",
            "<main></main>");
        Assert.InRange(visibilityDescendantTextScore.Metrics["text"], 0D, 0.99D);

        HtmlRoundTripScore visibilityInitialDescendantTextScore = HtmlRoundTripScorer.Compare(
            "<main><div style=\"visibility:hidden\"><span style=\"visibility:initial\">Visible</span></div></main>",
            "<main></main>");
        Assert.InRange(visibilityInitialDescendantTextScore.Metrics["text"], 0D, 0.99D);

        HtmlRoundTripScore visibilityCollapseTextScore = HtmlRoundTripScorer.Compare(
            "<main><span style=\"visibility:collapse\">hidden</span></main>",
            "<main></main>");
        Assert.Equal(1D, visibilityCollapseTextScore.Metrics["text"], 3);

        HtmlRoundTripScore pictureSourceScore = HtmlRoundTripScorer.Compare(
            "<main><picture><source srcset=\"wide.png\"><img src=\"small.png\"></picture></main>",
            "<main><img src=\"small.png\"></main>");
        Assert.InRange(pictureSourceScore.Metrics["images"], 0D, 0.99D);
        Assert.InRange(pictureSourceScore.Metrics["image-sources"], 0D, 0.99D);

        HtmlRoundTripScore pictureSelectionScore = HtmlRoundTripScorer.Compare(
            "<main><picture><source media=\"(min-width: 800px)\" type=\"image/avif\" sizes=\"100vw\" srcset=\"wide.avif\"><img src=\"small.png\"></picture></main>",
            "<main><picture><source media=\"print\" type=\"image/webp\" sizes=\"50vw\" srcset=\"wide.avif\"><img src=\"small.png\"></picture></main>");
        Assert.Equal(1D, pictureSelectionScore.Metrics["images"], 3);
        Assert.InRange(pictureSelectionScore.Metrics["image-sources"], 0D, 0.99D);

        HtmlRoundTripScore pictureSourceSrcScore = HtmlRoundTripScorer.Compare(
            "<main><picture><source src=\"unused.png\" data-src=\"unused-data.png\" srcset=\"wide.png\"><img src=\"small.png\"></picture></main>",
            "<main><picture><source srcset=\"wide.png\"><img src=\"small.png\"></picture></main>");
        Assert.InRange(pictureSourceSrcScore.Metrics["image-sources"], 0D, 0.99D);

        HtmlRoundTripScore figcaptionWrapperScore = HtmlRoundTripScorer.Compare(
            "<main><figure><img src=\"chart.png\"><figcaption>Quarterly result</figcaption></figure></main>",
            "<main><figure><img src=\"chart.png\"><p>Quarterly result</p></figure></main>");
        Assert.Equal(1D, figcaptionWrapperScore.Metrics["figures"], 3);
        Assert.Equal(1D, figcaptionWrapperScore.Metrics["figure-signatures"], 3);

        HtmlRoundTripScore mediaScore = HtmlRoundTripScorer.Compare(
            "<main><video src=\"movie.mp4\"><source src=\"movie-hd.mp4\"></video></main>",
            "<main></main>");
        Assert.InRange(mediaScore.Metrics["media"], 0D, 0.99D);
        Assert.InRange(mediaScore.Metrics["media-sources"], 0D, 0.99D);

        HtmlRoundTripScore mediaPlaybackScore = HtmlRoundTripScorer.Compare(
            "<main><video src=\"movie.mp4\" controls preload=\"metadata\"><track src=\"captions.vtt\" label=\"English\" default></video></main>",
            "<main><video src=\"movie.mp4\"><track src=\"captions.vtt\"></video></main>");
        Assert.Equal(1D, mediaPlaybackScore.Metrics["media"], 3);
        Assert.InRange(mediaPlaybackScore.Metrics["media-sources"], 0D, 0.99D);

        HtmlRoundTripScore mediaBooleanScore = HtmlRoundTripScorer.Compare(
            "<main><video src=\"movie.mp4\" controls autoplay loop muted><track src=\"captions.vtt\" default></video></main>",
            "<main><video src=\"movie.mp4\" controls=\"controls\" autoplay=\"autoplay\" loop=\"loop\" muted=\"muted\"><track src=\"captions.vtt\" default=\"default\"></video></main>");
        Assert.Equal(1D, mediaBooleanScore.Metrics["media-sources"], 3);

        HtmlRoundTripScore inertMediaSrcSetScore = HtmlRoundTripScorer.Compare(
            "<main><video><source src=\"movie.mp4\" srcset=\"wide.mp4 2x\"></video></main>",
            "<main><video><source src=\"movie.mp4\"></video></main>");
        Assert.Equal(1D, inertMediaSrcSetScore.Metrics["media-sources"], 3);

        HtmlRoundTripScore resolvedMediaScore = HtmlRoundTripScorer.Compare(
            "<html><head><base href=\"https://example.test/media/\"></head><body><main><video src=\"movie.mp4\" poster=\"poster.png\"><track src=\"captions.vtt\"></video></main></body></html>",
            "<main><video src=\"https://example.test/media/movie.mp4\" poster=\"https://example.test/media/poster.png\"><track src=\"https://example.test/media/captions.vtt\"></video></main>");
        Assert.Equal(1D, resolvedMediaScore.Metrics["media-sources"], 3);

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

        HtmlRoundTripScore definitionListScore = HtmlRoundTripScorer.Compare(
            "<main><dl><dt>Term</dt><dd>Definition</dd></dl></main>",
            "<main><p>Term</p><p>Definition</p></main>");
        Assert.InRange(definitionListScore.Metrics["lists"], 0D, 0.99D);
        Assert.InRange(definitionListScore.Metrics["list-items"], 0D, 0.99D);

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

        HtmlRoundTripScore formOwnerScore = HtmlRoundTripScorer.Compare(
            "<main><form id=\"save\" action=\"/save\"></form><form id=\"delete\" action=\"/delete\"></form><input form=\"save\" name=\"x\"><input form=\"delete\" name=\"y\"></main>",
            "<main><form id=\"save\" action=\"/save\"></form><form id=\"delete\" action=\"/delete\"></form><input form=\"delete\" name=\"x\"><input form=\"save\" name=\"y\"></main>");
        Assert.Equal(1D, formOwnerScore.Metrics["forms"], 3);
        Assert.InRange(formOwnerScore.Metrics["form-state"], 0D, 0.99D);

        HtmlRoundTripScore unresolvedFormOwnerScore = HtmlRoundTripScorer.Compare(
            "<main><input form=\"missing\" name=\"q\"></main>",
            "<main><input name=\"q\"></main>");
        Assert.InRange(unresolvedFormOwnerScore.Metrics["form-state"], 0D, 0.99D);

        HtmlRoundTripScore invalidExplicitOwnerInsideFormScore = HtmlRoundTripScorer.Compare(
            "<main><form id=\"f\"><input form=\"missing\" name=\"q\"></form></main>",
            "<main><form id=\"f\"><input name=\"q\"></form></main>");
        Assert.InRange(invalidExplicitOwnerInsideFormScore.Metrics["form-state"], 0D, 0.99D);

        HtmlRoundTripScore fieldsetDisabledScore = HtmlRoundTripScorer.Compare(
            "<main><form><fieldset disabled><input name=\"x\"></fieldset></form></main>",
            "<main><form><fieldset><input name=\"x\"></fieldset></form></main>");
        Assert.Equal(1D, fieldsetDisabledScore.Metrics["forms"], 3);
        Assert.InRange(fieldsetDisabledScore.Metrics["form-state"], 0D, 0.99D);

        HtmlRoundTripScore fieldsetLegendScore = HtmlRoundTripScorer.Compare(
            "<main><form><fieldset disabled><legend><input name=\"title\"></legend><input name=\"x\"></fieldset></form></main>",
            "<main><form><fieldset disabled><legend><input name=\"title\" data-fieldset-disabled=\"true\"></legend><input name=\"x\"></fieldset></form></main>");
        Assert.Equal(1D, fieldsetLegendScore.Metrics["forms"], 3);
        Assert.InRange(fieldsetLegendScore.Metrics["form-state"], 0D, 0.99D);

        HtmlRoundTripScore booleanAttributeScore = HtmlRoundTripScorer.Compare(
            "<main><form><input type=\"checkbox\" checked></form></main>",
            "<main><form><input type=\"checkbox\" checked=\"checked\"></form></main>");
        Assert.Equal(1D, booleanAttributeScore.Metrics["form-state"], 3);

        HtmlRoundTripScore defaultTypeScore = HtmlRoundTripScorer.Compare(
            "<main><form><input name=\"x\"><button>Save</button></form></main>",
            "<main><form><input type=\"text\" name=\"x\"><button type=\"submit\">Save</button></form></main>");
        Assert.Equal(1D, defaultTypeScore.Metrics["form-state"], 3);

        HtmlRoundTripScore invalidTypeScore = HtmlRoundTripScorer.Compare(
            "<main><form><input type=\"bogus\" name=\"x\"><button type=\"bogus\">Save</button></form></main>",
            "<main><form><input type=\"text\" name=\"x\"><button type=\"submit\">Save</button></form></main>");
        Assert.Equal(1D, invalidTypeScore.Metrics["form-state"], 3);

        HtmlRoundTripScore checkboxDefaultValueScore = HtmlRoundTripScorer.Compare(
            "<main><form><input type=\"checkbox\" name=\"agree\" checked><input type=\"radio\" name=\"tier\" checked></form></main>",
            "<main><form><input type=\"checkbox\" name=\"agree\" value=\"on\" checked><input type=\"radio\" name=\"tier\" value=\"on\" checked></form></main>");
        Assert.Equal(1D, checkboxDefaultValueScore.Metrics["form-state"], 3);

        HtmlRoundTripScore implicitOptionValueScore = HtmlRoundTripScorer.Compare(
            "<main><form><select><option selected>Gold</option></select></form></main>",
            "<main><form><select><option value=\"Gold\" selected>Gold</option></select></form></main>");
        Assert.Equal(1D, implicitOptionValueScore.Metrics["form-state"], 3);

        HtmlRoundTripScore implicitSelectedOptionScore = HtmlRoundTripScorer.Compare(
            "<main><form><select><option>Gold</option></select></form></main>",
            "<main><form><select><option selected>Gold</option></select></form></main>");
        Assert.Equal(1D, implicitSelectedOptionScore.Metrics["form-state"], 3);

        HtmlRoundTripScore inertOptionFormScore = HtmlRoundTripScorer.Compare(
            "<main><form id=\"f\"></form><select><option form=\"f\">Gold</option></select></main>",
            "<main><form id=\"f\"></form><select><option>Gold</option></select></main>");
        Assert.Equal(1D, inertOptionFormScore.Metrics["form-state"], 3);

        HtmlRoundTripScore resolvedFormOwnerScore = HtmlRoundTripScorer.Compare(
            "<html><head><base href=\"https://example.test/\"></head><body><main><form action=\"save\"><input name=\"x\"></form></main></body></html>",
            "<main><form action=\"https://example.test/save\"><input name=\"x\"></form></main>");
        Assert.Equal(1D, resolvedFormOwnerScore.Metrics["form-state"], 3);

        HtmlRoundTripScore requiredFormScore = HtmlRoundTripScorer.Compare(
            "<main><form><input name=\"email\" required pattern=\".+@.+\"></form></main>",
            "<main><form><input name=\"email\"></form></main>");
        Assert.Equal(1D, requiredFormScore.Metrics["forms"], 3);
        Assert.InRange(requiredFormScore.Metrics["form-state"], 0D, 0.99D);

        HtmlRoundTripScore imageSubmitterScore = HtmlRoundTripScorer.Compare(
            "<main><form><input type=\"image\" src=\"save-a.png\" alt=\"Save\"></form></main>",
            "<main><form><input type=\"image\" src=\"save-b.png\" alt=\"Save\"></form></main>");
        Assert.Equal(1D, imageSubmitterScore.Metrics["forms"], 3);
        Assert.InRange(imageSubmitterScore.Metrics["form-state"], 0D, 0.99D);

        HtmlRoundTripScore submitterOverrideScore = HtmlRoundTripScorer.Compare(
            "<main><form><button type=\"submit\" formmethod=\"get\" formenctype=\"text/plain\" formtarget=\"_blank\" formnovalidate>Go</button></form></main>",
            "<main><form><button type=\"submit\">Go</button></form></main>");
        Assert.Equal(1D, submitterOverrideScore.Metrics["forms"], 3);
        Assert.InRange(submitterOverrideScore.Metrics["form-state"], 0D, 0.99D);

        HtmlRoundTripScore inertSubmitterOverrideScore = HtmlRoundTripScorer.Compare(
            "<main><form><input type=\"text\" name=\"x\" formaction=\"/save\" formmethod=\"post\"></form></main>",
            "<main><form><input type=\"text\" name=\"x\"></form></main>");
        Assert.Equal(1D, inertSubmitterOverrideScore.Metrics["forms"], 3);
        Assert.Equal(1D, inertSubmitterOverrideScore.Metrics["form-state"], 3);

        HtmlRoundTripScore inertImageSubmitterAttributeScore = HtmlRoundTripScorer.Compare(
            "<main><form><input type=\"text\" name=\"x\" src=\"unused.png\" data-src=\"unused-data.png\" alt=\"Unused\"></form></main>",
            "<main><form><input type=\"text\" name=\"x\"></form></main>");
        Assert.Equal(1D, inertImageSubmitterAttributeScore.Metrics["forms"], 3);
        Assert.Equal(1D, inertImageSubmitterAttributeScore.Metrics["form-state"], 3);

        HtmlRoundTripScore linkScore = HtmlRoundTripScorer.Compare(
            "<main><a href=\"https://example.test/a\">same text</a></main>",
            "<main><a href=\"https://example.test/b\">same text</a></main>");
        Assert.Equal(1D, linkScore.Metrics["links"], 3);
        Assert.InRange(linkScore.Metrics["link-targets"], 0D, 0.99D);

        HtmlRoundTripScore linkBrowsingScore = HtmlRoundTripScorer.Compare(
            "<main><a href=\"https://example.test/report\" target=\"_blank\" rel=\"noopener\" download>Report</a></main>",
            "<main><a href=\"https://example.test/report\">Report</a></main>");
        Assert.Equal(1D, linkBrowsingScore.Metrics["links"], 3);
        Assert.InRange(linkBrowsingScore.Metrics["link-targets"], 0D, 0.99D);

        HtmlRoundTripScore relTokenScore = HtmlRoundTripScorer.Compare(
            "<main><a href=\"https://example.test/report\" rel=\"noopener noreferrer\">Report</a></main>",
            "<main><a href=\"https://example.test/report\" rel=\"NOREferrer noopener\">Report</a></main>");
        Assert.Equal(1D, relTokenScore.Metrics["link-targets"], 3);

        HtmlRoundTripScore downloadFilenameScore = HtmlRoundTripScorer.Compare(
            "<main><a href=\"https://example.test/report\" download=\"q1.pdf\">Report</a></main>",
            "<main><a href=\"https://example.test/report\" download=\"q2.pdf\">Report</a></main>");
        Assert.Equal(1D, downloadFilenameScore.Metrics["links"], 3);
        Assert.InRange(downloadFilenameScore.Metrics["link-targets"], 0D, 0.99D);

        HtmlRoundTripScore resolvedLinkScore = HtmlRoundTripScorer.Compare(
            "<html><head><base href=\"https://example.test/docs/\"></head><body><main><a href=\"page.html\">Docs</a></main></body></html>",
            "<main><a href=\"https://example.test/docs/page.html\">Docs</a></main>");
        Assert.Equal(1D, resolvedLinkScore.Metrics["link-targets"], 3);

        HtmlRoundTripScore areaLinkScore = HtmlRoundTripScorer.Compare(
            "<main><map name=\"chart\"><area href=\"https://example.test/a\" alt=\"A\"></map></main>",
            "<main><map name=\"chart\"><area href=\"https://example.test/b\" alt=\"A\"></map></main>");
        Assert.Equal(1D, areaLinkScore.Metrics["links"], 3);
        Assert.InRange(areaLinkScore.Metrics["link-targets"], 0D, 0.99D);

        HtmlRoundTripScore areaGeometryScore = HtmlRoundTripScorer.Compare(
            "<main><map name=\"chart\"><area href=\"https://example.test/a\" shape=\"rect\" coords=\"0,0,10,10\" alt=\"A\"></map></main>",
            "<main><map name=\"chart\"><area href=\"https://example.test/a\" shape=\"circle\" coords=\"5,5,5\" alt=\"A\"></map></main>");
        Assert.Equal(1D, areaGeometryScore.Metrics["links"], 3);
        Assert.InRange(areaGeometryScore.Metrics["link-targets"], 0D, 0.99D);

        HtmlRoundTripScore imageScore = HtmlRoundTripScorer.Compare(
            "<main><img src=\"https://example.test/a.png\" alt=\"Chart A\"></main>",
            "<main><img src=\"https://example.test/b.png\" alt=\"Chart B\"></main>");
        Assert.Equal(1D, imageScore.Metrics["images"], 3);
        Assert.InRange(imageScore.Metrics["image-sources"], 0D, 0.99D);

        HtmlRoundTripScore lazyImageScore = HtmlRoundTripScorer.Compare(
            "<main><img data-original=\"https://example.test/chart.png\" data-lazy-srcset=\"https://example.test/chart.png 1x, https://example.test/chart@2x.png 2x\" alt=\"Chart\"></main>",
            "<main><img alt=\"Chart\"></main>");
        Assert.Equal(1D, lazyImageScore.Metrics["images"], 3);
        Assert.InRange(lazyImageScore.Metrics["image-sources"], 0D, 0.99D);

        HtmlRoundTripScore resolvedImageScore = HtmlRoundTripScorer.Compare(
            "<html><head><base href=\"https://example.test/dir/\"></head><body><main><img src=\"chart.png\"></main></body></html>",
            "<main><img src=\"https://example.test/dir/chart.png\"></main>");
        Assert.Equal(1D, resolvedImageScore.Metrics["image-sources"], 3);

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
