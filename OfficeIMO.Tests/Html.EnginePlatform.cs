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
                <style>
                    body { color: #222; font-family: Aptos; }
                    article.report { color: #123456; }
                    .warning { color: #aa0000; }
                    p { color: #0000aa; }
                    p.force { color: #0000aa !important; }
                    a[href*=','] { font-weight: 700; }
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
                        <img src="/images/chart.png" alt="Revenue chart">
                    </picture>
                    <video poster="/media/poster.png">
                        <source src="/media/demo.mp4" type="video/mp4">
                    </video>
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
        Assert.Contains(resourceManifest.Resources, resource => resource.ElementName == "video" && resource.AttributeName == "poster" && resource.Kind == HtmlResourceKind.Media && resource.IsAllowed);
        Assert.Contains(resourceManifest.Resources, resource => resource.ElementName == "source" && resource.Kind == HtmlResourceKind.Media && resource.IsAllowed);
        Assert.Contains(resourceManifest.Resources, resource => resource.ElementName == "object" && resource.AttributeName == "data" && resource.DiagnosticCode == "HtmlResourceRejectedByPolicy");
        Assert.Contains(resourceManifest.Resources, resource => resource.DiagnosticCode == "ImageResourceRejectedByPolicy");
        Assert.Contains(resourceManifest.Resources, resource => resource.DiagnosticCode == "HyperlinkRejectedByPolicy");
        Assert.Contains(resourceManifest.Resources, resource => resource.DiagnosticCode == "ScriptResourceRejectedByPolicy");
        Assert.Contains(resourceManifest.Diagnostics.Diagnostics, diagnostic => diagnostic.Code == "HyperlinkRejectedByPolicy");

        HtmlRoundTripScore score = HtmlRoundTripScorer.Compare(sourceHtml, roundTripHtml);
        Assert.InRange(score.Score, 0.60D, 1.00D);
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
    }
}
