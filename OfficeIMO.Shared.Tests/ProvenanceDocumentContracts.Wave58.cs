using System;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void ForeignMathMlIframeSrcdocIsNotAProvenanceCarrier() {
        string nested = "<link rel=\"c2pa-manifest\" href=\"data:application/c2pa;base64," +
            Convert.ToBase64String(CreateManifestStore()) + "\">";
        string html = "<math><iframe srcdoc=\"" + nested.Replace("&", "&amp;").Replace("\"", "&quot;") +
            "\"></iframe></math>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);

        Assert.Empty(report.Evidence);
    }

    [Fact]
    public void ForeignMathMlStyleIsNotAProvenanceCarrier() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<math><style>.box{background-image:url('" + dataUri + "')}</style><mtext class=\"box\">x</mtext></math>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);

        Assert.Empty(report.Evidence);
    }

    [Fact]
    public void SvgStyleRemainsAnActiveProvenanceCarrier() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<svg><style>.box{background-image:url('" + dataUri + "')}</style><rect class=\"box\"/></svg>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void SrcsetCandidateLimitCountsRejectedDescriptors() {
        IReadOnlyList<HtmlSrcSetCandidate> candidates = HtmlSrcSetParser.Parse(
            "first.png nope, second.png also-nope, third.png 1x",
            maxCandidates: 2);

        Assert.Empty(candidates);
    }

    [Fact]
    public void IgnoredSelectStartTagsDoNotConsumeTheDomEntryBudget() {
        string html = "<select>" + string.Concat(Enumerable.Repeat("<div>", 100)) + "</select>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(
            html,
            new OfficeProvenanceOptions { MaxContainerEntries = 4 });

        Assert.Empty(report.Evidence);
    }
}
