using OfficeIMO.Reader.Benchmarks.Comparison;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderComparisonEvidenceTests {
    [Fact]
    public void Corpus_CoversThePlannedFormatAndMalformedInputContracts() {
        IReadOnlyList<ReaderComparisonCase> cases = ReaderComparisonCorpus.Create();

        Assert.Equal(
            new[] { "csv", "docx", "epub", "html", "malformed-pdf", "msg", "pdf", "pptx", "xlsx", "zip" },
            cases.Select(item => item.Id).OrderBy(value => value, StringComparer.Ordinal).ToArray());
        Assert.All(cases, item => Assert.NotEmpty(item.Bytes));
        Assert.Contains(cases.SelectMany(item => item.Probes), probe => probe.Kind == ReaderComparisonProbeKind.RichTable);
        Assert.Contains(cases.SelectMany(item => item.Probes), probe => probe.Kind == ReaderComparisonProbeKind.RichAsset);
        Assert.Contains(cases.SelectMany(item => item.Probes), probe => probe.Kind == ReaderComparisonProbeKind.LocationPage);
        Assert.Contains(cases.SelectMany(item => item.Probes), probe => probe.Kind == ReaderComparisonProbeKind.RejectsMalformedInput);
    }

    [Fact]
    public void Scorer_DoesNotPenalizeMarkdownOnlyToolsForRichResultProbes() {
        var probes = new[] {
            new ReaderComparisonProbe("text", ReaderComparisonProbeKind.ContainsText, "retained"),
            new ReaderComparisonProbe("table", ReaderComparisonProbeKind.RichTable),
            new ReaderComparisonProbe("location", ReaderComparisonProbeKind.LocationPage)
        };

        IReadOnlyList<ReaderComparisonProbeResult> results = ReaderComparisonScorer.ScoreMarkdown(
            "retained",
            probes,
            rejected: false);

        Assert.True(results[0].Applied);
        Assert.True(results[0].Passed);
        Assert.False(results[1].Applied);
        Assert.False(results[2].Applied);
    }

    [Fact]
    public void OfficeIMOComparison_ProducesScoredDeterministicResultsForEveryCase() {
        string output = Path.Combine(Path.GetTempPath(), "officeimo-reader-comparison-tests-" + Guid.NewGuid().ToString("N"));
        try {
            IReadOnlyList<ReaderComparisonCase> cases = ReaderComparisonCorpus.Create();

            ReaderComparisonToolResult result = ReaderComparisonCommand.RunOfficeIMO(cases, output);

            Assert.Equal(cases.Count, result.Cases.Count);
            Assert.All(result.Cases, item => Assert.Equal("success", item.Status));
            Assert.All(result.Cases, item => Assert.True(item.Deterministic));
            Assert.All(result.Cases, item => Assert.True(item.AppliedProbes > 0));
        } finally {
            if (Directory.Exists(output)) Directory.Delete(output, recursive: true);
        }
    }
}