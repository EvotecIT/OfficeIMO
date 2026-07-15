using OfficeIMO.Reader;
using OfficeIMO.Reader.Benchmarks.Comparison;
using System.Threading;
using System.Threading.Tasks;
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
    public void Scorer_RequiresTheExpectedNestedSourcePathMarker() {
        var probe = new ReaderComparisonProbe(
            "nested-path",
            ReaderComparisonProbeKind.LocationPath,
            "docs/evidence.md");
        var document = new OfficeDocumentReadResult {
            Chunks = new[] {
                new ReaderChunk { Location = new ReaderLocation { Path = "evidence-archive.zip" } }
            }
        };

        ReaderComparisonProbeResult weakLocation = ReaderComparisonScorer.ScoreOfficeDocument(
            string.Empty,
            document,
            new[] { probe },
            rejected: false)[0];
        document.Chunks = new[] {
            new ReaderChunk { Location = new ReaderLocation { Path = "evidence-archive.zip::docs/evidence.md" } }
        };
        ReaderComparisonProbeResult nestedLocation = ReaderComparisonScorer.ScoreOfficeDocument(
            string.Empty,
            document,
            new[] { probe },
            rejected: false)[0];

        Assert.False(weakLocation.Passed);
        Assert.True(nestedLocation.Passed);
    }

    [Fact]
    public async Task FileRunner_RemovesStaleOutputAndRequiresFreshOutput() {
        string output = Path.Combine(Path.GetTempPath(), "officeimo-reader-runner-" + Guid.NewGuid().ToString("N") + ".md");
        await File.WriteAllTextAsync(output, "stale output");
        try {
            ReaderComparisonProcessOutput result = await ReaderComparisonProcessRunner.RunAsync(
                DotNetRunner("file", "--version"),
                inputPath: output + ".input",
                outputPath: output,
                CancellationToken.None);

            Assert.Equal("failed", result.Status);
            Assert.Contains("did not create", result.Error, StringComparison.Ordinal);
            Assert.False(File.Exists(output));
        } finally {
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public async Task Runner_RejectsTruncatedStdout() {
        ReaderComparisonRunnerConfiguration configuration = DotNetRunner("stdout", "--info");
        configuration.MaxOutputBytes = 1024;

        ReaderComparisonProcessOutput result = await ReaderComparisonProcessRunner.RunAsync(
            configuration,
            inputPath: "unused",
            outputPath: "unused",
            CancellationToken.None);

        Assert.Equal("failed", result.Status);
        Assert.Contains("truncated", result.Error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void RepeatRunnerFailure_IsPropagatedIntoTheCaseOutcome() {
        var first = new ReaderComparisonProcessOutput { Status = "success", Markdown = "first" };
        var second = new ReaderComparisonProcessOutput { Status = "timed-out", Error = "timeout" };

        (string status, string? error) = ReaderComparisonCommand.ResolveRepeatOutcome(first, second);

        Assert.Equal("timed-out", status);
        Assert.Contains("Repeat run timed-out", error, StringComparison.Ordinal);
    }

    [Fact]
    public void ExpectedExternalRejection_IsReportedAsSuccess() {
        var rejected = new ReaderComparisonProcessOutput {
            Status = "failed",
            Error = "invalid input",
            Rejected = true
        };

        (string status, string? error) = ReaderComparisonCommand.ResolveRepeatOutcome(
            rejected,
            rejected,
            expectsRejection: true);

        Assert.Equal("success", status);
        Assert.Null(error);
    }

    [Fact]
    public void UnavailableExternalRunner_IsNotAValidRejection() {
        var unavailable = new ReaderComparisonProcessOutput {
            Status = "unavailable",
            Rejected = true
        };

        Assert.Equal(
            "unavailable",
            ReaderComparisonCommand.ResolveExternalStatus(unavailable, expectsRejection: true));
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
            ReaderComparisonCaseResult msg = Assert.Single(result.Cases, item => item.CaseId == "msg");
            Assert.True(Assert.Single(msg.Probes, probe => probe.Id == "attachment-content").Passed);
        } finally {
            if (Directory.Exists(output)) Directory.Delete(output, recursive: true);
        }
    }

    private static ReaderComparisonRunnerConfiguration DotNetRunner(string outputMode, params string[] arguments) {
        return new ReaderComparisonRunnerConfiguration {
            Name = "dotnet-test-runner",
            FileName = Environment.GetEnvironmentVariable("DOTNET_HOST_PATH") ?? "dotnet",
            Arguments = arguments.ToList(),
            OutputMode = outputMode,
            TimeoutSeconds = 30,
            MaxOutputBytes = 4096
        };
    }
}
