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
        Assert.All(
            cases.Where(item => item.Id is "docx" or "pptx" or "xlsx"),
            item => Assert.Contains("2026-07-15T12:00:00.0000000Z", ReadPackageCoreProperties(item.Bytes), StringComparison.Ordinal));
        IReadOnlyDictionary<string, byte[]> repeatedPackages = ReaderComparisonCorpus.Create()
            .Where(item => item.Id is "docx" or "pptx" or "xlsx" or "epub" or "zip")
            .ToDictionary(item => item.Id, item => item.Bytes, StringComparer.Ordinal);
        Assert.All(
            cases.Where(item => item.Id is "docx" or "pptx" or "xlsx" or "epub" or "zip"),
            item => Assert.Equal(item.Bytes, repeatedPackages[item.Id]));

        foreach (ReaderComparisonCase item in cases.Where(item => item.Id is "epub" or "zip")) {
            using var stream = new MemoryStream(item.Bytes);
            using var archive = new System.IO.Compression.ZipArchive(
                stream,
                System.IO.Compression.ZipArchiveMode.Read);
            Assert.All(
                archive.Entries,
                entry => Assert.Equal(
                    new DateTime(2026, 7, 15, 12, 0, 0),
                    entry.LastWriteTime.DateTime));
        }

        ReaderComparisonCase epub = Assert.Single(cases, item => item.Id == "epub");
        using (var stream = new MemoryStream(epub.Bytes))
        using (var archive = new System.IO.Compression.ZipArchive(
            stream,
            System.IO.Compression.ZipArchiveMode.Read)) {
            string package = ReadPackageEntry(archive, "OEBPS/content.opf");
            Assert.Contains("unique-identifier=\"book-id\"", package, StringComparison.Ordinal);
            Assert.Contains("<dc:identifier id=\"book-id\">", package, StringComparison.Ordinal);
            Assert.Contains("<dc:language>en</dc:language>", package, StringComparison.Ordinal);
            Assert.Contains("properties=\"nav\"", package, StringComparison.Ordinal);
            Assert.Contains("epub:type=\"toc\"", ReadPackageEntry(archive, "OEBPS/nav.xhtml"), StringComparison.Ordinal);
        }
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
    public void Scorer_RequiresNonEmptyExpectedLinkAndImageTargets() {
        var probes = new[] {
            new ReaderComparisonProbe(
                "link",
                ReaderComparisonProbeKind.MarkdownLink,
                "Policy",
                "https://example.com/policy"),
            new ReaderComparisonProbe(
                "image",
                ReaderComparisonProbeKind.MarkdownImage,
                "Diagram",
                "data:image/png;base64,")
        };

        IReadOnlyList<ReaderComparisonProbeResult> empty = ReaderComparisonScorer.ScoreMarkdown(
            "[Policy]()\n![Diagram]()",
            probes,
            rejected: false);
        IReadOnlyList<ReaderComparisonProbeResult> wrong = ReaderComparisonScorer.ScoreMarkdown(
            "[Policy](https://example.com/wrong)\n![Diagram](image.png)",
            probes,
            rejected: false);
        IReadOnlyList<ReaderComparisonProbeResult> retained = ReaderComparisonScorer.ScoreMarkdown(
            "[Policy](https://example.com/policy)\n![Diagram](data:image/png;base64,AAAA)",
            probes,
            rejected: false);

        Assert.All(empty, result => Assert.False(result.Passed));
        Assert.All(wrong, result => Assert.False(result.Passed));
        Assert.All(retained, result => Assert.True(result.Passed));
    }

    [Fact]
    public void Scorer_DoesNotTreatImageSyntaxAsARetainedLink() {
        var probe = new ReaderComparisonProbe(
            "link",
            ReaderComparisonProbeKind.MarkdownLink,
            "Policy",
            "https://example.com/policy");

        ReaderComparisonProbeResult imageOnly = Assert.Single(ReaderComparisonScorer.ScoreMarkdown(
            "![Policy](https://example.com/policy)",
            new[] { probe },
            rejected: false));
        ReaderComparisonProbeResult link = Assert.Single(ReaderComparisonScorer.ScoreMarkdown(
            "[Policy](https://example.com/policy)",
            new[] { probe },
            rejected: false));

        Assert.False(imageOnly.Passed);
        Assert.True(link.Passed);
    }

    [Fact]
    public void Scorer_RequiresTheExpectedPageLocation() {
        var probe = new ReaderComparisonProbe(
            "page-location",
            ReaderComparisonProbeKind.LocationPage,
            expectedPage: 2);
        var document = new OfficeDocumentReadResult {
            Chunks = new[] { new ReaderChunk { Location = new ReaderLocation { Page = 1 } } }
        };

        ReaderComparisonProbeResult wrongPage = Assert.Single(ReaderComparisonScorer.ScoreOfficeDocument(
            string.Empty,
            document,
            new[] { probe },
            rejected: false));
        document.Chunks = new[] { new ReaderChunk { Location = new ReaderLocation { Page = 2 } } };
        ReaderComparisonProbeResult expectedPage = Assert.Single(ReaderComparisonScorer.ScoreOfficeDocument(
            string.Empty,
            document,
            new[] { probe },
            rejected: false));

        Assert.False(wrongPage.Passed);
        Assert.True(expectedPage.Passed);
    }

    [Fact]
    public void Scorer_RequiresExpectedOfficeLocationValues() {
        ReaderComparisonProbe[] probes = {
            new ReaderComparisonProbe(
                "heading-location",
                ReaderComparisonProbeKind.LocationHeading,
                "Evidence policy"),
            new ReaderComparisonProbe(
                "sheet-location",
                ReaderComparisonProbeKind.LocationSheet,
                "Evidence"),
            new ReaderComparisonProbe(
                "slide-location",
                ReaderComparisonProbeKind.LocationSlide,
                expectedSlide: 1)
        };
        var document = new OfficeDocumentReadResult {
            Chunks = new[] {
                new ReaderChunk {
                    Location = new ReaderLocation {
                        HeadingPath = "Wrong heading",
                        Sheet = "Sheet1",
                        Slide = 2
                    }
                }
            }
        };

        IReadOnlyList<ReaderComparisonProbeResult> wrong = ReaderComparisonScorer.ScoreOfficeDocument(
            string.Empty,
            document,
            probes,
            rejected: false);
        document.Chunks = new[] {
            new ReaderChunk {
                Location = new ReaderLocation {
                    HeadingPath = "Evidence policy",
                    Sheet = "Evidence",
                    Slide = 1
                }
            }
        };
        IReadOnlyList<ReaderComparisonProbeResult> expected = ReaderComparisonScorer.ScoreOfficeDocument(
            string.Empty,
            document,
            probes,
            rejected: false);

        Assert.All(wrong, result => Assert.False(result.Passed));
        Assert.All(expected, result => Assert.True(result.Passed));
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
            Assert.False(result.Rejected);
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
        Assert.False(result.Rejected);
    }

    [Fact]
    public async Task Runner_DoesNotTreatTruncatedNonZeroOutputAsConverterRejection() {
        ReaderComparisonRunnerConfiguration configuration = DotNetRunner(
            "stdout",
            "exec",
            "officeimo-reader-missing-" + new string('x', 2048) + ".dll");
        configuration.MaxOutputBytes = 1024;

        ReaderComparisonProcessOutput result = await ReaderComparisonProcessRunner.RunAsync(
            configuration,
            inputPath: "unused",
            outputPath: "unused",
            CancellationToken.None);

        Assert.Equal("failed", result.Status);
        Assert.Contains("truncated", result.Error, StringComparison.OrdinalIgnoreCase);
        Assert.False(result.Rejected);
    }

    [Fact]
    public async Task FileRunner_DoesNotTreatMissingOutputAsConverterRejection() {
        string output = Path.Combine(
            Path.GetTempPath(),
            "officeimo-reader-missing-output-" + Guid.NewGuid().ToString("N") + ".md");
        try {
            ReaderComparisonProcessOutput result = await ReaderComparisonProcessRunner.RunAsync(
                DotNetRunner("file", "exec", "officeimo-reader-missing-assembly.dll"),
                inputPath: "unused",
                outputPath: output,
                CancellationToken.None);

            Assert.Equal("failed", result.Status);
            Assert.Contains("did not create", result.Error, StringComparison.Ordinal);
            Assert.False(result.Rejected);
        } finally {
            if (File.Exists(output)) File.Delete(output);
        }
    }

    [Fact]
    public async Task Runner_UsesOnlyACompletedNonZeroExitAsConverterRejection() {
        ReaderComparisonProcessOutput result = await ReaderComparisonProcessRunner.RunAsync(
            DotNetRunner("stdout", "exec", "officeimo-reader-missing-assembly.dll"),
            inputPath: "unused",
            outputPath: "unused",
            CancellationToken.None);

        Assert.Equal("failed", result.Status);
        Assert.True(result.Rejected);
    }

    [Fact]
    public async Task Runner_ClassifiesAMissingExecutableAsUnavailable() {
        var configuration = new ReaderComparisonRunnerConfiguration {
            Name = "missing-runner",
            FileName = "officeimo-reader-missing-runner-" + Guid.NewGuid().ToString("N"),
            OutputMode = "stdout",
            TimeoutSeconds = 30,
            MaxOutputBytes = 4096
        };

        ReaderComparisonProcessOutput result = await ReaderComparisonProcessRunner.RunAsync(
            configuration,
            inputPath: "unused",
            outputPath: "unused",
            CancellationToken.None);

        Assert.Equal("unavailable", result.Status);
        Assert.False(result.Rejected);
    }

    [Fact]
    public async Task Runner_ClassifiesANonExecutableUnixFileAsUnavailable() {
        if (OperatingSystem.IsWindows()) return;
        string executable = Path.Combine(
            Path.GetTempPath(),
            "officeimo-reader-non-executable-" + Guid.NewGuid().ToString("N"));
        await File.WriteAllTextAsync(executable, "#!/bin/sh\nexit 0\n");
        File.SetUnixFileMode(executable, UnixFileMode.UserRead | UnixFileMode.UserWrite);
        try {
            var configuration = new ReaderComparisonRunnerConfiguration {
                Name = "non-executable-runner",
                FileName = executable,
                OutputMode = "stdout",
                TimeoutSeconds = 30,
                MaxOutputBytes = 4096
            };

            ReaderComparisonProcessOutput result = await ReaderComparisonProcessRunner.RunAsync(
                configuration,
                inputPath: "unused",
                outputPath: "unused",
                CancellationToken.None);

            Assert.Equal("unavailable", result.Status);
            Assert.False(result.Rejected);
        } finally {
            if (File.Exists(executable)) File.Delete(executable);
        }
    }

    [Fact]
    public async Task Runner_TimesOutWhenACompletedWrapperLeavesPipeHoldingDescendants() {
        if (OperatingSystem.IsWindows()) return;
        var configuration = new ReaderComparisonRunnerConfiguration {
            Name = "pipe-holding-descendant",
            FileName = "/bin/sh",
            Arguments = new List<string> { "-c", "sleep 30 &" },
            OutputMode = "stdout",
            TimeoutSeconds = 1,
            MaxOutputBytes = 4096
        };

        ReaderComparisonProcessOutput result = await ReaderComparisonProcessRunner.RunAsync(
            configuration,
            inputPath: "unused",
            outputPath: "unused",
            CancellationToken.None);

        Assert.Equal("timed-out", result.Status);
        Assert.False(result.Rejected);
    }

    [Fact]
    public void RepeatRunnerFailure_IsPropagatedIntoTheCaseOutcome() {
        var first = new ReaderComparisonProcessOutput { Status = "success", Markdown = "first" };
        var second = new ReaderComparisonProcessOutput { Status = "timed-out", Error = "timeout" };

        (string status, string? error) = ReaderComparisonCommand.ResolveRepeatOutcome(first, second);

        Assert.Equal("timed-out", status);
        Assert.Contains("Repeat run timed-out", error, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("case.md", "case.repeat.md")]
    [InlineData("case", "case.repeat")]
    public void RepeatOutputPath_PreservesTheOutputExtension(string outputPath, string expected) {
        Assert.Equal(expected, ReaderComparisonCommand.GetRepeatOutputPath(outputPath));
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

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void ExpectedExternalRejection_MustBeStableAcrossRepeats(
        bool firstRejected,
        bool secondRejected) {
        var first = new ReaderComparisonProcessOutput {
            Status = firstRejected ? "failed" : "success",
            Rejected = firstRejected
        };
        var second = new ReaderComparisonProcessOutput {
            Status = secondRejected ? "failed" : "success",
            Rejected = secondRejected
        };

        (string status, string? error) = ReaderComparisonCommand.ResolveRepeatOutcome(
            first,
            second,
            expectsRejection: true);

        Assert.Equal("failed", status);
        Assert.Contains("did not preserve", error, StringComparison.Ordinal);
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

    private static string ReadPackageCoreProperties(byte[] packageBytes) {
        using var stream = new MemoryStream(packageBytes);
        using var archive = new System.IO.Compression.ZipArchive(stream, System.IO.Compression.ZipArchiveMode.Read);
        System.IO.Compression.ZipArchiveEntry entry = Assert.Single(
            archive.Entries,
            item => string.Equals(item.FullName, "docProps/core.xml", StringComparison.Ordinal));
        using StreamReader reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }

    private static string ReadPackageEntry(
        System.IO.Compression.ZipArchive archive,
        string path) {
        System.IO.Compression.ZipArchiveEntry entry = Assert.Single(
            archive.Entries,
            item => string.Equals(item.FullName, path, StringComparison.Ordinal));
        using StreamReader reader = new StreamReader(entry.Open());
        return reader.ReadToEnd();
    }
}
