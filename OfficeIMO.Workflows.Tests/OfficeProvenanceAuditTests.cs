using System.Text;
using System.Text.Json;
using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows.Tests;

public sealed class OfficeProvenanceAuditTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-provenance-audit-" + Guid.NewGuid().ToString("N"));
    public OfficeProvenanceAuditTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, true);
    private string Write(string name, string text) { string path = Path.Combine(_root, name); Directory.CreateDirectory(Path.GetDirectoryName(path)!); File.WriteAllText(path, text, new UTF8Encoding(false)); return path; }
    [Fact]
    public async Task AssessmentReportsSkippedChecksAndExactInputDigest() {
        string path = Write("text.txt", "a\u202Eb");
        var request = new OfficeProvenanceWorkflowRequest { InputPath = path, Operation = OfficeProvenanceWorkflowOperation.Assess };
        request.Assessment.InspectTextIntegrity = false;
        var disabled = await new OfficeWorkflowRunner().RunProvenanceAsync(request);
        Assert.True(disabled.Succeeded, disabled.Summary);
        Assert.Equal(OfficeProvenanceCheckStatus.Disabled, disabled.Assessment!.TextIntegrityStatus);
        Assert.Equal(OfficeProvenanceCheckStatus.NotConfigured, disabled.Assessment.VerificationStatus);
        Assert.DoesNotContain("0 text-integrity", disabled.Summary);
        Assert.Equal(OfficeTextIntegrityReview.ComputeSha256(File.ReadAllBytes(path)), disabled.InputSha256);
        using var json = JsonDocument.Parse(OfficeProvenanceReportSerializer.Serialize(disabled));
        Assert.Equal("Disabled", json.RootElement.GetProperty("checks").GetProperty("textIntegrity").GetString());
        request.Assessment.InspectTextIntegrity = true;
        var completed = await new OfficeWorkflowRunner().RunProvenanceAsync(request);
        Assert.True(OfficeProvenanceAudit.HasFindings(completed));
        Assert.Equal(OfficeProvenanceCheckStatus.Completed, completed.Assessment!.TextIntegrityStatus);
    }
    [Theory]
    [InlineData("text")]
    [InlineData("verifier")]
    [InlineData("detector")]
    [InlineData("cancel")]
    public async Task FailedAssessmentRetainsAttemptedCheckCoverage(string failure) {
        string path = Write("text.txt", failure == "text" ? new string('\u202E', 4097) : "ok");
        using var cancellation = new CancellationTokenSource();
        var runner = new OfficeWorkflowRunner(failure is "verifier" or "cancel" ? new ThrowingVerifier(cancellation, failure == "cancel") : null,
            failure == "detector" ? [new ThrowingDetector()] : null);
        var result = await runner.RunProvenanceAsync(new() { InputPath = path, Operation = OfficeProvenanceWorkflowOperation.Assess }, cancellationToken: cancellation.Token);
        Assert.False(result.Succeeded);
        Assert.Null(result.Assessment);
        Assert.Equal(OfficeProvenanceCheckStatus.Completed, result.Checks.Structural);
        Assert.Equal(failure == "text" ? OfficeProvenanceCheckStatus.Failed : OfficeProvenanceCheckStatus.Completed, result.Checks.TextIntegrity);
        Assert.Equal(failure is "verifier" or "cancel" ? OfficeProvenanceCheckStatus.Failed : OfficeProvenanceCheckStatus.NotConfigured, result.Checks.Verification);
        Assert.Equal(failure == "detector" ? OfficeProvenanceCheckStatus.Failed : OfficeProvenanceCheckStatus.NotConfigured, result.Checks.ProviderSignals);
        using var json = JsonDocument.Parse(OfficeProvenanceReportSerializer.Serialize(result));
        Assert.Equal(result.Checks.TextIntegrity.ToString(), json.RootElement.GetProperty("checks").GetProperty("textIntegrity").GetString());
        using var sarif = JsonDocument.Parse(OfficeProvenanceSarif.Serialize([result]));
        var run = sarif.RootElement.GetProperty("runs")[0];
        Assert.Equal("Completed", run.GetProperty("artifacts")[0].GetProperty("properties").GetProperty("structuralStatus").GetString());
        Assert.Equal(result.Checks.Verification.ToString(), run.GetProperty("artifacts")[0].GetProperty("properties").GetProperty("verificationStatus").GetString());
        Assert.False(run.GetProperty("invocations")[0].GetProperty("executionSuccessful").GetBoolean());
    }
    private sealed class ThrowingVerifier(CancellationTokenSource cancellation, bool cancel) : IOfficeProvenanceVerifier {
        public string Name => "test";
        public OfficeProvenanceVerificationResult Verify(string path, OfficeProvenanceVerificationOptions? options = null) {
            if (cancel) { cancellation.Cancel(); cancellation.Token.ThrowIfCancellationRequested(); }
            throw new InvalidOperationException("Verifier failure");
        }
    }
    private sealed class ThrowingDetector : IOfficeProvenanceSignalDetector {
        public string Name => "test";
        public OfficeProvenanceSignalKind SignalKind => OfficeProvenanceSignalKind.DeterministicArtifact;
        public OfficeProvenanceSignalResult Detect(string path) => throw new InvalidOperationException("Detector failure");
    }
    [Fact]
    public async Task DiscoveryIsBoundedFilteredAndDoesNotHideExplicitMissingInputs() {
        string first = Write("first.txt", "ok"); Write("sub/second.txt", "\u202E"); Write("bin/generated.txt", "\u202E"); Write("skip.txt", "\u202E"); Write("unknown.zzz", "unsupported");
        var selected = new OfficeProvenanceAuditRequest { Inputs = [_root, first], Include = ["*.txt"], Exclude = ["skip.txt"] };
        Assert.Equal(2, OfficeProvenanceAudit.Discover(selected).Count);
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceAudit.Discover(new() { Inputs = [_root], MaximumItems = 1 }));
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceAudit.Discover(new() { Inputs = [_root], Include = ["*.not-there"] }));
        var result = Assert.Single(await OfficeProvenanceAudit.RunAsync(new() { Inputs = [Path.Combine(_root, "missing.txt")] }));
        Assert.Equal(OfficeWorkflowFailureKind.InputNotFound, result.FailureKind);
        Assert.Equal(OfficeProvenanceCheckStatus.NotRequested, result.Checks.Structural);
        Assert.Equal(Path.Combine(_root, "missing.txt"), result.InputPath);
    }
    [Fact]
    public async Task AuditRejectsDirectoryReplacedByLinkAfterDiscovery() {
        string selected = Path.Combine(_root, "selected");
        string queued = Path.Combine(selected, "queued");
        string outside = Path.Combine(_root, "outside");
        Directory.CreateDirectory(queued);
        Directory.CreateDirectory(outside);
        File.WriteAllText(Path.Combine(queued, "report.txt"), "selected input");
        File.WriteAllText(Path.Combine(outside, "report.txt"), "\u202E outside input");

        var result = Assert.Single(await OfficeProvenanceAudit.RunAsync(
            new OfficeProvenanceAuditRequest { Inputs = [selected] },
            new SwapDirectoryBeforeAssessment(queued, outside, Path.Combine(_root, "parked"))));

        Assert.False(result.Succeeded);
        Assert.Null(result.InputSha256);
        Assert.Null(result.Assessment);
    }

    private sealed class SwapDirectoryBeforeAssessment(string queued, string outside, string parked)
        : IOfficeProvenanceWorkflowRunner {
        public Task<OfficeProvenanceWorkflowResult> RunProvenanceAsync(OfficeProvenanceWorkflowRequest request,
            IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default) =>
            new OfficeWorkflowRunner().RunProvenanceAsync(request, progress, cancellationToken);

        public Task<IReadOnlyList<OfficeProvenanceWorkflowResult>> RunProvenanceBatchAsync(
            IEnumerable<OfficeProvenanceWorkflowRequest> requests, OfficeProvenanceWorkflowBatchOptions? options = null,
            IProgress<OfficeWorkflowProgress>? progress = null, CancellationToken cancellationToken = default) {
            Directory.Move(queued, parked);
            Directory.CreateSymbolicLink(queued, outside);
            return new OfficeWorkflowRunner().RunProvenanceBatchAsync(requests, options, progress, cancellationToken);
        }
    }
    [Theory]
    [InlineData(OfficeWorkflowConflictPolicy.Fail)]
    [InlineData(OfficeWorkflowConflictPolicy.Replace)]
    [InlineData(OfficeWorkflowConflictPolicy.Rename)]
    public async Task PublicationGuardProtectsOwnedDestinations(OfficeWorkflowConflictPolicy policy) {
        string input = Write("page.html", "<html><body>before</body></html>");
        string output = Path.Combine(_root, "copy.html");
        var guard = new DestinationGuard(output);
        var result = await new OfficeWorkflowRunner().RunProvenanceAsync(new() { InputPath = input,
            Operation = OfficeProvenanceWorkflowOperation.Remove, OutputPath = output, ConflictPolicy = policy, PublicationGuard = guard });
        Assert.False(File.Exists(output));
        Assert.True(guard.Calls > 0);
        if (policy == OfficeWorkflowConflictPolicy.Rename) { Assert.True(result.Succeeded, result.Summary); Assert.NotEqual(output, result.OutputPath); Assert.True(File.Exists(result.OutputPath)); }
        else Assert.False(result.Succeeded);
    }
    private sealed class DestinationGuard(string owned) : IOfficeWorkflowPublicationGuard {
        internal int Calls;
        public ValueTask<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken token) { Calls++; return ValueTask.FromResult(!string.Equals(path, owned, StringComparison.Ordinal)); }
    }
    [Fact]
    public async Task ChangedReviewedFileCannotProduceAnArtifactAndSarifContainsFailures() {
        string input = Write("page.html", "<html><body>before</body></html>");
        var runner = new OfficeWorkflowRunner();
        var review = await runner.RunProvenanceAsync(new() { InputPath = input, Operation = OfficeProvenanceWorkflowOperation.Assess });
        File.WriteAllText(input, "<html><body>after</body></html>");
        string output = Path.Combine(_root, "copy.html");
        var result = await runner.RunProvenanceAsync(new() { InputPath = input, Operation = OfficeProvenanceWorkflowOperation.Remove, OutputPath = output, ExpectedInputSha256 = review.InputSha256 });
        Assert.False(result.Succeeded); Assert.False(File.Exists(output));
        using var sarif = JsonDocument.Parse(OfficeProvenanceSarif.Serialize([result]));
        var run = sarif.RootElement.GetProperty("runs")[0];
        Assert.False(run.GetProperty("invocations")[0].GetProperty("executionSuccessful").GetBoolean());
        Assert.Equal("officeimo.execution", run.GetProperty("results")[0].GetProperty("ruleId").GetString());
    }
}
