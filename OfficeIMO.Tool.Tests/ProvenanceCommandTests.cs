using System.Text.Json;
using OfficeIMO.Provenance;
using OfficeIMO.Tool.Commands.Provenance;
using OfficeIMO.Workflows;
using Xunit;

namespace OfficeIMO.Tool.Tests;

public sealed class ProvenanceCommandTests {
    [Fact]
    public async Task CapabilitiesReturnVersionedOwnerCatalog() {
        ToolResult result = await RunAsync(["provenance", "capabilities"]);

        Assert.Equal((int)OfficeImoToolExitCode.Success, result.ExitCode);
        using JsonDocument json = JsonDocument.Parse(result.Output);
        Assert.Equal("officeimo.provenance.capabilities.v1", json.RootElement.GetProperty("schema").GetString());
        JsonElement capabilities = json.RootElement.GetProperty("capabilities");
        Assert.Contains(capabilities.EnumerateArray(), item =>
            item.GetProperty("id").GetString() == "word-openxml" &&
            item.GetProperty("ownerPackage").GetString() == "OfficeIMO.Word");
        Assert.Equal(string.Empty, result.Error);
    }

    [Fact]
    public async Task InspectReturnsStableJsonWithoutChangingTopLevelInspectAlias() {
        using var scope = new TestDirectory();
        string input = scope.Write("page.html", HtmlWithManifest("body"));

        ToolResult result = await RunAsync(["provenance", "inspect", input]);

        Assert.Equal((int)OfficeImoToolExitCode.Success, result.ExitCode);
        using JsonDocument json = JsonDocument.Parse(result.Output);
        Assert.Equal("officeimo.provenance.result.v1", json.RootElement.GetProperty("schema").GetString());
        Assert.Equal("Inspect", json.RootElement.GetProperty("operation").GetString());
        Assert.Equal("OfficeIMO.Html", json.RootElement.GetProperty("ownerPackage").GetString());
        Assert.Equal(1, json.RootElement.GetProperty("inspection").GetProperty("evidence").GetArrayLength());
    }

    [Fact]
    public async Task AssessCanReportUnicodeIntegrityInJson() {
        using var scope = new TestDirectory();
        string input = scope.Write("page.html", HtmlWithManifest("review \u202Ethis"));

        ToolResult result = await RunAsync(["provenance", "assess", input]);

        Assert.Equal((int)OfficeImoToolExitCode.Success, result.ExitCode);
        using JsonDocument json = JsonDocument.Parse(result.Output);
        JsonElement findings = json.RootElement.GetProperty("assessment").GetProperty("textIntegrity");
        Assert.Contains(findings.EnumerateArray(), item => item.GetProperty("kind").GetString() == "BidirectionalControl");
    }

    [Fact]
    public async Task AssessmentTextOutputIncludesVerificationIntegrityAndProviderEvidence() {
        using var scope = new TestDirectory();
        string input = scope.Write("page.html", "<!doctype html><html><body>review \u202Ethis</body></html>");
        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner(
            new InvalidVerifier(),
            [new InconclusiveDetector()]).RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Assess,
                InputPath = input
            });
        Assert.True(result.Succeeded, result.Summary);
        using var output = new StringWriter();

        await ProvenanceOutput.WriteResultAsync(output, result, ProvenanceOutputFormat.Text);

        string text = output.ToString();
        Assert.Contains("Verification: test-verifier | status=Invalid", text, StringComparison.Ordinal);
        Assert.Contains("Verification finding: content binding failed", text, StringComparison.Ordinal);
        Assert.Contains("PotentiallyDangerous | BidirectionalControl | U+202E", text, StringComparison.Ordinal);
        Assert.Contains("Provider signal: test-detector | StatisticalTextWatermark | status=Inconclusive", text, StringComparison.Ordinal);
        Assert.Contains("Provider finding: sample too short", text, StringComparison.Ordinal);
    }

    [Fact]
    public async Task RemovePublishesVerifiedArtifactAndRefusesReplacementByDefault() {
        using var scope = new TestDirectory();
        string input = scope.Write("page.html", HtmlWithManifest("keep"));
        string output = Path.Combine(scope.Path, "cleaned.html");

        ToolResult first = await RunAsync(["provenance", "remove", input, "--output", output]);
        ToolResult refused = await RunAsync(["provenance", "remove", input, "--output", output]);
        ToolResult replaced = await RunAsync(["provenance", "remove", input, "--output", output, "--force"]);

        Assert.Equal((int)OfficeImoToolExitCode.Success, first.ExitCode);
        Assert.Equal((int)OfficeImoToolExitCode.OutputFailed, refused.ExitCode);
        Assert.Equal((int)OfficeImoToolExitCode.Success, replaced.ExitCode);
        Assert.DoesNotContain("c2pa-manifest", File.ReadAllText(output), StringComparison.OrdinalIgnoreCase);
        using JsonDocument json = JsonDocument.Parse(replaced.Output);
        Assert.True(json.RootElement.GetProperty("wasChanged").GetBoolean());
        Assert.Equal(0, json.RootElement.GetProperty("after").GetProperty("evidence").GetArrayLength());
    }

    [Fact]
    public async Task BatchInspectIsBoundedAndMachineReadable() {
        using var scope = new TestDirectory();
        string first = scope.Write("first.html", HtmlWithManifest("one"));
        string second = scope.Write("second.html", "<html><body>two</body></html>");

        ToolResult result = await RunAsync([
            "provenance", "batch", "inspect", first, second, "--max-items", "2"
        ]);

        Assert.Equal((int)OfficeImoToolExitCode.Success, result.ExitCode);
        using JsonDocument json = JsonDocument.Parse(result.Output);
        Assert.Equal("officeimo.provenance.batch.v1", json.RootElement.GetProperty("schema").GetString());
        Assert.Equal(2, json.RootElement.GetProperty("results").GetArrayLength());
    }

    [Fact]
    public async Task BatchRemoveRejectsDuplicateBasenamesBeforeForceCanOverwrite() {
        using var scope = new TestDirectory();
        string firstDirectory = Path.Combine(scope.Path, "first");
        string secondDirectory = Path.Combine(scope.Path, "second");
        string outputDirectory = Path.Combine(scope.Path, "output");
        Directory.CreateDirectory(firstDirectory);
        Directory.CreateDirectory(secondDirectory);
        string first = Path.Combine(firstDirectory, "page.html");
        string second = Path.Combine(secondDirectory, "page.html");
        File.WriteAllText(first, HtmlWithManifest("first"));
        File.WriteAllText(second, HtmlWithManifest("second"));

        ToolResult result = await RunAsync([
            "provenance", "batch", "remove", first, second,
            "--output-directory", outputDirectory, "--force"
        ]);

        Assert.Equal((int)OfficeImoToolExitCode.Usage, result.ExitCode);
        Assert.Contains("same output path", result.Error, StringComparison.Ordinal);
        Assert.False(Directory.Exists(outputDirectory));
    }

    [Fact]
    public async Task PreCancelledBatchReturnsCancelledExitCodeAndResult() {
        using var scope = new TestDirectory();
        string first = scope.Write("first.html", "<html><body>first</body></html>");
        string second = scope.Write("second.html", "<html><body>second</body></html>");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        ToolResult result = await RunAsync(
            ["provenance", "batch", "inspect", first, second],
            cancellation.Token);

        Assert.Equal((int)OfficeImoToolExitCode.Cancelled, result.ExitCode);
        using JsonDocument json = JsonDocument.Parse(result.Output);
        JsonElement item = Assert.Single(json.RootElement.GetProperty("results").EnumerateArray());
        Assert.Equal("Cancelled", item.GetProperty("status").GetString());
    }

    [Fact]
    public async Task BatchCancellationTakesExitCodePrecedenceOverAnEarlierFailure() {
        using var scope = new TestDirectory();
        string input = scope.Write("input.html", "<html><body>input</body></html>");
        string missing = Path.Combine(scope.Path, "missing.html");
        using var output = new StringWriter();
        using var error = new StringWriter();

        int exitCode = await ProvenanceCommand.RunAsync(
            ["batch", "inspect", input],
            output,
            error,
            runner: new FailureThenCancellationRunner(missing, input));

        Assert.Equal((int)OfficeImoToolExitCode.Cancelled, exitCode);
        using JsonDocument json = JsonDocument.Parse(output.ToString());
        Assert.Collection(
            json.RootElement.GetProperty("results").EnumerateArray(),
            item => Assert.Equal("Failed", item.GetProperty("status").GetString()),
            item => Assert.Equal("Cancelled", item.GetProperty("status").GetString()));
    }

    [Fact]
    public async Task BatchRemoveReturnsStructuredMixedResultsWhenAnInputIsMissing() {
        using var scope = new TestDirectory();
        string input = scope.Write("first.html", HtmlWithManifest("first"));
        string missing = Path.Combine(scope.Path, "missing.html");
        string outputDirectory = Path.Combine(scope.Path, "output");

        ToolResult result = await RunAsync([
            "provenance", "batch", "remove", input, missing,
            "--output-directory", outputDirectory
        ]);

        Assert.Equal((int)OfficeImoToolExitCode.InputNotFound, result.ExitCode);
        using JsonDocument json = JsonDocument.Parse(result.Output);
        JsonElement.ArrayEnumerator results = json.RootElement.GetProperty("results").EnumerateArray();
        Assert.Collection(
            results,
            item => Assert.Equal("Completed", item.GetProperty("status").GetString()),
            item => Assert.Equal("InputNotFound", item.GetProperty("failureKind").GetString()));
        Assert.True(File.Exists(Path.Combine(outputDirectory, "first.provenance-cleaned.html")));
        Assert.False(File.Exists(Path.Combine(outputDirectory, "missing.provenance-cleaned.html")));
    }

    [Fact]
    public async Task UnsafeOrMeaninglessRemovalOptionsReturnUsage() {
        using var scope = new TestDirectory();
        string input = scope.Write("page.html", HtmlWithManifest("body"));

        ToolResult result = await RunAsync([
            "provenance", "remove", input, "--keep-c2pa", "--keep-external-c2pa", "--keep-ai-source"
        ]);

        Assert.Equal((int)OfficeImoToolExitCode.Usage, result.ExitCode);
        Assert.Contains("at least one selected carrier", result.Error, StringComparison.Ordinal);
    }

    [Fact]
    public async Task MissingInputUsesStableInputNotFoundExitCode() {
        using var scope = new TestDirectory();
        string missing = Path.Combine(scope.Path, "missing.html");

        ToolResult result = await RunAsync(["provenance", "inspect", missing]);

        Assert.Equal((int)OfficeImoToolExitCode.InputNotFound, result.ExitCode);
        using JsonDocument json = JsonDocument.Parse(result.Output);
        Assert.Equal("InputNotFound", json.RootElement.GetProperty("failureKind").GetString());
    }

    [Fact]
    public void CliByteLimitsFlowIntoEveryOwningParserBoundary() {
        const long inputLimit = 768L * 1024L * 1024L;
        const long outputLimit = 1024L * 1024L * 1024L;
        ProvenanceArguments parsed = ProvenanceArguments.Parse([
            "remove", "input.html", "--max-input-bytes", inputLimit.ToString(),
            "--max-output-bytes", outputLimit.ToString()
        ]);

        OfficeProvenanceWorkflowRequest request = ProvenanceCommand.CreateRequest(
            parsed,
            "input.html",
            "output.html");

        Assert.Equal(inputLimit, request.Limits.MaximumInputBytes);
        Assert.Equal(outputLimit, request.Limits.MaximumOutputBytes);
        Assert.Equal(inputLimit, request.Inspection.MaxAssetBytes);
        Assert.Equal(inputLimit, request.Assessment.Structural.MaxAssetBytes);
        Assert.Equal(inputLimit, request.Assessment.TextIntegrity.MaxEncodedBytes);
        Assert.Equal(inputLimit, request.Removal.Limits.MaxAssetBytes);
        Assert.Equal(outputLimit, request.Removal.MaxOutputBytes);
    }

    [Fact]
    public void CliLongLimitsClampOnlyMaterializingParserOptions() {
        long configured = (long)int.MaxValue + 4096L;
        ProvenanceArguments parsed = ProvenanceArguments.Parse([
            "remove", "input.html", "--max-input-bytes", configured.ToString(),
            "--max-output-bytes", configured.ToString()
        ]);

        OfficeProvenanceWorkflowRequest request = ProvenanceCommand.CreateRequest(
            parsed,
            "input.html",
            "output.html");

        Assert.Equal(configured, request.Limits.MaximumInputBytes);
        Assert.Equal(configured, request.Limits.MaximumOutputBytes);
        Assert.Equal(int.MaxValue, request.Inspection.MaxAssetBytes);
        Assert.Equal(int.MaxValue, request.Assessment.Structural.MaxAssetBytes);
        Assert.Equal(int.MaxValue, request.Assessment.TextIntegrity.MaxEncodedBytes);
        Assert.Equal(int.MaxValue, request.Removal.Limits.MaxAssetBytes);
        Assert.Equal(int.MaxValue, request.Removal.MaxOutputBytes);
    }

    private static string HtmlWithManifest(string body) =>
        "<!doctype html><html><head><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body>" + body + "</body></html>";

    private static async Task<ToolResult> RunAsync(string[] args, CancellationToken cancellationToken = default) {
        await using var output = new MemoryStream();
        using var error = new StringWriter();
        int exitCode = await OfficeImoToolApp.RunAsync(args, Stream.Null, output, error, cancellationToken);
        return new ToolResult(exitCode, Encoding.UTF8.GetString(output.ToArray()), error.ToString());
    }

    private sealed record ToolResult(int ExitCode, string Output, string Error);

    private sealed class InvalidVerifier : IOfficeProvenanceVerifier {
        public string Name => "test-verifier";

        public OfficeProvenanceVerificationResult Verify(
            string filePath,
            OfficeProvenanceVerificationOptions? options = null) =>
            new(OfficeProvenanceVerificationStatus.Invalid, Name, ["content binding failed"]);
    }

    private sealed class InconclusiveDetector : IOfficeProvenanceSignalDetector {
        public string Name => "test-detector";
        public OfficeProvenanceSignalKind SignalKind => OfficeProvenanceSignalKind.StatisticalTextWatermark;

        public OfficeProvenanceSignalResult Detect(string filePath) =>
            new(Name, SignalKind, OfficeProvenanceSignalStatus.Inconclusive, ["sample too short"]);
    }

    private sealed class FailureThenCancellationRunner : IOfficeProvenanceWorkflowRunner {
        private readonly string _missingPath;
        private readonly string _existingPath;
        private readonly OfficeWorkflowRunner _inner = new OfficeWorkflowRunner();

        internal FailureThenCancellationRunner(string missingPath, string existingPath) {
            _missingPath = missingPath;
            _existingPath = existingPath;
        }

        public Task<OfficeProvenanceWorkflowResult> RunProvenanceAsync(
            OfficeProvenanceWorkflowRequest request,
            IProgress<OfficeWorkflowProgress>? progress = null,
            CancellationToken cancellationToken = default) =>
            _inner.RunProvenanceAsync(request, progress, cancellationToken);

        public async Task<IReadOnlyList<OfficeProvenanceWorkflowResult>> RunProvenanceBatchAsync(
            IEnumerable<OfficeProvenanceWorkflowRequest> requests,
            OfficeProvenanceWorkflowBatchOptions? options = null,
            IProgress<OfficeWorkflowProgress>? progress = null,
            CancellationToken cancellationToken = default) {
            OfficeProvenanceWorkflowResult failure = await _inner.RunProvenanceAsync(
                new OfficeProvenanceWorkflowRequest {
                    Operation = OfficeProvenanceWorkflowOperation.Inspect,
                    InputPath = _missingPath
                });
            using var cancelled = new CancellationTokenSource();
            cancelled.Cancel();
            OfficeProvenanceWorkflowResult cancellation = await _inner.RunProvenanceAsync(
                new OfficeProvenanceWorkflowRequest {
                    Operation = OfficeProvenanceWorkflowOperation.Inspect,
                    InputPath = _existingPath
                },
                cancellationToken: cancelled.Token);
            return [failure, cancellation];
        }
    }

    private sealed class TestDirectory : IDisposable {
        internal TestDirectory() {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "officeimo-tool-provenance-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        internal string Path { get; }

        internal string Write(string fileName, string contents) {
            string path = System.IO.Path.Combine(Path, fileName);
            File.WriteAllText(path, contents, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
            return path;
        }

        public void Dispose() {
            try {
                if (Directory.Exists(Path)) Directory.Delete(Path, recursive: true);
            } catch (IOException) {
                // Best-effort test cleanup.
            } catch (UnauthorizedAccessException) {
                // Best-effort test cleanup.
            }
        }
    }
}
