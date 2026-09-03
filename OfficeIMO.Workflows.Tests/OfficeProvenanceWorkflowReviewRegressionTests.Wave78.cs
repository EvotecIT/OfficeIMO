using System.Runtime.InteropServices;
using OfficeIMO.Provenance;

namespace OfficeIMO.Workflows.Tests;

public sealed partial class OfficeProvenanceWorkflowTests {
    [Theory]
    [InlineData("../claim.c2pa")]
    [InlineData("%2e%2e/claim.c2pa")]
    public async Task AssessmentRejectsParentRelativeManifestBeforeCallingProviders(string manifestReference) {
        using var scope = new TempScope();
        string assetDirectory = Path.Combine(scope.Path, "assets");
        Directory.CreateDirectory(assetDirectory);
        File.WriteAllText(Path.Combine(scope.Path, "claim.c2pa"), "outside snapshot child");
        string input = Path.Combine(assetDirectory, "page.html");
        File.WriteAllText(
            input,
            "<!doctype html><html><head><link rel=\"c2pa-manifest\" href=\"" + manifestReference +
            "\"></head><body>body</body></html>");
        var verifier = new NestedRelativeManifestVerifier("claim.c2pa", "unused");

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner(verifier).RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Assess,
                InputPath = input
            });

        Assert.False(result.Succeeded);
        Assert.Contains("cannot be bound within the immutable snapshot", result.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.Null(verifier.ObservedDirectory);
    }

    [Fact]
    public async Task RemovalSharesExpandedByteBudgetAcrossWorkflowStages() {
        using var scope = new TempScope();
        string input = scope.Write(
            "page.html",
            "<!doctype html><html><head><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head>" +
            "<body><img src=\"data:image/png;base64,AQIDBA==\"></body></html>");
        string output = Path.Combine(scope.Path, "cleaned.html");
        var request = new OfficeProvenanceWorkflowRequest {
            Operation = OfficeProvenanceWorkflowOperation.Remove,
            InputPath = input,
            OutputPath = output
        };
        request.Removal.Limits.MaxExpandedContainerBytes = 4;

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner().RunProvenanceAsync(request);

        Assert.False(result.Succeeded);
        Assert.Contains("expanded-container limit", result.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(output));
    }

    [Fact]
    public void SnapshotStorageFailuresUseTheExecutionFailureContract() {
        var exception = new IOException("snapshot storage unavailable");

        OfficeWorkflowFailureKind kind = OfficeWorkflowRunner.ClassifyFailure(
            exception,
            OfficeWorkflowRunner.WorkflowFailureStage.Snapshot);

        Assert.Equal(OfficeWorkflowFailureKind.OperationFailed, kind);
        Assert.Equal("snapshot", OfficeWorkflowRunner.GetDiagnosticStage(
            OfficeWorkflowRunner.WorkflowFailureStage.Snapshot));
    }

    [Fact]
    public async Task BatchReplacementRejectsDestinationRetargetedToAnotherInput() {
#if NET8_0_OR_GREATER
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows)) return;
        using var scope = new TempScope();
        string firstInput = scope.Write("first.html", HtmlWithExternalManifest("first"));
        string secondInput = scope.Write("second.html", "<!doctype html><html><body>second input</body></html>");
        string benignDestination = scope.Write("existing.html", "existing destination");
        string output = Path.Combine(scope.Path, "cleaned.html");
        File.CreateSymbolicLink(output, benignDestination);
        var progress = new RetargetingBatchOutputProgress("first", output, secondInput);

        IReadOnlyList<OfficeProvenanceWorkflowResult> results = await new OfficeWorkflowRunner().RunProvenanceBatchAsync([
            new OfficeProvenanceWorkflowRequest {
                Id = "first",
                Operation = OfficeProvenanceWorkflowOperation.Remove,
                InputPath = firstInput,
                OutputPath = output,
                ConflictPolicy = OfficeWorkflowConflictPolicy.Replace
            },
            new OfficeProvenanceWorkflowRequest {
                Id = "second",
                Operation = OfficeProvenanceWorkflowOperation.Inspect,
                InputPath = secondInput
            }
        ], progress: progress);

        Assert.True(progress.Retargeted);
        Assert.Equal(OfficeWorkflowFailureKind.OutputFailed, results[0].FailureKind);
        Assert.Contains("overlaps another batch request", results[0].Summary, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("<!doctype html><html><body>second input</body></html>", File.ReadAllText(secondInput));
#endif
    }

#if NET8_0_OR_GREATER
    private sealed class RetargetingBatchOutputProgress(
        string requestId,
        string outputPath,
        string targetPath) : IProgress<OfficeWorkflowProgress> {
        internal bool Retargeted { get; private set; }

        public void Report(OfficeWorkflowProgress value) {
            if (Retargeted || value.RequestId != requestId || value.Stage != "publish") return;
            File.Delete(outputPath);
            File.CreateSymbolicLink(outputPath, targetPath);
            Retargeted = true;
        }
    }
#endif
}
