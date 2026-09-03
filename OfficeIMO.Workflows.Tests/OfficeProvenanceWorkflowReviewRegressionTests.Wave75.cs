namespace OfficeIMO.Workflows.Tests;

public sealed partial class OfficeProvenanceWorkflowTests {
    [Fact]
    public async Task EmptyBatchMaterializationPreservesEnumeratorCancellation() {
        using var cancellation = new CancellationTokenSource();

        IEnumerable<OfficeProvenanceWorkflowRequest> Requests() {
            cancellation.Cancel();
            yield break;
        }

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            new OfficeWorkflowRunner().RunProvenanceBatchAsync(
                Requests(),
                cancellationToken: cancellation.Token));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task AssessmentRejectsAbsoluteFileManifestDependenciesEvenWhenStructurallyAmbiguous(
        bool useBaseElement) {
        using var scope = new TempScope();
        string manifest = scope.Write("claim.c2pa", "mutable claim");
        string manifestReference = new Uri(manifest).AbsoluteUri;
        string head = useBaseElement
            ? $"<base href=\"{new Uri(scope.Path + Path.DirectorySeparatorChar).AbsoluteUri}\"><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"><link rel=\"c2pa-manifest\" href=\"claim.c2pa\">"
            : $"<link rel=\"c2pa-manifest\" href=\"{manifestReference}\">";
        string input = scope.Write("page.html", $"<!doctype html><html><head>{head}</head><body>body</body></html>");
        var verifier = new NestedRelativeManifestVerifier("claim.c2pa", "mutable claim");

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner(verifier).RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Assess,
                InputPath = input
            });

        Assert.False(result.Succeeded);
        Assert.Contains("absolute file-based external provenance manifest", result.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.Null(verifier.ObservedDirectory);
    }

    [Fact]
    public async Task AssessmentRejectsRootedFileBaseBeforeCallingProviders() {
        using var scope = new TempScope();
        string input = scope.Write(
            "page.html",
            "<!doctype html><html><head><base href=\"/outside/\"><link rel=\"c2pa-manifest\" href=\"claim.c2pa\"></head><body>body</body></html>");
        var verifier = new NestedRelativeManifestVerifier("claim.c2pa", "unused");

        OfficeProvenanceWorkflowResult result = await new OfficeWorkflowRunner(verifier).RunProvenanceAsync(
            new OfficeProvenanceWorkflowRequest {
                Operation = OfficeProvenanceWorkflowOperation.Assess,
                InputPath = input
            });

        Assert.False(result.Succeeded);
        Assert.Contains("absolute file-based external provenance manifest", result.Summary, StringComparison.OrdinalIgnoreCase);
        Assert.Null(verifier.ObservedDirectory);
    }
}
