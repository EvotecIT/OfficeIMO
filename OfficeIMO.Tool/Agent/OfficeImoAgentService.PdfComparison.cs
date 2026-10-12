using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    internal async Task<AgentPdfWorkflowResult> ComparePdfAsync(string sourceId, string comparisonSourceId,
        string outputPath, PdfWorkflowSettings settings, string? actualPages = null,
        string? comparisonPasswordEnvironmentVariable = null, int maxOutputCharacters = 4000,
        CancellationToken cancellationToken = default) {
        settings.Validate(); maxOutputCharacters = ValidateOutputBudget(maxOutputCharacters);
        var actualSettings = new PdfWorkflowSettings {
            Pages = actualPages, PasswordEnvironmentVariable = comparisonPasswordEnvironmentVariable,
            MaximumInputBytes = settings.MaximumInputBytes, MaximumOutputBytes = settings.MaximumOutputBytes
        };
        actualSettings.Validate();
        string input = ResolvePdfInput(_registry.Resolve(sourceId, cancellationToken).Path);
        string actual = ResolvePdfInput(_registry.Resolve(comparisonSourceId, cancellationToken).Path);
        void VerifySources(CancellationToken token) {
            _ = ResolvePdfInput(_registry.Resolve(sourceId, token).Path);
            _ = ResolvePdfInput(_registry.Resolve(comparisonSourceId, token).Path);
        }
        OfficeWorkflowStreamInput Source(string path) => new(Path.GetFileName(path), token => {
            VerifySources(token);
            return Task.FromResult<Stream>(File.OpenRead(_pathPolicy.ResolveInput(path)));
        });
        string destination = PreparePdfOutput(outputPath, [input, actual], false, settings.Overwrite, "compare", maxOutputCharacters, ".html");
        var result = await new OfficeWorkflowRunner().RunAsync(new OfficeWorkflowRequest {
            Operation = OfficeWorkflowOperation.Compare, InputPath = input, ComparisonPath = actual,
            InputStream = Source(input), ComparisonStream = Source(actual), OutputPath = destination,
            ComparisonExpectedPages = settings.Selector(), ComparisonActualPages = actualSettings.Selector(),
            PdfPassword = PdfPassword(settings), ComparisonPdfPassword = PdfPassword(actualSettings),
            ConflictPolicy = settings.ConflictPolicy, Limits = settings.Limits(),
            PublicationGuard = new PdfRootPublicationGuard(_pathPolicy, [input, actual], VerifySources)
        }, cancellationToken: cancellationToken).ConfigureAwait(false);
        return PdfResult("compare", result.Status, result.FailureKind, result.OutputPath, result.OutputBytes,
            result.OutputPath is null ? [] : [new AgentPdfArtifact { Path = result.OutputPath, SizeBytes = result.OutputBytes }],
            result.Diagnostics, maxOutputCharacters, summary: result.Summary);
    }
}
