using System.Security.Cryptography;
using OfficeIMO.AI;

internal sealed record EvaluationAttempt(OfficeAiResult? Result, string? Failure, EvaluationScore? Score,
    bool ContractPassed, string? ReportSha256);

// One case owns inference, export, scoring and report fingerprinting. Failures do not discard other cases.
internal static class EvaluationCaseRunner {
    internal static async Task<EvaluationAttempt> ExecuteAsync(EvaluationCase item, OfficeAiRequest request,
        IOfficeAiExecutor executor, string sourcePath, string directory, CancellationToken operationToken, CancellationToken deadlineToken) {
        OfficeAiResult? result = null;
        string? failure = null;
        EvaluationScore? score = null;
        bool accepted = false;
        string? reportHash = null;
        try {
            var document = await DocumentInputs.ReadAsync(item.Source, sourcePath, item.Images, request.Pages, request.Limits, operationToken);
            result = await new OfficeAiEngine(executor).RunAsync(document, request, cancellationToken: operationToken);
            await ArtifactWriter.SaveAsync(directory, document, result, operationToken);
            reportHash = Convert.ToHexString(SHA256.HashData(
                await File.ReadAllBytesAsync(Path.Combine(directory, "report.json"), operationToken))).ToLowerInvariant();
            score = item.Gold.Score(result);
            accepted = result.Status != OfficeAiResultStatus.InvalidResponse && score.ContractPassed;
        } catch (OperationCanceledException) when (!deadlineToken.IsCancellationRequested) { failure = "case-timeout"; }
          catch (OperationCanceledException) { throw; }
          catch (Exception error) when (error is not OutOfMemoryException) { failure = error.GetType().Name; }

        return new(result, failure, score, accepted, reportHash);
    }
}
