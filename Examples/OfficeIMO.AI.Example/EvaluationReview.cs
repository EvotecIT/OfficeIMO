using System.Security.Cryptography;
using System.Text.Json;
using System.Text.Json.Serialization;

// Independent labels bind to the exact saved source/evidence/result report, rather than case names alone.
internal sealed record EvaluationReviewAnnotation(string Id, int Repetition, string ReportSha256,
    int? UnsupportedClaims, int? OmittedFacts, int? IncorrectRelationships, string? Notes) {
    public bool IsComplete => UnsupportedClaims.HasValue && OmittedFacts.HasValue && IncorrectRelationships.HasValue
        && !string.IsNullOrWhiteSpace(Notes);
    public bool Passed => IsComplete && UnsupportedClaims == 0 && OmittedFacts == 0 && IncorrectRelationships == 0;
}
internal sealed record EvaluationReviewAnnotations(string Schema, string Reviewer,
    IReadOnlyList<EvaluationReviewAnnotation> Annotations, string EvaluationSha256);

internal static class EvaluationReview {
    internal static readonly JsonSerializerOptions JsonOptions = new() {
        WriteIndented = true, PropertyNameCaseInsensitive = true, MaxDepth = 32,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow
    };

    public static async Task<int> RunAsync(ExampleOptions options, CancellationToken token) {
        string directory = Path.GetFullPath(options.ReviewEvaluationPath!);
        byte[] evaluationBytes = await DocumentInputs.ReadFileAsync(Path.Combine(directory, "evaluation.json"), 8_000_000, token);
        byte[] reviewBytes = await DocumentInputs.ReadFileAsync(options.AnnotationsPath!, 1_000_000, token);
        using var evaluation = JsonDocument.Parse(evaluationBytes);
        if (evaluation.RootElement.GetProperty("schema").GetString() != "officeimo.ai.evaluation.v3")
            throw new InvalidDataException("Review requires evaluation schema v3.");
        var reviews = JsonSerializer.Deserialize<EvaluationReviewAnnotations>(reviewBytes, JsonOptions)
            ?? throw new InvalidDataException("Review annotations are empty.");
        if (reviews.Schema != "officeimo.ai.semantic-review.v2" || string.IsNullOrWhiteSpace(reviews.Reviewer)
            || reviews.Annotations is null || reviews.Annotations.Count > 300)
            throw new InvalidDataException("Provide a reviewer and bounded semantic-review v2 annotations.");
        string evaluationHash = Convert.ToHexString(SHA256.HashData(evaluationBytes));
        if (!string.Equals(evaluationHash, reviews.EvaluationSha256, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("The semantic review does not match the saved evaluation contract. Review the changed gold and outcomes again.");
        var byRun = new Dictionary<(string, int), EvaluationReviewAnnotation>();
        foreach (var review in reviews.Annotations) {
            if (review is null || string.IsNullOrWhiteSpace(review.Id) || review.Repetition is < 1 or > 3
                || review.ReportSha256 is null || review.ReportSha256.Length != 64 || !review.ReportSha256.All(Uri.IsHexDigit)
                || review.UnsupportedClaims < 0 || review.OmittedFacts < 0 || review.IncorrectRelationships < 0
                || !byRun.TryAdd((review.Id, review.Repetition), review))
                throw new InvalidDataException("Invalid or duplicate review annotation.");
        }
        var rows = new List<object>();
        var seen = new HashSet<(string, int)>();
        int passed = 0, pending = 0, failedContracts = 0, failedAssessments = 0;
        foreach (var row in evaluation.RootElement.GetProperty("cases").EnumerateArray()) {
            if (rows.Count >= 300) throw new InvalidDataException("Evaluation exceeds the review limit.");
            string id = row.GetProperty("Id").GetString() ?? string.Empty;
            int repetition = row.GetProperty("repetition").GetInt32();
            if (id.Length == 0 || id.Length > 100 || id.Any(value => value is not (>= 'a' and <= 'z' or >= '0' and <= '9' or '-'))
                || repetition is < 1 or > 3 || !seen.Add((id, repetition)))
                throw new InvalidDataException("Invalid or duplicate evaluation run.");
            bool contractPassed = row.GetProperty("contractPassed").GetBoolean();
            if (!contractPassed) failedContracts++;
            bool assessed = byRun.TryGetValue((id, repetition), out var review);
            if (assessed) {
                byte[] report = await DocumentInputs.ReadFileAsync(Path.Combine(directory, id, "run-" + repetition, "report.json"), 8_000_000, token);
                string actual = Convert.ToHexString(SHA256.HashData(report));
                if (!string.Equals(actual, review!.ReportSha256, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("The semantic review does not match the saved report. Review the changed result again.");
                string sourceFile = row.GetProperty("sourceFile").GetString() ?? string.Empty;
                if (sourceFile is not ("source.txt" or "source.pdf" or "source.png" or "source.jpg" or "source.jpeg" or "source.webp"))
                    throw new InvalidDataException("Invalid evaluation source filename.");
                byte[] source = await DocumentInputs.ReadFileAsync(Path.Combine(directory, id, "run-" + repetition, sourceFile), 64_000_000, token);
                if (!string.Equals(Convert.ToHexString(SHA256.HashData(source)), row.GetProperty("sourceHash").GetString(), StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("The saved evaluation source has changed. Review the original captured source.");
                assessed = review.IsComplete;
            }
            if (!assessed) pending++;
            else if (!review!.Passed) failedAssessments++;
            bool qualityPassed = contractPassed && assessed && review!.Passed;
            if (qualityPassed) passed++;
            rows.Add(new { id, repetition, contractPassed, semanticAssessment = assessed ? review!.Passed ? "passed" : "failed" : "pending",
                passed = qualityPassed, review });
        }
        if (byRun.Keys.Any(key => !seen.Contains(key))) throw new InvalidDataException("Annotations contain an unknown evaluation run.");
        int total = evaluation.RootElement.GetProperty("total").GetInt32();
        if (total < rows.Count || total <= 0 || total > 300) throw new InvalidDataException("Invalid evaluation run count.");
        pending += total - rows.Count;
        var reportOutput = new { schema = "officeimo.ai.reviewed-evaluation.v1", reviews.Reviewer,
            evaluationSha256 = Convert.ToHexString(SHA256.HashData(evaluationBytes)).ToLowerInvariant(),
            corpus = evaluation.RootElement.GetProperty("corpus").GetString(), passed, pending, total,
            qualityPassed = passed == total, cases = rows };
        token.ThrowIfCancellationRequested();
        string outputPath = Path.GetFullPath(options.OutputPath!);
        string temporary = outputPath + ".review-" + Guid.NewGuid().ToString("N") + ".tmp";
        try {
            await using (var output = new FileStream(temporary, FileMode.CreateNew, FileAccess.Write, FileShare.None)) {
                await JsonSerializer.SerializeAsync(output, reportOutput, JsonOptions, token);
                await output.FlushAsync(token);
            }
            token.ThrowIfCancellationRequested();
            File.Move(temporary, outputPath); // Same-directory, exclusive publication after complete serialization.
        } finally {
            if (File.Exists(temporary)) File.Delete(temporary);
        }
        Console.WriteLine($"Quality passed: {passed}/{total}; pending independent review: {pending}.");
        return passed == total ? 0 : pending > 0 && failedContracts == 0 && failedAssessments == 0 ? 4 : 1;
    }
}
