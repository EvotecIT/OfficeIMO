using System.Security.Cryptography;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class EvaluationReviewTests {
    [Theory]
    [InlineData(0, 0, 0, true)]
    [InlineData(1, 0, 0, false)]
    [InlineData(0, 1, 0, false)]
    [InlineData(0, 0, 1, false)]
    public void IndependentLabelsDecideMeaningEvenWhenMarkersMatch(int unsupported, int omitted, int relationships, bool passed) {
        var result = new OfficeAiResult {
            RequestId = "test", SourceHash = "test", SnapshotHash = "test", Status = OfficeAiResultStatus.Completed,
            Profile = new() { Id = "fixture", Provider = "fixture", Model = "fixture", IsLocal = true },
            Claims = [new("The document does NOT say 17 or 23 or 2031-11-19.", [new("e1", 1, null, false)])]
        };
        var score = new EvaluationGold(FactMarkers: ["17", "23", "2031-11-19"]).Score(result);
        Assert.True(score.ContractPassed); // Mechanical recall does not decide entailment.
        Assert.True(score.SemanticReviewRequired);
        var annotation = new EvaluationReviewAnnotation("case", 1, new string('a', 64), unsupported, omitted, relationships, "Checked against the independent source gold.");
        Assert.Equal(passed, annotation.Passed);
        Assert.False((annotation with { UnsupportedClaims = null }).Passed);
        Assert.False((annotation with { Notes = null }).Passed);
        Assert.False(new EvaluationGold(FactMarkers: ["17"]).Score(result with {
            Claims = [new(result.Claims[0].Text, [])]
        }).ContractPassed);
    }

    [Theory]
    [InlineData("pass", 0)]
    [InlineData("unsupported", 1)]
    [InlineData("pending", 4)]
    [InlineData("stale", 2)]
    [InlineData("changed-source", 2)]
    [InlineData("duplicate", 2)]
    [InlineData("unknown", 2)]
    [InlineData("changed-gold", 2)]
    [InlineData("changed-contract", 2)]
    public async Task OfflineReviewUsesExactSavedReportsAndIndependentLabels(string scenario, int expected) {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-evaluation-review-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(Path.Combine(root, "case", "run-1"));
        try {
            byte[] report = System.Text.Encoding.UTF8.GetBytes("source and proposed result");
            await File.WriteAllBytesAsync(Path.Combine(root, "case", "run-1", "report.json"), report);
            await File.WriteAllTextAsync(Path.Combine(root, "evaluation.json"), JsonSerializer.Serialize(new {
                schema = "officeimo.ai.evaluation.v3", corpus = "fixture", total = 1,
                cases = new[] { new { Id = "case", repetition = 1, contractPassed = scenario != "changed-contract", gold = "original", sourceFile = "source.txt", sourceHash = Convert.ToHexString(SHA256.HashData(report)) } }
            }));
            await File.WriteAllBytesAsync(Path.Combine(root, "case", "run-1", "source.txt"), scenario == "changed-source" ? [1, 2, 3] : report);
            var annotation = new EvaluationReviewAnnotation(scenario == "unknown" ? "other" : "case", 1,
                scenario == "stale" ? new string('0', 64) : Convert.ToHexString(SHA256.HashData(report)),
                scenario == "pending" ? null : scenario == "unsupported" ? 1 : 0, 0, 0, "Compared source and output.");
            await File.WriteAllTextAsync(Path.Combine(root, "labels.json"), JsonSerializer.Serialize(new EvaluationReviewAnnotations(
                "officeimo.ai.semantic-review.v2", "independent reviewer", scenario == "duplicate" ? [annotation, annotation] : [annotation],
                Convert.ToHexString(SHA256.HashData(await File.ReadAllBytesAsync(Path.Combine(root, "evaluation.json")))))));
            if (scenario is "changed-gold" or "changed-contract") {
                string evaluationPath = Path.Combine(root, "evaluation.json");
                string original = await File.ReadAllTextAsync(evaluationPath);
                await File.WriteAllTextAsync(evaluationPath, scenario == "changed-gold"
                    ? original.Replace("original", "changed") : original.Replace("false", "true"));
            }
            var options = ExampleOptions.Parse(["--review-evaluation", root, "--annotations", Path.Combine(root, "labels.json"),
                "--output", Path.Combine(root, "reviewed.json")]);
            if (expected == 2) {
                await Assert.ThrowsAsync<InvalidDataException>(() => EvaluationReview.RunAsync(options, default));
                Assert.False(File.Exists(options.OutputPath));
            } else {
                Assert.Equal(expected, await EvaluationReview.RunAsync(options, default));
                using var saved = JsonDocument.Parse(await File.ReadAllTextAsync(options.OutputPath!));
                Assert.Equal(expected == 0, saved.RootElement.GetProperty("qualityPassed").GetBoolean());
                await Assert.ThrowsAsync<IOException>(() => EvaluationReview.RunAsync(options, default));
            }
        } finally { Directory.Delete(root, true); }
    }
}
