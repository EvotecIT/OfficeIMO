using System.Text;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.AI.Tests;

public sealed class EvaluationCaseRunnerTests {
    [Fact]
    public async Task ExportFailureIsRecordedWithoutInventingAnAnnotationOrStoppingTheNextCase() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-case-export-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            var request = new OfficeAiRequest { Operation = OfficeAiOperation.Ask, Instruction = "What is the value?" };
            var item = new EvaluationCase("fixture", ".txt", Encoding.UTF8.GetBytes("Value: 42."), request, false, "42",
                new EvaluationGold(FactMarkers: ["42"]));
            string failed = Path.Combine(root, "failed");
            Directory.CreateDirectory(Path.Combine(failed, "report.json")); // Reachable output-write failure after inference.
            var attempt = await EvaluationCaseRunner.ExecuteAsync(item, request, new Executor(), "source.txt", failed, default, default);
            Assert.NotNull(attempt.Result);
            Assert.False(attempt.ContractPassed);
            Assert.NotNull(attempt.Failure);
            Assert.Null(attempt.ReportSha256);
            var next = await EvaluationCaseRunner.ExecuteAsync(item, request, new Executor(), "source.txt", Path.Combine(root, "next"), default, default);
            Assert.True(next.ContractPassed);
            Assert.Null(next.Failure);
            Assert.Equal(64, next.ReportSha256!.Length);
            Assert.True(File.Exists(Path.Combine(root, "next", "report.json")));
        } finally { Directory.Delete(root, true); }
    }

    private sealed class Executor : IOfficeAiExecutor {
        public OfficeAiExecutionProfile Profile => new() { Id = "fixture", Provider = "fixture", Model = "fixture", IsLocal = true };
        public Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) {
            using var input = JsonDocument.Parse(request.InputJson);
            string id = input.RootElement.GetProperty("evidence")[0].GetProperty("id").GetString()!;
            return Task.FromResult(new OfficeAiExecutionResponse(JsonSerializer.Serialize(new {
                claims = new[] { new { text = "The value is 42.", evidence = new[] { new { id, quote = "42" } } } },
                fields = Array.Empty<object>(), blocks = Array.Empty<object>(), tables = Array.Empty<object>()
            })));
        }
    }
}
