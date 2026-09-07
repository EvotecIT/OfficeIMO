using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.AI;

internal static class EvaluationRunner {
    public static async Task<int> RunAsync(ExampleOptions options, CancellationToken cancellationToken) {
        string output = options.OutputPath ?? throw new ArgumentException("Evaluation requires an output directory.");
        if (Directory.Exists(output)) throw new IOException("Evaluation output directory already exists.");
        Directory.CreateDirectory(output);
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        deadline.CancelAfter(TimeSpan.FromMinutes(15));
        IReadOnlyList<EvaluationCase> cases = EvaluationCorpus.Create();
        if (options.CaseId is not null) cases = new[] { cases.Single(item => item.Id == options.CaseId) };
        using var executor = await ExampleExecution.ConnectAsync(options, images: true, deadline.Token);
        var rows = new List<object>();
        int passed = 0;
        foreach (EvaluationCase item in cases) {
            deadline.Token.ThrowIfCancellationRequested();
            string directory = Path.Combine(output, item.Id);
            Directory.CreateDirectory(directory);
            string sourcePath = Path.Combine(directory, "source" + item.Extension);
            await File.WriteAllBytesAsync(sourcePath, item.Source, deadline.Token);
            OfficeAiRequest request = item.Request with { AllowRemoteProcessing = options.AllowRemote, IncludeImages = item.Images };
            using var operation = CancellationTokenSource.CreateLinkedTokenSource(deadline.Token);
            operation.CancelAfter(request.Limits.Timeout);
            var timer = Stopwatch.StartNew();
            OfficeAiDocument document = await DocumentInputs.ReadAsync(item.Source, sourcePath, item.Images, request.Pages, request.Limits, operation.Token);
            var recorded = new RecordingExecutor(executor);
            var engine = new OfficeAiEngine(recorded);
            OfficeAiResult result = await engine.RunAsync(document, request, cancellationToken: operation.Token);
            await File.WriteAllTextAsync(Path.Combine(directory, "provider-responses.json"), JsonSerializer.Serialize(recorded.Responses), operation.Token);
            ArtifactWriter.Save(directory, document, result);
            bool accepted = result.Status != OfficeAiResultStatus.InvalidResponse && item.Check(result);
            if (accepted) passed++;
            rows.Add(new {
                item.Id, item.Expected, passed = accepted, status = result.Status.ToString(), document.SourceHash,
                sourceBytes = item.Source.Length, elapsedMilliseconds = timer.ElapsedMilliseconds,
                result.InputTokens, result.OutputTokens, result.Diagnostics,
                processed = result.ProcessedEvidenceIds.Count, omitted = result.OmittedEvidenceIds.Count,
                emptyPages = result.EmptyPages.Count, claims = result.Claims.Count, fields = result.Fields.Count, tables = result.Tables.Count
            });
            Console.WriteLine($"{item.Id}: {(accepted ? "PASS" : "FAIL")} ({result.Status})");
        }
        var report = new {
            schema = "officeimo.ai.evaluation.v1", corpus = EvaluationCorpus.Version, utc = DateTimeOffset.UtcNow,
            profile = executor.Profile, runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            platform = System.Runtime.InteropServices.RuntimeInformation.OSDescription,
            assembly = typeof(OfficeAiEngine).Assembly.GetName().Version?.ToString(),
            threshold = "Every case must meet its exact field/table/status assertion. Claim assertions are smoke checks; semantic support requires separate review.",
            passed, total = cases.Count, cases = rows
        };
        await File.WriteAllTextAsync(Path.Combine(output, "evaluation.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }), deadline.Token);
        return passed == cases.Count ? 0 : 1;
    }

    // Recording is confined to the explicitly selected synthetic corpus, never the general file-processing path.
    private sealed class RecordingExecutor(IOfficeAiExecutor inner) : IOfficeAiExecutor {
        public OfficeAiExecutionProfile Profile => inner.Profile;
        public List<OfficeAiExecutionResponse> Responses { get; } = new();
        public int MeasureRequestCharacters(OfficeAiExecutionRequest request) => inner.MeasureRequestCharacters(request);
        public async Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) {
            OfficeAiExecutionResponse response = await inner.ExecuteAsync(request, cancellationToken);
            Responses.Add(response); return response;
        }
    }
}
