using System.Diagnostics;
using System.Text.Json;
using OfficeIMO.AI;

internal static class EvaluationRunner {
    public static async Task<int> RunAsync(ExampleOptions options, CancellationToken cancellationToken) {
        string output = options.OutputPath ?? throw new ArgumentException("Evaluation requires an output directory.");
        if (Directory.Exists(output)) throw new IOException("Evaluation output directory already exists.");
        Directory.CreateDirectory(output);
        using var deadline = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
        deadline.CancelAfter(TimeSpan.FromHours(1));
        IReadOnlyList<EvaluationCase> selected = EvaluationCorpus.Create().Where(item => options.Split == "all" || item.Split == options.Split).ToArray();
        var excluded = options.TextOnly ? selected.Where(item => item.Images).Select(item => item.Id).ToArray() : Array.Empty<string>();
        IReadOnlyList<EvaluationCase> cases = selected.Where(item => !options.TextOnly || !item.Images).ToArray();
        if (options.CaseId is not null) cases = new[] { cases.Single(item => item.Id == options.CaseId) };
        using var executor = await ExampleExecution.ConnectAsync(options, images: cases.Any(item => item.Images), deadline.Token);
        var rows = new List<object>();
        int passed = 0;
        foreach (EvaluationCase item in cases) {
            for (int repetition = 1; repetition <= options.Repeat; repetition++) {
                deadline.Token.ThrowIfCancellationRequested();
                string directory = Path.Combine(output, item.Id, "run-" + repetition);
                Directory.CreateDirectory(directory);
                string sourcePath = Path.Combine(directory, "source" + item.Extension);
                await File.WriteAllBytesAsync(sourcePath, item.Source, deadline.Token);
                OfficeAiRequest request = item.Request with { AllowRemoteProcessing = options.AllowRemote, IncludeImages = item.Images };
                using var operation = CancellationTokenSource.CreateLinkedTokenSource(deadline.Token);
                operation.CancelAfter(request.Limits.Timeout);
                var timer = Stopwatch.StartNew();
                long allocatedBefore = GC.GetTotalAllocatedBytes();
                using var process = Process.GetCurrentProcess();
                long peakWorkingSet = process.WorkingSet64;
                using var sampler = new Timer(_ => {
                    try { process.Refresh(); RecordMaximum(ref peakWorkingSet, process.WorkingSet64); }
                    catch (InvalidOperationException) { }
                }, null, 0, 100);
                var recorded = new RecordingExecutor(executor);
                OfficeAiResult? result = null;
                string? failure = null;
                EvaluationScore? score = null;
                bool accepted = false;
                try {
                    OfficeAiDocument document = await DocumentInputs.ReadAsync(item.Source, sourcePath, item.Images, request.Pages, request.Limits, operation.Token);
                    result = await new OfficeAiEngine(recorded).RunAsync(document, request, cancellationToken: operation.Token);
                    await ArtifactWriter.SaveAsync(directory, document, result, operation.Token);
                    score = item.Gold.Score(result);
                    accepted = result.Status != OfficeAiResultStatus.InvalidResponse && score.Passed;
                } catch (OperationCanceledException) when (!deadline.IsCancellationRequested) { failure = "case-timeout"; }
                  catch (OperationCanceledException) { throw; }
                  catch (Exception error) when (error is not OutOfMemoryException) {
                    failure = error.GetType().Name;
                }
                await sampler.DisposeAsync();
                await File.WriteAllTextAsync(Path.Combine(directory, "provider-responses.json"), JsonSerializer.Serialize(recorded.Responses), deadline.Token);
                if (accepted) passed++;
                rows.Add(new {
                    item.Id, item.Split, repetition, item.Expected, gold = item.Gold, score, passed = accepted,
                    status = result?.Status.ToString(), failure,
                    sourceHash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(item.Source)).ToLowerInvariant(),
                    sourceBytes = item.Source.Length, elapsedMilliseconds = timer.ElapsedMilliseconds,
                    sampledEvaluationProcessPeakWorkingSetBytes = peakWorkingSet,
                    evaluationProcessManagedAllocationDeltaBytes = GC.GetTotalAllocatedBytes() - allocatedBefore,
                    requestAttempts = recorded.Attempts, maxMeasuredRequestCharacters = recorded.MaximumRequestCharacters,
                    result?.InputTokens, result?.OutputTokens, result?.Diagnostics, synthesis = result?.SynthesisStatus.ToString(),
                    processed = result?.ProcessedEvidenceIds.Count, omitted = result?.OmittedEvidenceIds.Count,
                    processedCharacters = result?.ProcessedTextRanges.Sum(range => (long)range.Length),
                    emptyPages = result?.EmptyPages.Count, claims = result?.Claims.Count, fields = result?.Fields.Count, tables = result?.Tables.Count
                });
                Console.WriteLine($"{item.Id} [{repetition}/{options.Repeat}]: {(accepted ? "PASS" : "FAIL")} ({result?.Status.ToString() ?? failure})");
                var report = new {
                    schema = "officeimo.ai.evaluation.v2", corpus = EvaluationCorpus.Version, utc = DateTimeOffset.UtcNow,
                    profile = executor.Profile, options.Split, options.Repeat, excludedImageCases = excluded,
                    runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
                    platform = System.Runtime.InteropServices.RuntimeInformation.OSDescription,
                    assembly = typeof(OfficeAiEngine).Assembly.GetName().Version?.ToString(),
                    threshold = "Every selected repetition must meet declared gold. Fact markers are recall smoke checks, not semantic entailment. Review source/report pairs separately.",
                    memoryScope = "Sampled evaluation process only; excludes model server, GPU and child processes. Allocation delta includes process background work.",
                    passed, completed = rows.Count, total = cases.Count * options.Repeat, cases = rows
                };
                await File.WriteAllTextAsync(Path.Combine(output, "evaluation.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }), deadline.Token);
            }
        }
        return passed == cases.Count * options.Repeat ? 0 : 1;
    }

    private static void RecordMaximum(ref long target, long value) {
        long previous;
        do { previous = Interlocked.Read(ref target); if (previous >= value) return; }
        while (Interlocked.CompareExchange(ref target, value, previous) != previous);
    }

    // Recording is confined to the explicitly selected synthetic corpus, never general file processing.
    private sealed class RecordingExecutor(IOfficeAiExecutor inner) : IOfficeAiExecutor {
        public OfficeAiExecutionProfile Profile => inner.Profile;
        public List<OfficeAiExecutionResponse> Responses { get; } = new();
        public int Attempts { get; private set; }
        public int MaximumRequestCharacters { get; private set; }
        public int MeasureRequestCharacters(OfficeAiExecutionRequest request) => inner.MeasureRequestCharacters(request);
        public async Task<OfficeAiExecutionResponse> ExecuteAsync(OfficeAiExecutionRequest request, CancellationToken cancellationToken = default) {
            Attempts++; MaximumRequestCharacters = Math.Max(MaximumRequestCharacters, MeasureRequestCharacters(request));
            OfficeAiExecutionResponse response = await inner.ExecuteAsync(request, cancellationToken);
            Responses.Add(response); return response;
        }
    }
}
