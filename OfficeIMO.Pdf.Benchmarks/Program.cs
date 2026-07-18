using System.Diagnostics;
using System.Text.Json;
using OfficeIMO.Pdf;

PdfPerformanceBudget budget = JsonSerializer.Deserialize<PdfPerformanceBudget>(
    File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "pdf-performance-budgets.json")),
    new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
    ?? throw new InvalidOperationException("PDF performance budget manifest is invalid.");

byte[] corpus = CreateCorpus();
(PdfPerformanceMeasurement cold, PdfPerformanceMeasurement cached) = MeasureWorkflows(corpus);
double speedup = cold.ElapsedMilliseconds / Math.Max(cached.ElapsedMilliseconds, 0.001D);
double allocationReduction = cold.AllocatedBytes / (double)Math.Max(cached.AllocatedBytes, 1L);

Console.WriteLine($"Corpus: {corpus.Length:N0} bytes");
Print(cold);
Print(cached);
Console.WriteLine($"Cached speedup: {speedup:F2}x");
Console.WriteLine($"Cached allocation reduction: {allocationReduction:F2}x");

bool verify = args.Contains("--verify-budgets", StringComparer.OrdinalIgnoreCase);
if (!verify) {
    return 0;
}

var failures = new List<string>();
foreach (PdfPerformanceMeasurement measurement in new[] { cold, cached }) {
    if (measurement.ElapsedMilliseconds > budget.MaxElapsedMilliseconds) {
        failures.Add($"{measurement.Name}: {measurement.ElapsedMilliseconds:F1} ms exceeded {budget.MaxElapsedMilliseconds:F1} ms.");
    }

    if (measurement.AllocatedBytes > budget.MaxAllocatedBytes) {
        failures.Add($"{measurement.Name}: {measurement.AllocatedBytes:N0} allocated bytes exceeded {budget.MaxAllocatedBytes:N0}.");
    }
}

if (speedup < budget.MinimumCachedSpeedup) {
    failures.Add($"Cached workflow speedup {speedup:F2}x was below {budget.MinimumCachedSpeedup:F2}x.");
}

if (allocationReduction < budget.MinimumCachedAllocationReduction) {
    failures.Add($"Cached workflow allocation reduction {allocationReduction:F2}x was below {budget.MinimumCachedAllocationReduction:F2}x.");
}

foreach (string failure in failures) {
    Console.Error.WriteLine("BUDGET FAILURE: " + failure);
}

return failures.Count == 0 ? 0 : 1;

static byte[] CreateCorpus() {
    PdfDocument document = PdfDocument.Create(new PdfOptions {
        CompressContentStreams = true,
        DefaultFontSize = 10
    }).Meta(title: "OfficeIMO.Pdf performance corpus");

    for (int page = 1; page <= 60; page++) {
        document
            .H1("Operational report " + page)
            .Paragraph(paragraph => paragraph.Text(
                "This deterministic page exercises parsing, text extraction, inspection, diagnostics, " +
                "stream decoding, page-tree traversal, and one-call analysis through the public facade."))
            .Table(new[] {
                new[] { "Metric", "Value", "Status" },
                new[] { "Documents", (page * 37).ToString(), "Healthy" },
                new[] { "Rules", (page * 11).ToString(), "Reviewed" },
                new[] { "Signals", (page * 19).ToString(), "Observed" }
            });
        if (page < 60) {
            document.PageBreak();
        }
    }

    return document.ToBytes();
}

static long RunColdWorkflow(byte[] corpus) {
    string text = PdfDocument.Open(corpus).Read.Text();
    PdfDocumentInfo info = PdfDocument.Open(corpus).Inspect();
    PdfDocumentPreflight preflight = PdfDocument.Open(corpus).Preflight();
    PdfAnalysisReport analysis = PdfDocument.Open(corpus).Analyze();
    return text.Length + info.PageCount + preflight.Diagnostics.Count + analysis.Diagnostics.ObjectCount;
}

static long RunCachedWorkflow(byte[] corpus) {
    PdfDocument document = PdfDocument.Open(corpus);
    string text = document.Read.Text();
    PdfDocumentInfo info = document.Inspect();
    PdfDocumentPreflight preflight = document.Preflight();
    PdfAnalysisReport analysis = document.Analyze();
    return text.Length + info.PageCount + preflight.Diagnostics.Count + analysis.Diagnostics.ObjectCount;
}

static (PdfPerformanceMeasurement Cold, PdfPerformanceMeasurement Cached) MeasureWorkflows(byte[] corpus) {
    const int sampleCount = 7;
    RunColdWorkflow(corpus);
    RunCachedWorkflow(corpus);

    var coldSamples = new List<PdfPerformanceSample>(sampleCount);
    var cachedSamples = new List<PdfPerformanceSample>(sampleCount);
    for (int sample = 0; sample < sampleCount; sample++) {
        // Alternate order so sustained machine load cannot systematically favor either workflow.
        if (sample % 2 == 0) {
            coldSamples.Add(MeasureOnce(corpus, RunColdWorkflow));
            cachedSamples.Add(MeasureOnce(corpus, RunCachedWorkflow));
        } else {
            cachedSamples.Add(MeasureOnce(corpus, RunCachedWorkflow));
            coldSamples.Add(MeasureOnce(corpus, RunColdWorkflow));
        }
    }

    return (
        Summarize("cold-workflow", coldSamples),
        Summarize("cached-workflow", cachedSamples));
}

static PdfPerformanceSample MeasureOnce(byte[] corpus, Func<byte[], long> operation) {
    GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
    long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
    var stopwatch = Stopwatch.StartNew();
    long output = operation(corpus);
    stopwatch.Stop();
    long allocated = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;
    if (output <= 0) {
        throw new InvalidOperationException("PDF performance workflow produced no observable output.");
    }

    return new PdfPerformanceSample(stopwatch.Elapsed.TotalMilliseconds, allocated, output);
}

static PdfPerformanceMeasurement Summarize(string name, IReadOnlyList<PdfPerformanceSample> samples) {
    long output = samples[0].Output;
    if (samples.Any(sample => sample.Output != output)) {
        throw new InvalidOperationException(name + " produced inconsistent output between samples.");
    }

    double elapsed = samples
        .Select(sample => sample.ElapsedMilliseconds)
        .OrderBy(value => value)
        .ElementAt(samples.Count / 2);
    long allocated = samples
        .Select(sample => sample.AllocatedBytes)
        .OrderBy(value => value)
        .ElementAt(samples.Count / 2);
    return new PdfPerformanceMeasurement(name, elapsed, allocated, output);
}

static void Print(PdfPerformanceMeasurement measurement) {
    Console.WriteLine(
        $"{measurement.Name,-16} {measurement.ElapsedMilliseconds,8:F1} ms " +
        $"{measurement.AllocatedBytes / 1048576D,8:F1} MiB allocated");
}

internal sealed record PdfPerformanceBudget(
    double MaxElapsedMilliseconds,
    long MaxAllocatedBytes,
    double MinimumCachedSpeedup,
    double MinimumCachedAllocationReduction);

internal sealed record PdfPerformanceMeasurement(
    string Name,
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long Output);

internal sealed record PdfPerformanceSample(
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long Output);
