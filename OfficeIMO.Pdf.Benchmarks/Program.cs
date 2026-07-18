using System.Diagnostics;
using System.Text.Json;
using OfficeIMO.Pdf;

PdfPerformanceBudget budget = JsonSerializer.Deserialize<PdfPerformanceBudget>(
    File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "pdf-performance-budgets.json")),
    new JsonSerializerOptions { PropertyNameCaseInsensitive = true })
    ?? throw new InvalidOperationException("PDF performance budget manifest is invalid.");

byte[] corpus = CreateCorpus();
PdfPerformanceMeasurement cold = Measure("cold-workflow", corpus, RunColdWorkflow);
PdfPerformanceMeasurement cached = Measure("cached-workflow", corpus, RunCachedWorkflow);
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

static PdfPerformanceMeasurement Measure(string name, byte[] corpus, Func<byte[], long> operation) {
    operation(corpus);
    GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
    long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
    var stopwatch = Stopwatch.StartNew();
    long output = operation(corpus);
    stopwatch.Stop();
    long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
    if (output <= 0) {
        throw new InvalidOperationException(name + " produced no observable output.");
    }

    return new PdfPerformanceMeasurement(name, stopwatch.Elapsed.TotalMilliseconds, allocated, output);
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
