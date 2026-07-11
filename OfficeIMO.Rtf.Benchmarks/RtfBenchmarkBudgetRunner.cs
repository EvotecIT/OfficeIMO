using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Text.Json;
using OfficeIMO.Html;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Rtf;
using OfficeIMO.Rtf.Markdown;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;

namespace OfficeIMO.Rtf.Benchmarks;

internal static class RtfBenchmarkBudgetRunner {
    private static readonly JsonSerializerOptions JsonOptions = new JsonSerializerOptions {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    private static readonly string[] Operations = {
        "Parse", "Lossless", "SemanticWrite", "Html", "Markdown", "Pdf", "WordModel", "Word", "Reader"
    };

    public static int RunProbe(string[] args) {
        if (args.Length != 2 || !Operations.Contains(args[0], StringComparer.OrdinalIgnoreCase)) {
            Console.Error.WriteLine("Usage: --probe <Parse|Lossless|SemanticWrite|Html|Markdown|Pdf|WordModel|Word|Reader> <Small|Medium|Large>");
            return 2;
        }

        try {
            RtfBenchmarkMeasurement measurement = Measure(args[0], args[1]);
            Console.WriteLine(JsonSerializer.Serialize(measurement, JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    public static int Verify(string[] args) {
        string? scaleFilter = GetOption(args, "--scale");
        string? jsonPath = GetOption(args, "--json");
        IReadOnlyList<string> scales = string.IsNullOrWhiteSpace(scaleFilter)
            ? RtfBenchmarkCorpus.Scales
            : new[] { RtfBenchmarkCorpus.Get(scaleFilter!).Scale };
        RtfBenchmarkBudgetManifest manifest = LoadManifest();
        var measurements = new List<RtfBenchmarkMeasurement>();
        var failures = new List<string>();

        foreach (string scale in scales) {
            RtfBenchmarkScaleBudget? scaleBudget = manifest.Scales.FirstOrDefault(item =>
                string.Equals(item.Scale, scale, StringComparison.OrdinalIgnoreCase));
            foreach (string operation in Operations) {
                RtfBenchmarkMeasurement measurement = RunChildProbe(operation, scale);
                measurements.Add(measurement);
                RtfBenchmarkBudget? budget = manifest.Budgets.FirstOrDefault(item =>
                    string.Equals(item.Operation, operation, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(item.Scale, scale, StringComparison.OrdinalIgnoreCase));
                if (budget == null) {
                    failures.Add($"Missing budget for {operation}/{scale}.");
                    continue;
                }

                if (scaleBudget == null || measurement.InputBytes < scaleBudget.MinInputBytes || measurement.InputBytes > scaleBudget.MaxInputBytes) {
                    failures.Add($"{operation}/{scale}: input size {measurement.InputBytes} is outside the declared scale range.");
                }
                if (measurement.ElapsedMilliseconds > budget.MaxElapsedMilliseconds) failures.Add($"{operation}/{scale}: elapsed {measurement.ElapsedMilliseconds:F1} ms > {budget.MaxElapsedMilliseconds} ms.");
                if (measurement.AllocatedBytes > budget.MaxAllocatedBytes) failures.Add($"{operation}/{scale}: allocations {measurement.AllocatedBytes} > {budget.MaxAllocatedBytes} bytes.");
                if (measurement.PeakWorkingSetBytes > budget.MaxPeakWorkingSetBytes) failures.Add($"{operation}/{scale}: peak working set {measurement.PeakWorkingSetBytes} > {budget.MaxPeakWorkingSetBytes} bytes.");
                if (measurement.OutputBytes > budget.MaxOutputBytes) failures.Add($"{operation}/{scale}: output {measurement.OutputBytes} > {budget.MaxOutputBytes} bytes.");
                Console.WriteLine($"{operation,-13} {scale,-6} {measurement.ElapsedMilliseconds,9:F1} ms {measurement.MebibytesPerSecond,9:F2} MiB/s {measurement.AllocatedBytes / 1048576d,9:F2} MiB alloc {measurement.PeakWorkingSetBytes / 1048576d,9:F2} MiB peak");
            }
        }

        var report = new RtfBenchmarkBudgetReport(DateTimeOffset.UtcNow, measurements, failures);
        if (!string.IsNullOrWhiteSpace(jsonPath)) {
            string fullPath = Path.GetFullPath(jsonPath!);
            Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
            File.WriteAllText(fullPath, JsonSerializer.Serialize(report, JsonOptions));
        }

        foreach (string failure in failures) Console.Error.WriteLine("BUDGET FAILURE: " + failure);
        return failures.Count == 0 ? 0 : 1;
    }

    private static RtfBenchmarkMeasurement Measure(string operation, string scale) {
        RtfBenchmarkFixture fixture = RtfBenchmarkCorpus.Get(scale);
        RtfDocument? preparedDocument = string.Equals(operation, "Parse", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(operation, "Lossless", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(operation, "SemanticWrite", StringComparison.OrdinalIgnoreCase)
            ? null
            : RtfDocument.Read(fixture.Rtf).Document;
        RtfPdfSaveOptions? preparedPdfOptions = string.Equals(operation, "Pdf", StringComparison.OrdinalIgnoreCase)
            ? RtfBenchmarkSupport.CreatePdfSaveOptions()
            : null;
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        long outputBytes;

        if (string.Equals(operation, "Parse", StringComparison.OrdinalIgnoreCase)) {
            RtfReadResult result = RtfDocument.Read(fixture.Rtf);
            outputBytes = Encoding.UTF8.GetByteCount(result.ToRtfLossless());
        } else if (string.Equals(operation, "Lossless", StringComparison.OrdinalIgnoreCase)) {
            outputBytes = Encoding.UTF8.GetByteCount(RtfDocument.Read(fixture.Rtf).EditLossless().ToRtf());
        } else if (string.Equals(operation, "SemanticWrite", StringComparison.OrdinalIgnoreCase)) {
            outputBytes = Encoding.UTF8.GetByteCount(RtfDocument.Read(fixture.Rtf).Document.ToRtf());
        } else if (string.Equals(operation, "Html", StringComparison.OrdinalIgnoreCase)) {
            outputBytes = Encoding.UTF8.GetByteCount(preparedDocument!.ToHtml());
        } else if (string.Equals(operation, "Markdown", StringComparison.OrdinalIgnoreCase)) {
            outputBytes = Encoding.UTF8.GetByteCount(preparedDocument!.ToMarkdown());
        } else if (string.Equals(operation, "Pdf", StringComparison.OrdinalIgnoreCase)) {
            outputBytes = preparedDocument!.ToPdfDocument(preparedPdfOptions).ToBytes().LongLength;
        } else if (string.Equals(operation, "WordModel", StringComparison.OrdinalIgnoreCase)) {
            using WordDocument word = preparedDocument!.ToWordDocument();
            outputBytes = 0;
        } else if (string.Equals(operation, "Word", StringComparison.OrdinalIgnoreCase)) {
            using WordDocument word = preparedDocument!.ToWordDocument();
            using MemoryStream stream = word.SaveAsMemoryStream();
            outputBytes = stream.Length;
        } else {
            ReaderChunk[] chunks = DocumentReaderRtfExtensions.ReadRtfDocument(preparedDocument!).ToArray();
            outputBytes = chunks.Sum(chunk => Encoding.UTF8.GetByteCount(chunk.Markdown ?? chunk.Text ?? string.Empty));
        }

        stopwatch.Stop();
        long allocatedBytes = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        double seconds = Math.Max(stopwatch.Elapsed.TotalSeconds, 0.000001d);
        return new RtfBenchmarkMeasurement(
            operation,
            fixture.Scale,
            fixture.InputBytes,
            outputBytes,
            stopwatch.Elapsed.TotalMilliseconds,
            allocatedBytes,
            process.PeakWorkingSet64,
            fixture.InputBytes / 1048576d / seconds);
    }

    private static RtfBenchmarkMeasurement RunChildProbe(string operation, string scale) {
        string processPath = Environment.ProcessPath ?? throw new InvalidOperationException("Unable to resolve benchmark process path.");
        var startInfo = new ProcessStartInfo {
            FileName = processPath,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        if (string.Equals(Path.GetFileNameWithoutExtension(processPath), "dotnet", StringComparison.OrdinalIgnoreCase)) {
            startInfo.ArgumentList.Add(Assembly.GetEntryAssembly()!.Location);
        }
        startInfo.ArgumentList.Add("--probe");
        startInfo.ArgumentList.Add(operation);
        startInfo.ArgumentList.Add(scale);
        using Process child = Process.Start(startInfo) ?? throw new InvalidOperationException("Unable to start benchmark probe process.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) throw new InvalidOperationException($"Probe {operation}/{scale} failed: {error}");
        return JsonSerializer.Deserialize<RtfBenchmarkMeasurement>(output, JsonOptions)
            ?? throw new InvalidOperationException($"Probe {operation}/{scale} returned no measurement.");
    }

    private static RtfBenchmarkBudgetManifest LoadManifest() {
        string path = Path.Combine(AppContext.BaseDirectory, "rtf-benchmark-budgets.json");
        return JsonSerializer.Deserialize<RtfBenchmarkBudgetManifest>(File.ReadAllText(path), JsonOptions)
            ?? throw new InvalidOperationException("RTF benchmark budget manifest is invalid.");
    }

    private static string? GetOption(string[] args, string name) {
        int index = Array.FindIndex(args, item => string.Equals(item, name, StringComparison.OrdinalIgnoreCase));
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }
}

internal sealed record RtfBenchmarkMeasurement(
    string Operation,
    string Scale,
    long InputBytes,
    long OutputBytes,
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long PeakWorkingSetBytes,
    double MebibytesPerSecond);

internal sealed record RtfBenchmarkBudgetReport(
    DateTimeOffset MeasuredAtUtc,
    IReadOnlyList<RtfBenchmarkMeasurement> Measurements,
    IReadOnlyList<string> Failures);

internal sealed class RtfBenchmarkBudgetManifest {
    public int Version { get; set; }
    public string Description { get; set; } = string.Empty;
    public List<RtfBenchmarkScaleBudget> Scales { get; set; } = new List<RtfBenchmarkScaleBudget>();
    public List<RtfBenchmarkBudget> Budgets { get; set; } = new List<RtfBenchmarkBudget>();
}

internal sealed class RtfBenchmarkScaleBudget {
    public string Scale { get; set; } = string.Empty;
    public long MinInputBytes { get; set; }
    public long MaxInputBytes { get; set; }
}

internal sealed class RtfBenchmarkBudget {
    public string Operation { get; set; } = string.Empty;
    public string Scale { get; set; } = string.Empty;
    public double MaxElapsedMilliseconds { get; set; }
    public long MaxAllocatedBytes { get; set; }
    public long MaxPeakWorkingSetBytes { get; set; }
    public long MaxOutputBytes { get; set; }
}
