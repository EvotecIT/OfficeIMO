using System.Diagnostics;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json;
using Microsoft.Playwright;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PeachPDF;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Measures whole-corpus HTML-to-PDF cost per engine in isolated worker processes.
/// Fidelity analysis stays in HeadToHead/h2h.py so one independent rasteriser judges every engine.
/// </summary>
internal static class HtmlHeadToHeadRunner {
    private static readonly string[] AllEngines = { "OfficeIMO", "PeachPDF", "PdfPecker", "Chromium" };
    private static readonly JsonSerializerOptions Json = new() { WriteIndented = true, PropertyNamingPolicy = JsonNamingPolicy.CamelCase };

    internal static async Task<int> RunAsync(string[] args) {
        string corpus = Path.GetFullPath(Required(args, "--corpus"));
        string output = Path.GetFullPath(Required(args, "--output"));
        string[] engines = (Option(args, "--engines")?.Split(',') ?? AllEngines);
        string[]? selected = Option(args, "--docs")?.Split(',');
        int coldRuns = int.Parse(Option(args, "--cold") ?? "3", CultureInfo.InvariantCulture);
        int warmIterations = int.Parse(Option(args, "--warm") ?? "15", CultureInfo.InvariantCulture);
        int warmSeconds = int.Parse(Option(args, "--warm-seconds") ?? "10", CultureInfo.InvariantCulture);
        // Tiered JIT promotes a method after about 30 calls, so a slow engine measured after three
        // warm-ups would still run partly unoptimised code while a fast one is fully optimised.
        int warmupIterations = int.Parse(Option(args, "--warmup") ?? "30", CultureInfo.InvariantCulture);
        int warmupSeconds = int.Parse(Option(args, "--warmup-seconds") ?? "20", CultureInfo.InvariantCulture);
        int degree = int.Parse(Option(args, "--parallel") ?? "4", CultureInfo.InvariantCulture);
        int timeoutSeconds = int.Parse(Option(args, "--timeout") ?? "120", CultureInfo.InvariantCulture);
        HashSet<string> parallelDocuments = new((Option(args, "--parallel-docs") ?? "invoice,tour,report-30p").Split(','));
        Directory.CreateDirectory(output);
        string[] documents = Directory.GetFiles(corpus, "*.html")
            .Where(path => selected == null || selected.Contains(Path.GetFileNameWithoutExtension(path)))
            .OrderBy(path => path, StringComparer.Ordinal).ToArray();

        var records = new List<WorkerRecord>();
        string resultsPath = Path.Combine(output, "measurements.jsonl");
        await using var results = new StreamWriter(resultsPath, append: true);
        foreach (string document in documents) {
            string name = Path.GetFileNameWithoutExtension(document);
            foreach (string engine in engines) {
                string engineDirectory = Path.Combine(output, engine);
                Directory.CreateDirectory(engineDirectory);
                string pdf = Path.Combine(engineDirectory, name + ".pdf");
                bool failedCold = false;
                for (int run = 1; run <= coldRuns && !failedCold; run++) {
                    WorkerRecord record = await RunWorkerAsync(engine, document, pdf, "cold", 1, warmSeconds, degree, timeoutSeconds, warmupIterations, warmupSeconds);
                    failedCold = record.Status != "Passed";
                    await Append(results, name, engine, record);
                }
                if (failedCold) continue;
                await Append(results, name, engine,
                    await RunWorkerAsync(engine, document, pdf, "warm", warmIterations, warmSeconds, degree, timeoutSeconds, warmupIterations, warmupSeconds));
                if (parallelDocuments.Contains(name)) {
                    await Append(results, name, engine,
                        await RunWorkerAsync(engine, document, pdf, "parallel", warmIterations, warmSeconds, degree, timeoutSeconds, warmupIterations, warmupSeconds));
                }
            }
        }
        Console.WriteLine("HTML_HEAD_TO_HEAD_MEASUREMENTS=" + resultsPath);
        return 0;
    }

    private static async Task Append(StreamWriter results, string document, string engine, WorkerRecord record) {
        await results.WriteLineAsync(JsonSerializer.Serialize(record with { Document = document, Engine = engine },
            new JsonSerializerOptions(JsonSerializerDefaults.Web)));
        await results.FlushAsync();
        Console.WriteLine($"{document,-24} {engine,-10} {record.Mode,-8} {record.Status,-8} first={record.FirstMilliseconds:0.0}ms/{record.FirstCpuMilliseconds:0}cpu warm={record.WarmMedianMilliseconds:0.0}ms/{record.WarmCpuMedianMilliseconds:0.0}cpu {record.Detail}");
    }

    private static async Task<WorkerRecord> RunWorkerAsync(string engine, string html, string pdf, string mode,
        int iterations, int warmSeconds, int degree, int timeoutSeconds, int warmupIterations, int warmupSeconds) {
        string resultPath = Path.GetTempFileName();
        string entry = System.Reflection.Assembly.GetEntryAssembly()!.Location;
        string host = Environment.ProcessPath ?? "dotnet";
        var start = new ProcessStartInfo(host) {
            UseShellExecute = false, RedirectStandardError = true, RedirectStandardOutput = true, CreateNoWindow = true
        };
        if (Path.GetFileNameWithoutExtension(host) == "dotnet") start.ArgumentList.Add(entry);
        foreach (string argument in new[] { "html-head-to-head-worker", "--engine", engine, "--html", html, "--pdf", pdf, "--mode", mode,
                     "--iterations", iterations.ToString(CultureInfo.InvariantCulture), "--seconds", warmSeconds.ToString(CultureInfo.InvariantCulture),
                     "--degree", degree.ToString(CultureInfo.InvariantCulture), "--result", resultPath,
                     "--warmup", warmupIterations.ToString(CultureInfo.InvariantCulture),
                     "--warmup-seconds", warmupSeconds.ToString(CultureInfo.InvariantCulture) }) {
            start.ArgumentList.Add(argument);
        }
        var wall = Stopwatch.StartNew();
        using Process process = Process.Start(start) ?? throw new InvalidOperationException("Could not start worker.");
        // Only the Chromium path has helper processes worth the cost of ps sampling on a busy host.
        ProcessTreeMemorySampler? sampler = engine == "Chromium" ? new ProcessTreeMemorySampler(process) : null;
        Task<string> error = process.StandardError.ReadToEndAsync();
        Task<string> outputText = process.StandardOutput.ReadToEndAsync();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(timeoutSeconds));
        string status = "Passed";
        string detail = "";
        try {
            await process.WaitForExitAsync(timeout.Token);
            if (process.ExitCode != 0) {
                status = "Failed";
                string text = await error + await outputText;
                detail = text.Split('\n').FirstOrDefault(line => line.Length > 0)?.Trim() ?? "exit " + process.ExitCode;
            }
        } catch (OperationCanceledException) {
            status = "Timeout";
            detail = "Exceeded " + timeoutSeconds + " s";
            process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync();
        }
        wall.Stop();
        long treePeak = 0;
        if (sampler != null) {
            await sampler.StopAsync();
            treePeak = sampler.CreateEvidence().PeakWorkingSetBytes;
            await sampler.DisposeAsync();
        }
        WorkerRecord record = status == "Passed" && new FileInfo(resultPath).Length > 0
            ? JsonSerializer.Deserialize<WorkerRecord>(File.ReadAllText(resultPath), new JsonSerializerOptions(JsonSerializerDefaults.Web))!
            : new WorkerRecord { Mode = mode };
        File.Delete(resultPath);
        return record with { Status = status, Detail = detail.Length > 300 ? detail[..300] : detail, ProcessWallMilliseconds = wall.Elapsed.TotalMilliseconds,
            PeakRssBytes = Math.Max(record.PeakRssBytes, treePeak) };
    }

    internal static async Task<int> RunWorkerAsync(string[] args) {
        string engine = Required(args, "--engine");
        string html = await File.ReadAllTextAsync(Required(args, "--html"));
        string pdfPath = Required(args, "--pdf");
        string mode = Required(args, "--mode");
        int iterations = int.Parse(Required(args, "--iterations"), CultureInfo.InvariantCulture);
        int seconds = int.Parse(Required(args, "--seconds"), CultureInfo.InvariantCulture);
        int degree = int.Parse(Required(args, "--degree"), CultureInfo.InvariantCulture);
        int warmupIterations = int.Parse(Option(args, "--warmup") ?? "30", CultureInfo.InvariantCulture);
        int warmupSeconds = int.Parse(Option(args, "--warmup-seconds") ?? "20", CultureInfo.InvariantCulture);
        var record = new WorkerRecord { Mode = mode };

        using Process self = Process.GetCurrentProcess();
        double Cpu() { self.Refresh(); return self.TotalProcessorTime.TotalMilliseconds; }
        double cpuBefore = Cpu();
        // The first document includes browser start-up where the engine has one.
        var first = Stopwatch.StartNew();
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        Renderer renderer = await Renderer.CreateAsync(engine);
        byte[] bytes = await renderer.RenderAsync(html);
        first.Stop();
        record = record with { FirstMilliseconds = first.Elapsed.TotalMilliseconds, FirstCpuMilliseconds = Cpu() - cpuBefore, FirstAllocatedBytes = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore, OutputBytes = bytes.Length };
        RequirePdf(bytes, engine, "first");
        await File.WriteAllBytesAsync(pdfPath, bytes);

        async Task<int> WarmUpAsync() {
            var warming = Stopwatch.StartNew();
            int warmups = 0;
            while (warmups < 3 || (warmups < warmupIterations && warming.Elapsed.TotalSeconds < warmupSeconds)) {
                RequirePdf(await renderer.RenderAsync(html), engine, "warm-up");
                warmups++;
            }
            return warmups;
        }

        if (mode == "warm") {
            record = record with { WarmupIterations = await WarmUpAsync() };
            var times = new List<double>();
            var allocations = new List<long>();
            var cpu = new List<double>();
            var budget = Stopwatch.StartNew();
            int gen2Before = GC.CollectionCount(2);
            while (times.Count < 3 || (times.Count < iterations && budget.Elapsed.TotalSeconds < seconds)) {
                long before = GC.GetTotalAllocatedBytes(precise: true);
                double cpuStart = Cpu();
                var one = Stopwatch.StartNew();
                byte[] warm = await renderer.RenderAsync(html);
                one.Stop();
                RequirePdf(warm, engine, "warm");
                cpu.Add(Cpu() - cpuStart);
                allocations.Add(GC.GetTotalAllocatedBytes(precise: true) - before);
                times.Add(one.Elapsed.TotalMilliseconds);
            }
            times.Sort();
            allocations.Sort();
            cpu.Sort();
            record = record with { WarmCpuMedianMilliseconds = cpu[cpu.Count / 2], WarmMedianMilliseconds = times[times.Count / 2], WarmMinimumMilliseconds = times[0], WarmIterations = times.Count,
                WarmAllocatedBytes = allocations[allocations.Count / 2], Gen2Collections = GC.CollectionCount(2) - gen2Before };
        } else if (mode == "parallel") {
            record = record with { WarmupIterations = await WarmUpAsync() };
            var fleet = await RendererFleet.CreateAsync(engine, renderer, degree);
            var clock = Stopwatch.StartNew();
            int completed = 0;
            await Task.WhenAll(Enumerable.Range(0, degree).Select(worker => Task.Run(async () => {
                while (clock.Elapsed.TotalSeconds < seconds) {
                    await fleet[worker].RenderAsync(html);
                    Interlocked.Increment(ref completed);
                }
            })));
            clock.Stop();
            record = record with { ParallelDegree = degree, ParallelDocumentsPerSecond = completed / clock.Elapsed.TotalSeconds };
        }
        await File.WriteAllTextAsync(Required(args, "--result"),
            JsonSerializer.Serialize(record with { PeakRssBytes = NativePeakRss.Read() }, new JsonSerializerOptions(JsonSerializerDefaults.Web)));
        // Results are durable before teardown: a browser that will not close must not void a finished measurement.
        try {
            await renderer.DisposeAsync().AsTask().WaitAsync(TimeSpan.FromSeconds(3));
        } catch (TimeoutException) {
            Environment.Exit(0);
        }
        return 0;
    }

    // Every timed render must still produce a complete PDF; a failed or truncated result cannot count as a fast one.
    private static void RequirePdf(byte[] bytes, string engine, string phase) {
        static bool StartsWithPdf(byte[] value) => value.Length > 8 && value[0] == (byte)'%' && value[1] == (byte)'P'
            && value[2] == (byte)'D' && value[3] == (byte)'F' && value[4] == (byte)'-';
        bool endsWithEof = bytes.Length > 8 && System.Text.Encoding.ASCII.GetString(bytes, bytes.Length - Math.Min(bytes.Length, 32), Math.Min(bytes.Length, 32)).Contains("%%EOF", StringComparison.Ordinal);
        if (!StartsWithPdf(bytes) || !endsWithEof) throw new InvalidDataException(engine + " " + phase + " render did not return a complete PDF.");
    }

    /// <summary>Splits OfficeIMO's warm cost into parse, layout alone and the full conversion with writing.</summary>
    internal static int WriteOfficeImoStages(string[] args) {
        string corpus = Path.GetFullPath(Required(args, "--corpus"));
        foreach (string name in Required(args, "--docs").Split(',')) {
            string html = File.ReadAllText(Path.Combine(corpus, name + ".html"));
            var samples = new Dictionary<string, List<(double Milliseconds, long Bytes)>>();
            T Measure<T>(string stage, Func<T> work) {
                long before = GC.GetTotalAllocatedBytes(precise: true);
                var clock = Stopwatch.StartNew();
                T value = work();
                clock.Stop();
                if (!samples.TryGetValue(stage, out var list)) samples[stage] = list = new();
                list.Add((clock.Elapsed.TotalMilliseconds, GC.GetTotalAllocatedBytes(precise: true) - before));
                return value;
            }
            int iterations = int.Parse(Option(args, "--iterations") ?? "15", CultureInfo.InvariantCulture);
            int warmups = int.Parse(Option(args, "--warmup") ?? "30", CultureInfo.InvariantCulture);
            for (int iteration = 0; iteration < warmups + iterations; iteration++) {
                HtmlConversionDocument source = Measure("parse", () => HtmlConversionDocument.Parse(html));
                Measure("layout-only", () => HtmlRenderEngine.Render(source, new HtmlToPdfOptions()));
                source = HtmlConversionDocument.Parse(html);
                // ToPdf serialises once to report its pages, so this stage includes writing and
                // ToBytes afterwards only copies the retained output.
                using OfficeIMO.Pdf.PdfDocument document = Measure("convert-and-write", () => source.ToPdf());
                Measure("retained-bytes", () => document.ToBytes());
                if (iteration < warmups) samples.Clear();
            }
            foreach ((string stage, var list) in samples) {
                double milliseconds = list.Select(sample => sample.Milliseconds).Order().ElementAt(list.Count / 2);
                long bytes = list.Select(sample => sample.Bytes).Order().ElementAt(list.Count / 2);
                Console.WriteLine($"{name,-22} {stage,-18} {milliseconds,9:0.0} ms {bytes / 1e6,9:0.0} MB");
            }
        }
        return 0;
    }

    /// <summary>Writes the renderer's own loss report once, outside every timed path.</summary>
    internal static int WriteOfficeImoDiagnostics(string[] args) {
        string corpus = Path.GetFullPath(Required(args, "--corpus"));
        string outputPath = Path.GetFullPath(Required(args, "--output"));
        var report = new SortedDictionary<string, object>(StringComparer.Ordinal);
        foreach (string file in Directory.GetFiles(corpus, "*.html")) {
            string name = Path.GetFileNameWithoutExtension(file);
            try {
                HtmlPdfRenderRequestResult result = HtmlConversionDocument.Parse(File.ReadAllText(file)).RenderToPdfResult(
                    HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, new HtmlToPdfOptions()));
                report[name] = result.Output.Report.Warnings
                    .Where(diagnostic => diagnostic.LossKind != OfficeConversionLossKind.None)
                    .GroupBy(diagnostic => diagnostic.Code + ": " + diagnostic.Message)
                    .Select(group => new { Diagnostic = group.Key, Count = group.Count() }).ToArray();
            } catch (Exception exception) {
                report[name] = new[] { new { Diagnostic = exception.GetType().Name + ": " + exception.Message, Count = 1 } };
            }
        }
        File.WriteAllText(outputPath, JsonSerializer.Serialize(report, Json));
        return 0;
    }

    private sealed class RendererFleet {
        private readonly Renderer[] _renderers;
        private readonly bool _shared;
        private RendererFleet(Renderer[] renderers, bool shared) { _renderers = renderers; _shared = shared; }
        internal Renderer this[int index] => _renderers[_shared ? 0 : index];

        internal static async Task<RendererFleet> CreateAsync(string engine, Renderer existing, int degree) {
            if (engine is "OfficeIMO" or "PdfPecker") return new RendererFleet(new[] { existing }, shared: true);
            var renderers = new Renderer[degree];
            for (int index = 0; index < degree; index++) renderers[index] = await Renderer.CreateAsync(engine);
            return new RendererFleet(renderers, shared: false);
        }
    }

    /// <summary>Holds per-instance engine state: PeachPDF allows sequential generator reuse, not concurrent use.</summary>
    private sealed class Renderer : IAsyncDisposable {
        private readonly string _engine;
        private readonly IPlaywright? _playwright;
        private readonly IBrowser? _browser;
        private readonly IPage? _page;
        private readonly PdfGenerator? _peach;

        private Renderer(string engine, IPlaywright? playwright, IBrowser? browser, IPage? page, PdfGenerator? peach) {
            _engine = engine;
            _playwright = playwright;
            _browser = browser;
            _page = page;
            _peach = peach;
        }

        internal static async Task<Renderer> CreateAsync(string engine) {
            if (engine == "Chromium") {
                IPlaywright playwright = await Playwright.CreateAsync();
                IBrowser browser = await playwright.Chromium.LaunchAsync();
                return new Renderer(engine, playwright, browser, await browser.NewPageAsync(), null);
            }
            return engine switch {
                "PeachPDF" => new Renderer(engine, null, null, null, new PdfGenerator()),
                "OfficeIMO" or "PdfPecker" => new Renderer(engine, null, null, null, null),
                _ => throw new ArgumentException("Unknown engine " + engine)
            };
        }

        internal async Task<byte[]> RenderAsync(string html) {
            switch (_engine) {
                case "OfficeIMO": {
                    using OfficeIMO.Pdf.PdfDocument document = HtmlConversionDocument.Parse(html).ToPdf();
                    return document.ToBytes();
                }
                case "PeachPDF": {
                    var config = new PdfGenerateConfig {
                        PageSize = PeachPDF.PageSize.A4,
                        PageOrientation = PeachPDF.PageOrientation.Portrait,
                        EnableTaggedPdf = true
                    };
                    var document = await _peach!.GeneratePdf(html, config);
                    using var stream = new MemoryStream();
                    document.Save(stream);
                    return stream.ToArray();
                }
                case "PdfPecker":
                    return global::PdfPecker.HtmlToPdf.Convert(html, new global::PdfPecker.PdfOptions {
                        PageSize = global::PdfPecker.PageSize.A4,
                        Accessibility = global::PdfPecker.PdfAccessibility.Tagged
                    }, CancellationToken.None);
                default:
                    await _page!.SetContentAsync(html, new PageSetContentOptions { WaitUntil = WaitUntilState.Load });
                    return await _page.PdfAsync(new PagePdfOptions {
                        Format = "A4", PrintBackground = true, PreferCSSPageSize = true, Tagged = true
                    });
            }
        }

        public async ValueTask DisposeAsync() {
            if (_browser != null) await _browser.CloseAsync();
            _playwright?.Dispose();
        }
    }

    private static class NativePeakRss {
        [StructLayout(LayoutKind.Sequential)]
        private struct RUsage {
            public long UserSeconds, UserMicroseconds, SystemSeconds, SystemMicroseconds, MaxRss;
            public long Ixrss, Idrss, Isrss, Minflt, Majflt, Nswap, Inblock, Oublock, Msgsnd, Msgrcv, Nsignals, Nvcsw, Nivcsw;
        }

        [DllImport("libc", SetLastError = true)]
        private static extern int getrusage(int who, out RUsage usage);

        /// <summary>ru_maxrss: bytes on macOS, kilobytes on Linux; Windows uses the managed peak working set.</summary>
        internal static long Read() {
            if (OperatingSystem.IsWindows()) return Process.GetCurrentProcess().PeakWorkingSet64;
            if (getrusage(0, out RUsage usage) != 0) return 0;
            return OperatingSystem.IsMacOS() ? usage.MaxRss : usage.MaxRss * 1024;
        }
    }

    internal sealed record WorkerRecord {
        public string Document { get; init; } = "";
        public string Engine { get; init; } = "";
        public string Mode { get; init; } = "";
        public string Status { get; init; } = "Passed";
        public string Detail { get; init; } = "";
        public double FirstMilliseconds { get; init; }
        public long FirstAllocatedBytes { get; init; }
        public double FirstCpuMilliseconds { get; init; }
        public double WarmMedianMilliseconds { get; init; }
        public double WarmCpuMedianMilliseconds { get; init; }
        public double WarmMinimumMilliseconds { get; init; }
        public int WarmIterations { get; init; }
        public int WarmupIterations { get; init; }
        public long WarmAllocatedBytes { get; init; }
        public int Gen2Collections { get; init; }
        public int ParallelDegree { get; init; }
        public double ParallelDocumentsPerSecond { get; init; }
        public long OutputBytes { get; init; }
        public long PeakRssBytes { get; init; }
        public double ProcessWallMilliseconds { get; init; }
    }

    private static string? Option(string[] args, string name) {
        int index = Array.FindIndex(args, value => string.Equals(value, name, StringComparison.OrdinalIgnoreCase));
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }

    private static string Required(string[] args, string name) =>
        Option(args, name) ?? throw new ArgumentException(name + " is required.");
}
