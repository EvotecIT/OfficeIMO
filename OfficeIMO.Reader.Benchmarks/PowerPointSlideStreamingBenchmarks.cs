using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader.Benchmarks.Comparison;
using OfficeIMO.Reader.PowerPoint;

namespace OfficeIMO.Reader.Benchmarks;

/// <summary>
/// Compares reading a large deck through the loaded presentation model with reading it slide by slide through
/// <see cref="PowerPointSlideReader"/>. Each pair of lanes produces the same validated output.
/// </summary>
[MemoryDiagnoser]
[ShortRunJob(RuntimeMoniker.Net80)]
[BenchmarkCategory("Performance")]
public class PowerPointSlideStreamingBenchmarks {
    private string _path = string.Empty;

    [Params(1_000)]
    public int Slides { get; set; }

    [GlobalSetup]
    public void Setup() {
        _path = PowerPointStreamingDeck.Create(Slides);
        PowerPointStreamingDeck.Validate(_path);
    }

    [GlobalCleanup]
    public void Cleanup() => PowerPointStreamingDeck.Delete(_path);

    [Benchmark(Baseline = true)]
    public long LoadedExtraction() => PowerPointStreamingDeck.Run("loaded-extraction", _path);

    [Benchmark]
    public long StreamedExtraction() => PowerPointStreamingDeck.Run("streamed-extraction", _path);

    [Benchmark]
    public long LoadedFirstSlide() => PowerPointStreamingDeck.Run("loaded-first-slide", _path);

    [Benchmark]
    public long StreamedFirstSlide() => PowerPointStreamingDeck.Run("streamed-first-slide", _path);

    [Benchmark]
    public long ReaderTextLoaded() => PowerPointStreamingDeck.Run("reader-text-loaded", _path);

    [Benchmark]
    public long ReaderTextStreamed() => PowerPointStreamingDeck.Run("reader-text-streamed", _path);
}

/// <summary>The generated deck and the measured operations shared by the benchmark and the peak-memory probe.</summary>
internal static class PowerPointStreamingDeck {
    internal static readonly string[] Lanes = {
        "loaded-extraction", "streamed-extraction", "loaded-first-slide", "streamed-first-slide",
        "reader-text-loaded", "reader-text-streamed"
    };

    private static readonly PowerPointExtractionExtensions.PowerPointExtractOptions PlainText = new() { IncludeMarkdown = false };
    private static readonly ReaderOptions TextProjection = new() { Projection = ReaderProjection.Text, ComputeHashes = false };
    private static readonly OfficeDocumentReader Reader = new OfficeDocumentReaderBuilder().AddPowerPointHandler().Build();
    private static readonly byte[] Png = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==");

    // Each slide has a title, a bulleted text box, a 6x4 table, a picture and speaker notes.
    internal static string Create(int slides) {
        string path = Path.Combine(Path.GetTempPath(), "officeimo-pptx-streaming-" + slides.ToString(CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N") + ".pptx");
        using (PowerPointDocument presentation = PowerPointDocument.Create(path)) {
            for (int number = 1; number <= slides; number++) {
                PowerPointSlide slide = presentation.AddSlide();
                slide.AddTitle("Quarterly review " + number.ToString(CultureInfo.InvariantCulture));
                PowerPointTextBox body = slide.AddTextBox("Highlights for region " + number.ToString(CultureInfo.InvariantCulture));
                for (int item = 1; item <= 5; item++) {
                    body.AddParagraph("Point " + item.ToString(CultureInfo.InvariantCulture) + ": revenue, margin and pipeline commentary for the period.");
                }
                PowerPointTable table = slide.AddTable(6, 4);
                for (int row = 0; row < 6; row++) {
                    for (int column = 0; column < 4; column++) {
                        table.GetRow(row).GetCell(column).Text = row == 0
                            ? "Column " + (column + 1).ToString(CultureInfo.InvariantCulture)
                            : (number * 100 + row * 10 + column).ToString(CultureInfo.InvariantCulture);
                    }
                }
                using (var image = new MemoryStream(Png, writable: false)) slide.AddImage(image, OfficeImageFormat.Png);
                slide.Notes.Text = "Speaker notes for slide " + number.ToString(CultureInfo.InvariantCulture) + ": cover the table and the outlook.";
            }
            presentation.Save();
        }
        return path;
    }

    internal static void Delete(string path) {
        if (File.Exists(path)) File.Delete(path);
    }

    // Both extraction lanes must return the same chunks, and both Reader lanes the same text.
    internal static void Validate(string path) {
        string[] loaded;
        using (PowerPointDocument presentation = PowerPointDocument.Load(path, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
            loaded = presentation.ExtractMarkdownChunks(PlainText).Select(chunk => chunk.Text).ToArray();
        }
        string[] streamed;
        using (PowerPointSlideReader reader = PowerPointSlideReader.Open(path)) {
            streamed = reader.ExtractMarkdownChunks(PlainText).Select(chunk => chunk.Text).ToArray();
        }
        if (loaded.Length == 0 || !loaded.SequenceEqual(streamed, StringComparer.Ordinal)) {
            throw new InvalidOperationException("Streamed extraction does not match the loaded presentation.");
        }
        string modelText;
        using (PowerPointDocument presentation = PowerPointDocument.Load(path, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
            modelText = presentation.ReadDocument(path, TextProjection).GetText();
        }
        OfficeDocumentReadResult streamedResult = Reader.ReadDocument(path, TextProjection);
        if (!streamedResult.CapabilitiesUsed.Contains("officeimo.reader.powerpoint.streamed-text")
            || !string.Equals(modelText, streamedResult.GetText(), StringComparison.Ordinal)) {
            throw new InvalidOperationException("The streamed text projection does not match the model projection.");
        }
    }

    internal static long Run(string lane, string path) {
        switch (lane) {
            case "loaded-extraction": {
                using PowerPointDocument presentation = PowerPointDocument.Load(path, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
                return presentation.ExtractMarkdownChunks(PlainText).Sum(chunk => (long)chunk.Text.Length);
            }
            case "streamed-extraction": {
                using PowerPointSlideReader reader = PowerPointSlideReader.Open(path);
                return reader.ExtractMarkdownChunks(PlainText).Sum(chunk => (long)chunk.Text.Length);
            }
            case "loaded-first-slide": {
                using PowerPointDocument presentation = PowerPointDocument.Load(path, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
                return presentation.ExtractMarkdownChunks(PlainText).First().Text.Length;
            }
            case "streamed-first-slide": {
                using PowerPointSlideReader reader = PowerPointSlideReader.Open(path);
                return reader.ExtractMarkdownChunks(PlainText).First().Text.Length;
            }
            case "reader-text-loaded": {
                using PowerPointDocument presentation = PowerPointDocument.Load(path, new PowerPointLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
                return presentation.ReadDocument(path, TextProjection).Chunks.Count;
            }
            case "reader-text-streamed":
                return Reader.ReadDocument(path, TextProjection).Chunks.Count;
            default:
                throw new ArgumentOutOfRangeException(nameof(lane), lane, "Unknown PowerPoint streaming lane.");
        }
    }
}

/// <summary>
/// Measures each lane's peak memory in its own process, which BenchmarkDotNet does not report:
/// <c>powerpoint-streaming-memory [slides]</c>. The child runs with a 2 MiB generation-0 budget so the sampled managed
/// heap follows live data rather than uncollected garbage; the process peak working set includes the runtime itself.
/// </summary>
internal static class PowerPointStreamingMemoryCommand {
    internal static int Run(string[] args) {
        if (args.Length == 3 && args[0] == "probe") {
            Console.WriteLine(Probe(args[1], args[2]));
            return 0;
        }
        int slides = args.Length > 0 ? int.Parse(args[0], CultureInfo.InvariantCulture) : 1_000;
        string path = PowerPointStreamingDeck.Create(slides);
        try {
            PowerPointStreamingDeck.Validate(path);
            Console.WriteLine("| Lane | Peak managed heap growth (MiB) | Process peak working set (MiB) |");
            Console.WriteLine("| --- | ---: | ---: |");
            foreach (string lane in PowerPointStreamingDeck.Lanes) {
                string[] values = RunChild(lane, path).Split(' ');
                Console.WriteLine("| " + lane + " | " + Mebibytes(values[0]) + " | " + Mebibytes(values[1]) + " |");
            }
        } finally {
            PowerPointStreamingDeck.Delete(path);
        }
        return 0;
    }

    private static string Probe(string lane, string path) {
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        using Process process = Process.GetCurrentProcess();
        ReaderMemoryPeak peak;
        using (var sampler = new ReaderMemorySampler(process)) {
            GC.KeepAlive(PowerPointStreamingDeck.Run(lane, path));
            peak = sampler.Stop();
        }
        process.Refresh();
        return Math.Max(0, peak.ManagedHeapBytes - heapBefore).ToString(CultureInfo.InvariantCulture) + " "
            + process.PeakWorkingSet64.ToString(CultureInfo.InvariantCulture);
    }

    private static string RunChild(string lane, string path) {
        string processPath = Environment.ProcessPath
            ?? throw new InvalidOperationException("Unable to resolve the benchmark process path.");
        var startInfo = new ProcessStartInfo {
            FileName = processPath,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        startInfo.Environment["DOTNET_GCgen0size"] = "0x200000";
        if (string.Equals(Path.GetFileNameWithoutExtension(processPath), "dotnet", StringComparison.OrdinalIgnoreCase)) {
            startInfo.ArgumentList.Add(Assembly.GetEntryAssembly()!.Location);
        }
        startInfo.ArgumentList.Add("powerpoint-streaming-memory");
        startInfo.ArgumentList.Add("probe");
        startInfo.ArgumentList.Add(lane);
        startInfo.ArgumentList.Add(path);
        using Process child = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Unable to start the memory probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) throw new InvalidOperationException($"Memory probe '{lane}' failed: {error}");
        return output.Trim();
    }

    private static string Mebibytes(string bytes) =>
        (long.Parse(bytes, CultureInfo.InvariantCulture) / 1024d / 1024d).ToString("0.0", CultureInfo.InvariantCulture);
}
