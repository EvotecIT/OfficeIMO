using System.Diagnostics;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class HtmlStaticBudgetWorker {
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    internal static async Task<int> RunAsync(string[] args) {
        string resultPath = Path.GetFullPath(ReadRequiredOption(args, "--result"));
        string mode = ReadRequiredOption(args, "--mode").ToLowerInvariant();
        int iterations = int.Parse(ReadRequiredOption(args, "--iterations"), System.Globalization.CultureInfo.InvariantCulture);
        if (mode is not ("cold" or "warm")) throw new ArgumentException("--mode must be cold or warm.");
        if (iterations < 1 || iterations > 10) throw new ArgumentOutOfRangeException(nameof(iterations));

        HtmlRenderingAdvancedHeldOutCorpus corpus = HtmlRenderingAdvancedHeldOutCorpus.Load();
        if (mode == "warm") {
            _ = RenderIteration(corpus, 0);
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        var results = new List<HtmlStaticBudgetIteration>(iterations);
        for (int iteration = 1; iteration <= iterations; iteration++) {
            results.Add(RenderIteration(corpus, iteration));
        }
        Directory.CreateDirectory(Path.GetDirectoryName(resultPath)!);
        await File.WriteAllTextAsync(resultPath, JsonSerializer.Serialize(results, JsonOptions)).ConfigureAwait(false);
        return 0;
    }

    internal static ProcessStartInfo CreateStartInfo(
        string mode,
        int iterations,
        string resultPath,
        string workingDirectory) {
        string entryAssemblyPath = Assembly.GetEntryAssembly()?.Location
            ?? throw new InvalidOperationException("The HTML static budget executable path is unavailable.");
        string processPath = Environment.ProcessPath ?? "dotnet";
        var startInfo = new ProcessStartInfo {
            FileName = processPath,
            WorkingDirectory = workingDirectory,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };
        if (string.Equals(Path.GetFileNameWithoutExtension(processPath), "dotnet", StringComparison.OrdinalIgnoreCase)) {
            startInfo.ArgumentList.Add(entryAssemblyPath);
        }
        startInfo.ArgumentList.Add("html-static-budget-worker");
        startInfo.ArgumentList.Add("--mode");
        startInfo.ArgumentList.Add(mode);
        startInfo.ArgumentList.Add("--iterations");
        startInfo.ArgumentList.Add(iterations.ToString(System.Globalization.CultureInfo.InvariantCulture));
        startInfo.ArgumentList.Add("--result");
        startInfo.ArgumentList.Add(resultPath);
        return startInfo;
    }

    internal static async Task<IReadOnlyList<HtmlStaticBudgetIteration>> ReadResultAsync(string path) {
        await using FileStream stream = File.OpenRead(path);
        return await JsonSerializer.DeserializeAsync<List<HtmlStaticBudgetIteration>>(stream, JsonOptions).ConfigureAwait(false)
            ?? throw new InvalidDataException("The HTML static budget worker result was empty.");
    }

    private static HtmlStaticBudgetIteration RenderIteration(HtmlRenderingAdvancedHeldOutCorpus corpus, int number) {
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: false);
        var stopwatch = Stopwatch.StartNew();
        long outputBytes = 0;
        int outputCount = 0;
        var hashes = new List<string>();
        foreach (HtmlRenderingAdvancedHeldOutCase scenario in corpus.Cases) {
            HtmlConversionDocument source = scenario.LoadDocument();
            RenderPdf(source, scenario, HtmlRenderIntentProfile.PrintPaged, CreatePrintOptions(), scenario.Manifest.ExpectedPrintPageCount, hashes, ref outputBytes, ref outputCount);
            RenderImages(source, scenario, HtmlRenderIntentProfile.PrintPaged, CreatePrintOptions(), hashes, ref outputBytes, ref outputCount);
            RenderImages(source, scenario, HtmlRenderIntentProfile.ScreenFullPage, CreateScreenOptions(), hashes, ref outputBytes, ref outputCount);
            RenderPdf(source, scenario, HtmlRenderIntentProfile.ScreenSnapshotPaged, CreateSnapshotOptions(), null, hashes, ref outputBytes, ref outputCount);
            RenderImages(source, scenario, HtmlRenderIntentProfile.ScreenSnapshotPaged, CreateSnapshotOptions(), hashes, ref outputBytes, ref outputCount);
        }
        stopwatch.Stop();
        long allocated = Math.Max(0L, GC.GetTotalAllocatedBytes(precise: false) - allocatedBefore);
        string fingerprint = Hash(Encoding.UTF8.GetBytes(string.Join("\n", hashes)));
        return new HtmlStaticBudgetIteration(
            number,
            stopwatch.Elapsed.TotalMilliseconds,
            allocated,
            outputBytes,
            corpus.Cases.Count,
            outputCount,
            fingerprint);
    }

    private static void RenderPdf(
        HtmlConversionDocument source,
        HtmlRenderingAdvancedHeldOutCase scenario,
        HtmlRenderIntentProfile profile,
        HtmlRenderOptions options,
        int? expectedPages,
        ICollection<string> hashes,
        ref long outputBytes,
        ref int outputCount) {
        HtmlPdfRenderRequestResult result = source.RenderToPdfResult(HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Pdf, options));
        ValidateDocument(result.RenderResult.Document, scenario, expectedPages);
        AddOutput(result.ToBytes(), hashes, ref outputBytes, ref outputCount);
    }

    private static void RenderImages(
        HtmlConversionDocument source,
        HtmlRenderingAdvancedHeldOutCase scenario,
        HtmlRenderIntentProfile profile,
        HtmlRenderOptions options,
        ICollection<string> hashes,
        ref long outputBytes,
        ref int outputCount) {
        foreach (HtmlRenderEncoder encoder in new[] { HtmlRenderEncoder.Png, HtmlRenderEncoder.Svg }) {
            HtmlRenderResult result = HtmlRenderEngine.Execute(source, HtmlRenderRequest.Create(profile, encoder, options));
            ValidateDocument(result.Document, scenario, profile == HtmlRenderIntentProfile.PrintPaged ? scenario.Manifest.ExpectedPrintPageCount : null);
            foreach (OfficeImageExportResult image in result.ExportImages()) {
                AddOutput(image.Bytes, hashes, ref outputBytes, ref outputCount);
            }
        }
    }

    private static void ValidateDocument(HtmlRenderDocument document, HtmlRenderingAdvancedHeldOutCase scenario, int? expectedPages) {
        if (expectedPages.HasValue && document.Pages.Count != expectedPages.Value) {
            throw new InvalidDataException($"{scenario.Id} rendered {document.Pages.Count} pages; expected {expectedPages.Value}.");
        }
        if (document.Diagnostics.Any(diagnostic => diagnostic.Severity == HtmlDiagnosticSeverity.Error)) {
            throw new InvalidDataException(scenario.Id + " produced an error diagnostic.");
        }
        string text = string.Join(" ", document.Text.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
        foreach (string marker in scenario.Manifest.TextMarkers) {
            string normalized = string.Join(" ", marker.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));
            if (!text.Contains(normalized, StringComparison.Ordinal)) {
                throw new InvalidDataException(scenario.Id + " is missing text marker: " + marker + ".");
            }
        }
    }

    private static void AddOutput(byte[] bytes, ICollection<string> hashes, ref long outputBytes, ref int outputCount) {
        if (bytes.Length == 0) throw new InvalidDataException("A static rendering output was empty.");
        outputBytes += bytes.LongLength;
        outputCount++;
        hashes.Add(Hash(bytes));
    }

    private static HtmlRenderOptions CreateScreenOptions() => new() {
        Mode = HtmlRenderMode.Continuous, ViewportWidth = 768D, ViewportHeight = 900D,
        Margins = HtmlRenderMargins.All(0D), Scale = 1D, BackgroundColor = OfficeColor.White,
        UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(), FidelityPolicy = HtmlRenderFidelityPolicy.AllowDiagnosedLoss
    };

    private static HtmlRenderOptions CreatePrintOptions() => new() {
        Mode = HtmlRenderMode.Paged, ViewportWidth = 768D, ViewportHeight = 900D,
        PageSize = new OfficePageSize(8.27D, 11.69D), Margins = HtmlRenderMargins.All(40D),
        Scale = 1D, BackgroundColor = OfficeColor.White, UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
        FidelityPolicy = HtmlRenderFidelityPolicy.AllowDiagnosedLoss
    };

    private static HtmlRenderOptions CreateSnapshotOptions() => new() {
        Mode = HtmlRenderMode.Continuous, ViewportWidth = 768D, ViewportHeight = 900D,
        PageSize = new OfficePageSize(8D, 9.375D), Margins = HtmlRenderMargins.All(0D), HonorCssPageRules = false,
        Scale = 1D, BackgroundColor = OfficeColor.White, UrlPolicy = HtmlUrlPolicy.CreateWebOnlyProfile(),
        FidelityPolicy = HtmlRenderFidelityPolicy.AllowDiagnosedLoss
    };

    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static string ReadRequiredOption(string[] args, string option) {
        for (int index = 1; index < args.Length - 1; index++) {
            if (string.Equals(args[index], option, StringComparison.OrdinalIgnoreCase)) return args[index + 1];
        }
        throw new ArgumentException(option + " is required.");
    }
}
