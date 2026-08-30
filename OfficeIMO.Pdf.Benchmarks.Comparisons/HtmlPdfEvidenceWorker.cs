using System.Diagnostics;
using System.Reflection;
using System.Text.Json;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal sealed record HtmlPdfWorkerResult(
    double DurationMilliseconds,
    long ManagedAllocatedBytes);

internal static class HtmlPdfEvidenceWorker {
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    internal static async Task<int> RunAsync(string[] args) {
        string engineValue = ReadRequiredOption(args, "--engine");
        if (!Enum.TryParse(engineValue, ignoreCase: true, out HtmlPdfComparisonEngine engine) || !Enum.IsDefined(engine)) {
            throw new ArgumentException("--engine must be OfficeIMO, PeachPDF, ITextPdfHtml, or Chromium.");
        }

        string htmlPath = Path.GetFullPath(ReadRequiredOption(args, "--html"));
        string outputPath = Path.GetFullPath(ReadRequiredOption(args, "--output"));
        string resultPath = Path.GetFullPath(ReadRequiredOption(args, "--result"));
        string html = await File.ReadAllTextAsync(htmlPath).ConfigureAwait(false);

        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: false);
        var stopwatch = Stopwatch.StartNew();
        byte[] bytes;
        if (engine == HtmlPdfComparisonEngine.Chromium) {
            await using HtmlTinkerX.HtmlBrowserSession session =
                await HtmlPdfComparisonRenderers.OpenChromiumSessionAsync().ConfigureAwait(false);
            bytes = await HtmlPdfComparisonRenderers.RenderChromiumAsync(session, html).ConfigureAwait(false);
        } else {
            bytes = HtmlPdfComparisonRenderers.RenderManaged(engine, html);
        }
        stopwatch.Stop();
        long allocatedAfter = GC.GetTotalAllocatedBytes(precise: false);

        await File.WriteAllBytesAsync(outputPath, bytes).ConfigureAwait(false);
        var result = new HtmlPdfWorkerResult(
            stopwatch.Elapsed.TotalMilliseconds,
            Math.Max(0L, allocatedAfter - allocatedBefore));
        await File.WriteAllTextAsync(resultPath, JsonSerializer.Serialize(result, JsonOptions)).ConfigureAwait(false);
        return 0;
    }

    internal static ProcessStartInfo CreateStartInfo(
        HtmlPdfComparisonEngine engine,
        string htmlPath,
        string outputPath,
        string resultPath,
        string workingDirectory) {
        string entryAssemblyPath = Assembly.GetEntryAssembly()?.Location
            ?? throw new InvalidOperationException("The HTML-to-PDF evidence executable path is unavailable.");
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
        startInfo.ArgumentList.Add("html-evidence-worker");
        startInfo.ArgumentList.Add("--engine");
        startInfo.ArgumentList.Add(engine.ToString());
        startInfo.ArgumentList.Add("--html");
        startInfo.ArgumentList.Add(htmlPath);
        startInfo.ArgumentList.Add("--output");
        startInfo.ArgumentList.Add(outputPath);
        startInfo.ArgumentList.Add("--result");
        startInfo.ArgumentList.Add(resultPath);
        return startInfo;
    }

    internal static async Task<HtmlPdfWorkerResult> ReadResultAsync(string resultPath) {
        await using FileStream stream = File.OpenRead(resultPath);
        return await JsonSerializer.DeserializeAsync<HtmlPdfWorkerResult>(stream, JsonOptions).ConfigureAwait(false)
            ?? throw new InvalidDataException("The HTML-to-PDF evidence worker result was empty.");
    }

    private static string ReadRequiredOption(string[] args, string option) {
        for (int index = 1; index < args.Length - 1; index++) {
            if (string.Equals(args[index], option, StringComparison.OrdinalIgnoreCase)) return args[index + 1];
        }
        throw new ArgumentException(option + " is required.");
    }
}
