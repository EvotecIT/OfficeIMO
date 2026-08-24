using OfficeIMO.Confluence;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json;

namespace OfficeIMO.Confluence.Benchmarks;

internal static class ConfluenceEvidenceRunner {
    private static readonly int[] PageSizes = [16 * 1024, 1024 * 1024];
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        if (args.Length != 1 || !int.TryParse(args[0], out int pageCharacters) || pageCharacters <= 0) {
            Console.Error.WriteLine("Usage: evidence-probe <page-characters>");
            return 2;
        }
        try {
            Console.WriteLine(JsonSerializer.Serialize(Measure(pageCharacters), JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    internal static int RunEvidence(string[] args) {
        try {
            int repeat = GetPositiveIntOption(args, "--repeat", 3);
            string? jsonPath = GetOption(args, "--json");
            var measurements = new List<ConfluenceEvidenceMeasurement>(PageSizes.Length * repeat);
            foreach (int pageCharacters in PageSizes) {
                for (int iteration = 1; iteration <= repeat; iteration++) {
                    ConfluenceEvidenceMeasurement measurement = RunChildProbe(pageCharacters) with { Iteration = iteration };
                    measurements.Add(measurement);
                    Console.WriteLine(
                        $"{pageCharacters,8:N0} chars #{iteration}: {measurement.ElapsedMilliseconds,7:F3} ms, " +
                        $"{measurement.AllocatedBytes / 1048576D,6:F2} MiB alloc, " +
                        $"{measurement.RetainedManagedHeapGrowthBytes / 1048576D,6:F2} MiB retained, " +
                        $"{measurement.PeakManagedHeapGrowthBytes / 1048576D,6:F2} MiB managed peak, " +
                        $"{measurement.PeakWorkingSetGrowthBytes / 1048576D,6:F2} MiB process growth");
                }
            }

            var report = new ConfluenceEvidenceReport(
                DateTimeOffset.UtcNow,
                ResolveCommit(),
                ResolveSourceTreeDirty(),
                RuntimeInformation.FrameworkDescription,
                RuntimeInformation.OSDescription,
                RuntimeInformation.ProcessArchitecture.ToString(),
                Environment.ProcessorCount,
                repeat,
                measurements);
            if (!string.IsNullOrWhiteSpace(jsonPath)) {
                string fullPath = Path.GetFullPath(jsonPath!);
                string? directory = Path.GetDirectoryName(fullPath);
                if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
                File.WriteAllText(fullPath, JsonSerializer.Serialize(report, JsonOptions));
                Console.WriteLine("Wrote " + fullPath);
            }
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    private static ConfluenceEvidenceMeasurement Measure(int pageCharacters) {
        ConfluenceManagedSectionCorpus corpus = ConfluenceManagedSectionCorpusFactory.Create(pageCharacters);
        for (int index = 0; index < (pageCharacters < 100_000 ? 32 : 4); index++) {
            GC.KeepAlive(Apply(corpus));
        }

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        ConfluenceManagedSectionResult timedResult = Apply(corpus);
        stopwatch.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        Validate(corpus, timedResult);
        GC.KeepAlive(timedResult);

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        long workingSetBefore = process.WorkingSet64;
        ConfluenceManagedSectionResult result;
        ConfluenceMemoryPeak peak;
        using (var sampler = new ConfluenceMemorySampler(process)) {
            result = Apply(corpus);
            peak = sampler.Stop();
        }
        Validate(corpus, result);
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long retained = Math.Max(0, GC.GetTotalMemory(forceFullCollection: false) - heapBefore);
        process.Refresh();
        long absoluteProcessPeak = process.PeakWorkingSet64;
        GC.KeepAlive(result);

        return new ConfluenceEvidenceMeasurement(
            pageCharacters,
            1,
            Encoding.UTF8.GetByteCount(corpus.ExistingBody),
            Encoding.UTF8.GetByteCount(result.UpdatedBody),
            result.UpdatedBody.Length,
            stopwatch.Elapsed.TotalMilliseconds,
            allocated,
            retained,
            Math.Max(0, peak.ManagedHeapBytes - heapBefore),
            Math.Max(0, peak.WorkingSetBytes - workingSetBefore),
            absoluteProcessPeak,
            result.OriginalSha256,
            result.UpdatedSha256);
    }

    private static ConfluenceManagedSectionResult Apply(ConfluenceManagedSectionCorpus corpus) =>
        ConfluenceManagedSection.Apply(corpus.ExistingBody, ConfluenceManagedSectionCorpusFactory.SectionId, corpus.Replacement);

    private static void Validate(ConfluenceManagedSectionCorpus corpus, ConfluenceManagedSectionResult result) {
        if (!result.Changed || result.WasCreated || !result.UpdatedBody.Contains(corpus.Replacement, StringComparison.Ordinal)) {
            throw new InvalidOperationException("Managed-section evidence validation failed.");
        }
        if (string.Equals(result.OriginalSha256, result.UpdatedSha256, StringComparison.Ordinal)) {
            throw new InvalidOperationException("Managed-section evidence hashes did not change.");
        }
    }

    private static ConfluenceEvidenceMeasurement RunChildProbe(int pageCharacters) {
        string processPath = Environment.ProcessPath ?? throw new InvalidOperationException("No process path.");
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
        startInfo.ArgumentList.Add("evidence-probe");
        startInfo.ArgumentList.Add(pageCharacters.ToString(System.Globalization.CultureInfo.InvariantCulture));
        using Process child = Process.Start(startInfo) ?? throw new InvalidOperationException("Unable to start evidence probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) throw new InvalidOperationException("Evidence probe failed: " + error);
        return JsonSerializer.Deserialize<ConfluenceEvidenceMeasurement>(output, JsonOptions)
            ?? throw new InvalidOperationException("Evidence probe returned no measurement.");
    }

    private static string? GetOption(string[] args, string name) {
        int index = Array.FindIndex(args, value => string.Equals(value, name, StringComparison.OrdinalIgnoreCase));
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }

    private static int GetPositiveIntOption(string[] args, string name, int fallback) =>
        int.TryParse(GetOption(args, name), out int value) && value > 0 ? value : fallback;

    private static string ResolveCommit() => RunGit("rev-parse", "HEAD") ?? "unknown";
    private static bool ResolveSourceTreeDirty() => !string.IsNullOrWhiteSpace(RunGit("status", "--porcelain", "--untracked-files=no"));

    private static string? RunGit(params string[] arguments) {
        var startInfo = new ProcessStartInfo {
            FileName = "git",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        foreach (string argument in arguments) startInfo.ArgumentList.Add(argument);
        using Process? process = Process.Start(startInfo);
        if (process == null) return null;
        string output = process.StandardOutput.ReadToEnd();
        process.WaitForExit();
        return process.ExitCode == 0 ? output.Trim() : null;
    }
}
