using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace OfficeIMO.OneNote.Benchmarks;

internal static class OneNoteEvidenceRunner {
    private static readonly string[] Operations = { "CreateWrite", "Read", "ReadEditWrite" };
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        if (args.Length < 2 || args.Length > 3) {
            Console.Error.WriteLine("Usage: --probe <CreateWrite|Read|ReadEditWrite> <Small|Normal|Large> [source.one]");
            return 2;
        }
        try {
            Console.WriteLine(JsonSerializer.Serialize(
                Measure(args[0], args[1], args.Length == 3 ? args[2] : null),
                JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    internal static int RunEvidence(string[] args) {
        string? operationFilter = GetOption(args, "--operation");
        string? scaleFilter = GetOption(args, "--scale");
        string? jsonPath = GetOption(args, "--json");
        int repeat = GetPositiveIntOption(args, "--repeat", 1);
        string[] operations = SelectOperations(operationFilter);
        OneNoteBenchmarkScale[] scales = string.IsNullOrWhiteSpace(scaleFilter)
            ? OneNoteBenchmarkCorpus.Scales
            : new[] { OneNoteBenchmarkCorpus.Get(scaleFilter!) };
        var measurements = new List<OneNoteEvidenceMeasurement>();

        foreach (OneNoteBenchmarkScale scale in scales) {
            string sourcePath = CreateTemporarySource(scale);
            try {
                foreach (string operation in operations) {
                    for (var iteration = 1; iteration <= repeat; iteration++) {
                        OneNoteEvidenceMeasurement measurement = RunChildProbe(
                            operation,
                            scale.Name,
                            string.Equals(operation, "CreateWrite", StringComparison.OrdinalIgnoreCase) ? null : sourcePath)
                            with { Iteration = iteration };
                        measurements.Add(measurement);
                        Console.WriteLine(
                            $"{operation,-13} {scale.Name,-6} #{iteration,-2} " +
                            $"{measurement.ElapsedMilliseconds,9:F2} ms " +
                            $"{measurement.AllocatedBytes / 1048576D,9:F2} MiB alloc " +
                            $"{measurement.PeakWorkingSetBytes / 1048576D,9:F2} MiB peak " +
                            $"{measurement.OutputBytes / 1024D,9:F2} KiB output");
                    }
                }
            } finally {
                File.Delete(sourcePath);
            }
        }

        var report = new OneNoteEvidenceReport(
            DateTimeOffset.UtcNow,
            ResolveCommit(),
            ResolveSourceTreeDirty(),
            RuntimeInformation.FrameworkDescription,
            RuntimeInformation.OSDescription,
            RuntimeInformation.ProcessArchitecture.ToString(),
            Environment.ProcessorCount,
            measurements);
        if (!string.IsNullOrWhiteSpace(jsonPath)) {
            string fullPath = Path.GetFullPath(jsonPath!);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
            File.WriteAllText(fullPath, JsonSerializer.Serialize(report, JsonOptions));
            Console.WriteLine("Wrote " + fullPath);
        }
        return 0;
    }

    private static OneNoteEvidenceMeasurement Measure(string operation, string scaleName, string? sourcePath) {
        operation = SelectOperations(operation).Single();
        OneNoteBenchmarkScale scale = OneNoteBenchmarkCorpus.Get(scaleName);
        if (!string.Equals(operation, "CreateWrite", StringComparison.OrdinalIgnoreCase)
            && (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))) {
            throw new FileNotFoundException("An existing OneNote section is required for this operation.", sourcePath);
        }

        long inputBytes = sourcePath == null ? 0 : new FileInfo(sourcePath).Length;
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        OneNoteOperationResult result = Execute(operation, scale, sourcePath);
        stopwatch.Stop();
        long allocatedBytes = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        using Process process = Process.GetCurrentProcess();
        process.Refresh();

        OneNoteCorpusObservation observation = OneNoteBenchmarkCorpus.Observe(result.Section);
        Validate(operation, scale, result, observation);
        return new OneNoteEvidenceMeasurement(
            operation,
            scale.Name,
            1,
            scale.PageCount,
            inputBytes,
            result.OutputBytes,
            stopwatch.Elapsed.TotalMilliseconds,
            allocatedBytes,
            process.PeakWorkingSet64,
            observation.PageCount,
            observation.ParagraphCount,
            observation.StructuralFingerprint);
    }

    private static OneNoteOperationResult Execute(
        string operation,
        OneNoteBenchmarkScale scale,
        string? sourcePath) {
        if (string.Equals(operation, "CreateWrite", StringComparison.OrdinalIgnoreCase)) {
            OneNoteSection section = OneNoteBenchmarkCorpus.CreateSection(scale.PageCount);
            byte[] bytes = OneNoteSectionWriter.Write(section);
            return new OneNoteOperationResult(bytes, section);
        }

        using var input = new FileStream(sourcePath!, FileMode.Open, FileAccess.Read, FileShare.Read);
        OneNoteSection loaded = OneNoteSectionReader.Read(input);
        if (string.Equals(operation, "Read", StringComparison.OrdinalIgnoreCase)) {
            return new OneNoteOperationResult(null, loaded);
        }

        loaded.Pages.Add(OneNoteBenchmarkCorpus.CreateEditPage(loaded.Pages.Count));
        byte[] edited = OneNoteSectionWriter.Write(loaded);
        return new OneNoteOperationResult(edited, loaded);
    }

    private static void Validate(
        string operation,
        OneNoteBenchmarkScale scale,
        OneNoteOperationResult result,
        OneNoteCorpusObservation observation) {
        int expectedPages = scale.PageCount
            + (string.Equals(operation, "ReadEditWrite", StringComparison.OrdinalIgnoreCase) ? 1 : 0);
        int expectedParagraphs = expectedPages * 8;
        if (observation.PageCount != expectedPages || observation.ParagraphCount != expectedParagraphs) {
            throw new InvalidOperationException(
                $"{operation}/{scale.Name} observed {observation.PageCount} pages and " +
                $"{observation.ParagraphCount} paragraphs; expected {expectedPages} and {expectedParagraphs}.");
        }
        bool edited = string.Equals(operation, "ReadEditWrite", StringComparison.OrdinalIgnoreCase);
        if (edited && !observation.ContainsEditMarker) {
            throw new InvalidOperationException(operation + "/" + scale.Name + " lost the edit marker.");
        }
        OneNoteSection expectedSection = OneNoteBenchmarkCorpus.CreateSection(scale.PageCount);
        if (edited) expectedSection.Pages.Add(OneNoteBenchmarkCorpus.CreateEditPage(expectedSection.Pages.Count));
        OneNoteCorpusObservation expected = OneNoteBenchmarkCorpus.Observe(expectedSection);
        if (observation != expected) {
            throw new InvalidOperationException(operation + "/" + scale.Name + " failed exact ordered semantic validation.");
        }
        if (result.PackageBytes == null) return;

        using var stream = new MemoryStream(result.PackageBytes, writable: false);
        OneNoteCorpusObservation reopened = OneNoteBenchmarkCorpus.Observe(OneNoteSectionReader.Read(stream));
        if (reopened != expected) {
            throw new InvalidOperationException(operation + "/" + scale.Name + " failed exact semantic round-trip validation.");
        }
    }

    private static OneNoteEvidenceMeasurement RunChildProbe(
        string operation,
        string scale,
        string? sourcePath) {
        string processPath = Environment.ProcessPath
            ?? throw new InvalidOperationException("Unable to resolve benchmark process path.");
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
        foreach (string argument in new[] { "--probe", operation, scale }) startInfo.ArgumentList.Add(argument);
        if (sourcePath != null) startInfo.ArgumentList.Add(sourcePath);

        using Process child = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Unable to start OneNote benchmark probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) {
            throw new InvalidOperationException($"Probe {operation}/{scale} failed: {error}");
        }
        return JsonSerializer.Deserialize<OneNoteEvidenceMeasurement>(output, JsonOptions)
            ?? throw new InvalidOperationException($"Probe {operation}/{scale} returned no measurement.");
    }

    private static string CreateTemporarySource(OneNoteBenchmarkScale scale) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-OneNote-{scale.Name}-{Guid.NewGuid():N}.one");
        File.WriteAllBytes(path, OneNoteSectionWriter.Write(OneNoteBenchmarkCorpus.CreateSection(scale.PageCount)));
        return path;
    }

    private static string[] SelectOperations(string? filter) {
        if (string.IsNullOrWhiteSpace(filter)) return Operations;
        string? selected = Operations.FirstOrDefault(value => string.Equals(value, filter, StringComparison.OrdinalIgnoreCase));
        return selected == null
            ? throw new ArgumentException("Unknown OneNote benchmark operation: " + filter)
            : new[] { selected };
    }

    private static string? GetOption(string[] args, string name) {
        int index = Array.FindIndex(args, argument => string.Equals(argument, name, StringComparison.OrdinalIgnoreCase));
        if (index < 0) return null;
        if (index + 1 >= args.Length || args[index + 1].StartsWith("--", StringComparison.Ordinal)) {
            throw new ArgumentException(name + " requires a value.");
        }
        return args[index + 1];
    }

    private static int GetPositiveIntOption(string[] args, string name, int defaultValue) {
        string? value = GetOption(args, name);
        if (value == null) return defaultValue;
        return int.TryParse(value, out int parsed) && parsed > 0
            ? parsed
            : throw new ArgumentException(name + " must be a positive integer.");
    }

    private static string ResolveCommit() {
        string? value = Environment.GetEnvironmentVariable("GITHUB_SHA");
        if (!string.IsNullOrWhiteSpace(value)) return value;
        try {
            var startInfo = CreateGitStartInfo("rev-parse", "HEAD");
            using Process process = Process.Start(startInfo)!;
            string output = process.StandardOutput.ReadToEnd().Trim();
            process.WaitForExit();
            return process.ExitCode == 0 ? output : "unknown";
        } catch {
            return "unknown";
        }
    }

    private static bool ResolveSourceTreeDirty() {
        try {
            using Process tracked = Process.Start(CreateGitStartInfo("diff", "--quiet", "HEAD", "--"))!;
            tracked.WaitForExit();
            if (tracked.ExitCode != 0) return true;
            using Process untracked = Process.Start(CreateGitStartInfo("ls-files", "--others", "--exclude-standard"))!;
            string output = untracked.StandardOutput.ReadToEnd();
            untracked.WaitForExit();
            return untracked.ExitCode != 0 || !string.IsNullOrWhiteSpace(output);
        } catch {
            return true;
        }
    }

    private static ProcessStartInfo CreateGitStartInfo(params string[] arguments) {
        var startInfo = new ProcessStartInfo {
            FileName = "git",
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        foreach (string argument in arguments) startInfo.ArgumentList.Add(argument);
        return startInfo;
    }
}

internal sealed record OneNoteOperationResult(byte[]? PackageBytes, OneNoteSection Section) {
    internal long OutputBytes => PackageBytes?.LongLength ?? 0;
}

internal sealed record OneNoteEvidenceMeasurement(
    string Operation,
    string Scale,
    int Iteration,
    int ExpectedPageCount,
    long InputBytes,
    long OutputBytes,
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long PeakWorkingSetBytes,
    int ObservedPageCount,
    int ObservedParagraphCount,
    string StructuralFingerprint);

internal sealed record OneNoteEvidenceReport(
    DateTimeOffset MeasuredAtUtc,
    string SourceCommit,
    bool SourceTreeDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    IReadOnlyList<OneNoteEvidenceMeasurement> Measurements);
