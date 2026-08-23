using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;

namespace OfficeIMO.OpenDocument.Benchmarks;

internal static class OpenDocumentEvidenceRunner {
    private const string EditMarker = "OfficeIMO benchmark edit marker";
    private static readonly string[] Formats = { "ODT", "ODS", "ODP" };
    private static readonly string[] Operations = { "CreateSave", "OpenRead", "OpenEditSave" };
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        if (args.Length < 3 || args.Length > 4) {
            Console.Error.WriteLine("Usage: --probe <ODT|ODS|ODP> <CreateSave|OpenRead|OpenEditSave> <Small|Normal|Large> [source]");
            return 2;
        }
        try {
            string? sourcePath = args.Length == 4 ? args[3] : null;
            OpenDocumentEvidenceMeasurement measurement = Measure(args[0], args[1], args[2], sourcePath);
            Console.WriteLine(JsonSerializer.Serialize(measurement, JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    internal static int RunEvidence(string[] args) {
        string? formatFilter = GetOption(args, "--format");
        string? operationFilter = GetOption(args, "--operation");
        string? scaleFilter = GetOption(args, "--scale");
        string? jsonPath = GetOption(args, "--json");
        int repeat = GetPositiveIntOption(args, "--repeat", 1);
        string[] formats = Select(Formats, formatFilter, "format");
        string[] operations = Select(Operations, operationFilter, "operation");
        OpenDocumentBenchmarkScale[] scales = string.IsNullOrWhiteSpace(scaleFilter)
            ? OpenDocumentBenchmarkCorpus.Scales.ToArray()
            : new[] { OpenDocumentBenchmarkCorpus.Get(scaleFilter!) };
        var measurements = new List<OpenDocumentEvidenceMeasurement>();

        foreach (OpenDocumentBenchmarkScale scale in scales) {
            foreach (string format in formats) {
                string sourcePath = CreateTemporarySource(format, scale);
                try {
                    foreach (string operation in operations) {
                        for (var iteration = 1; iteration <= repeat; iteration++) {
                            OpenDocumentEvidenceMeasurement measurement = RunChildProbe(
                                format,
                                operation,
                                scale.Name,
                                string.Equals(operation, "CreateSave", StringComparison.OrdinalIgnoreCase) ? null : sourcePath)
                                with { Iteration = iteration };
                            measurements.Add(measurement);
                            Console.WriteLine(
                                $"{format,-3} {operation,-12} {scale.Name,-6} #{iteration,-2} " +
                                $"{measurement.ElapsedMilliseconds,10:F1} ms " +
                                $"{measurement.AllocatedBytes / 1048576D,10:F1} MiB alloc " +
                                $"{measurement.PeakWorkingSetBytes / 1048576D,10:F1} MiB peak " +
                                $"{measurement.OutputBytes / 1048576D,10:F2} MiB output");
                        }
                    }
                } finally {
                    File.Delete(sourcePath);
                }
            }
        }

        var report = new OpenDocumentEvidenceReport(
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
            Directory.CreateDirectory(Path.GetDirectoryName(fullPath)!);
            File.WriteAllText(fullPath, JsonSerializer.Serialize(report, JsonOptions));
            Console.WriteLine("Wrote " + fullPath);
        }
        return 0;
    }

    private static OpenDocumentEvidenceMeasurement Measure(
        string format,
        string operation,
        string scaleName,
        string? sourcePath) {
        format = Select(Formats, format, "format").Single();
        operation = Select(Operations, operation, "operation").Single();
        OpenDocumentBenchmarkScale scale = OpenDocumentBenchmarkCorpus.Get(scaleName);
        if (!string.Equals(operation, "CreateSave", StringComparison.OrdinalIgnoreCase)
            && (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))) {
            throw new FileNotFoundException("An existing source package is required for open workflows.", sourcePath);
        }

        long inputBytes = sourcePath == null ? 0 : new FileInfo(sourcePath).Length;
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        OpenDocumentOperationResult result = Execute(format, operation, scale, sourcePath);
        stopwatch.Stop();
        long allocatedBytes = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        long peakWorkingSetBytes = process.PeakWorkingSet64;

        OpenDocumentContentObservation observation = Validate(format, operation, scale, sourcePath, result);
        return new OpenDocumentEvidenceMeasurement(
            format,
            operation,
            scale.Name,
            1,
            ExpectedPrimaryCount(format, scale),
            inputBytes,
            result.OutputBytes,
            stopwatch.Elapsed.TotalMilliseconds,
            allocatedBytes,
            peakWorkingSetBytes,
            observation.ObservedCount,
            observation.Checksum);
    }

    private static OpenDocumentOperationResult Execute(
        string format,
        string operation,
        OpenDocumentBenchmarkScale scale,
        string? sourcePath) {
        if (string.Equals(operation, "CreateSave", StringComparison.OrdinalIgnoreCase)) {
            byte[] package = OpenDocumentBenchmarkCorpus.CreatePackage(format, scale);
            return new OpenDocumentOperationResult(package, ExpectedPrimaryCount(format, scale), 0);
        }
        if (string.Equals(operation, "OpenRead", StringComparison.OrdinalIgnoreCase)) {
            return Read(format, sourcePath!);
        }

        return Edit(format, sourcePath!);
    }

    private static OpenDocumentOperationResult Read(string format, string sourcePath) {
        switch (format) {
            case "ODT": {
                OdtDocument document = OdtDocument.Load(sourcePath);
                IReadOnlyList<OdtParagraph> paragraphs = document.Paragraphs;
                return new OpenDocumentOperationResult(null, paragraphs.Count, paragraphs.Sum(paragraph => (long)paragraph.Text.Length));
            }
            case "ODS": {
                OdsDocument document = OdsDocument.Load(sourcePath);
                OdsSheet sheet = document.Sheets.Single();
                long checksum = 0;
                long count = 0;
                foreach (OdsRowRun row in sheet.RowRuns) {
                    count += row.RepeatCount;
                    foreach (OdsCellRun cell in row.CellRuns) {
                        checksum += checked(row.RepeatCount * cell.RepeatCount * cell.Value.ToString().Length);
                    }
                }
                return new OpenDocumentOperationResult(null, count, checksum);
            }
            case "ODP": {
                OdpPresentation presentation = OdpPresentation.Load(sourcePath);
                long checksum = presentation.Slides.Sum(slide =>
                    slide.Shapes.OfType<OdpTextBox>().SelectMany(box => box.Paragraphs).Sum(paragraph => (long)paragraph.Text.Length));
                return new OpenDocumentOperationResult(null, presentation.Slides.Count, checksum);
            }
            default:
                throw new ArgumentException("Unknown format: " + format, nameof(format));
        }
    }

    private static OpenDocumentOperationResult Edit(string format, string sourcePath) {
        byte[] package;
        int count;
        switch (format) {
            case "ODT": {
                OdtDocument document = OdtDocument.Load(sourcePath);
                document.AddParagraph(EditMarker);
                count = document.Paragraphs.Count;
                package = document.ToBytes();
                break;
            }
            case "ODS": {
                OdsDocument document = OdsDocument.Load(sourcePath);
                OdsSheet sheet = document.Sheets.Single();
                long row = sheet.UsedRange?.LastRow + 1 ?? 0;
                sheet.Cell(row, 0).SetString(EditMarker);
                count = checked((int)(row + 1));
                package = document.ToBytes();
                break;
            }
            case "ODP": {
                OdpPresentation presentation = OdpPresentation.Load(sourcePath);
                presentation.AddSlide(EditMarker)
                    .AddTextBox(OdfRect.FromCentimeters(1, 1, 20, 2), EditMarker);
                count = presentation.Slides.Count;
                package = presentation.ToBytes();
                break;
            }
            default:
                throw new ArgumentException("Unknown format: " + format, nameof(format));
        }
        return new OpenDocumentOperationResult(package, count, 0);
    }

    private static OpenDocumentContentObservation Validate(
        string format,
        string operation,
        OpenDocumentBenchmarkScale scale,
        string? sourcePath,
        OpenDocumentOperationResult result) {
        int expected = ExpectedPrimaryCount(format, scale)
            + (string.Equals(operation, "OpenEditSave", StringComparison.OrdinalIgnoreCase) ? 1 : 0);
        byte[] bytes = result.PackageBytes ?? File.ReadAllBytes(sourcePath!);
        OpenDocumentContentObservation observation = InspectSerializedContent(format, bytes);
        if (observation.ObservedCount != expected) {
            throw new InvalidOperationException($"{format}/{operation} serialized {observation.ObservedCount} records; expected exactly {expected}.");
        }
        if (observation.Checksum <= 0) {
            throw new InvalidOperationException(format + "/" + operation + " produced no serialized content checksum.");
        }
        if (string.Equals(operation, "OpenEditSave", StringComparison.OrdinalIgnoreCase) && !observation.ContainsEditMarker) {
            throw new InvalidOperationException(format + "/" + operation + " did not preserve the edit marker.");
        }
        if (string.Equals(operation, "OpenRead", StringComparison.OrdinalIgnoreCase)
            && (result.ObservedCount != observation.ObservedCount || result.Checksum != observation.Checksum)) {
            throw new InvalidOperationException($"{format}/{operation} measured traversal did not match serialized content inspection.");
        }
        if (result.PackageBytes != null && result.OutputBytes <= 0) {
            throw new InvalidOperationException(format + "/" + operation + " produced an empty package.");
        }
        return observation;
    }

    private static OpenDocumentContentObservation InspectSerializedContent(string format, byte[] bytes) {
        using var stream = new MemoryStream(bytes, writable: false);
        switch (format) {
            case "ODT": {
                OdtDocument document = OdtDocument.Load(stream);
                EnsureStructurallyValid(document, format);
                IReadOnlyList<OdtParagraph> paragraphs = document.Paragraphs;
                return new OpenDocumentContentObservation(
                    paragraphs.Count,
                    paragraphs.Sum(paragraph => (long)paragraph.Text.Length),
                    paragraphs.Any(paragraph => string.Equals(paragraph.Text, EditMarker, StringComparison.Ordinal)));
            }
            case "ODS": {
                OdsDocument document = OdsDocument.Load(stream);
                EnsureStructurallyValid(document, format);
                OdsSheet sheet = document.Sheets.Single();
                long rowCount = 0;
                long checksum = 0;
                bool containsMarker = false;
                foreach (OdsRowRun row in sheet.RowRuns) {
                    rowCount += row.RepeatCount;
                    foreach (OdsCellRun cell in row.CellRuns) {
                        string value = cell.Value.ToString();
                        checksum += checked(row.RepeatCount * cell.RepeatCount * value.Length);
                        containsMarker |= string.Equals(value, EditMarker, StringComparison.Ordinal);
                    }
                }
                return new OpenDocumentContentObservation(rowCount, checksum, containsMarker);
            }
            case "ODP": {
                OdpPresentation presentation = OdpPresentation.Load(stream);
                EnsureStructurallyValid(presentation, format);
                long checksum = presentation.Slides.Sum(slide =>
                    slide.Shapes.OfType<OdpTextBox>().SelectMany(box => box.Paragraphs).Sum(paragraph => (long)paragraph.Text.Length));
                bool containsMarker = presentation.Slides.Any(slide =>
                    string.Equals(slide.Name, EditMarker, StringComparison.Ordinal)
                    || slide.Shapes.OfType<OdpTextBox>().SelectMany(box => box.Paragraphs)
                        .Any(paragraph => string.Equals(paragraph.Text, EditMarker, StringComparison.Ordinal)));
                return new OpenDocumentContentObservation(presentation.Slides.Count, checksum, containsMarker);
            }
            default:
                throw new ArgumentException("Unknown format: " + format, nameof(format));
        }
    }

    private static void EnsureStructurallyValid(OdfDocument document, string format) {
        OdfValidationResult validation = document.Validate();
        if (!validation.IsValid) {
            throw new InvalidOperationException(format + " produced invalid OpenDocument content: " +
                string.Join("; ", validation.Diagnostics.Where(diagnostic => diagnostic.Severity == OdfDiagnosticSeverity.Error)
                    .Take(5).Select(diagnostic => diagnostic.Message)));
        }
    }

    private static OpenDocumentEvidenceMeasurement RunChildProbe(
        string format,
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
        foreach (string argument in new[] { "--probe", format, operation, scale }) {
            startInfo.ArgumentList.Add(argument);
        }
        if (sourcePath != null) startInfo.ArgumentList.Add(sourcePath);

        using Process child = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Unable to start OpenDocument benchmark probe.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) {
            throw new InvalidOperationException($"Probe {format}/{operation}/{scale} failed: {error}");
        }
        return JsonSerializer.Deserialize<OpenDocumentEvidenceMeasurement>(output, JsonOptions)
            ?? throw new InvalidOperationException($"Probe {format}/{operation}/{scale} returned no measurement.");
    }

    private static string CreateTemporarySource(string format, OpenDocumentBenchmarkScale scale) {
        string path = Path.Combine(Path.GetTempPath(), $"OfficeIMO-OpenDocument-{Guid.NewGuid():N}.{format.ToLowerInvariant()}");
        File.WriteAllBytes(path, OpenDocumentBenchmarkCorpus.CreatePackage(format, scale));
        return path;
    }

    private static int ExpectedPrimaryCount(string format, OpenDocumentBenchmarkScale scale) =>
        format switch {
            "ODT" => scale.TextParagraphs,
            "ODS" => scale.SpreadsheetRows,
            "ODP" => scale.PresentationSlides,
            _ => throw new ArgumentException("Unknown format: " + format, nameof(format))
        };

    private static string[] Select(string[] values, string? filter, string label) {
        if (string.IsNullOrWhiteSpace(filter)) return values;
        string? selected = values.FirstOrDefault(value => string.Equals(value, filter, StringComparison.OrdinalIgnoreCase));
        return selected == null
            ? throw new ArgumentException($"Unknown {label} '{filter}'.")
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
            var startInfo = new ProcessStartInfo {
                FileName = "git",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            startInfo.ArgumentList.Add("rev-parse");
            startInfo.ArgumentList.Add("HEAD");
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
            var tracked = new ProcessStartInfo {
                FileName = "git",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            foreach (string argument in new[] { "diff", "--quiet", "HEAD", "--" }) tracked.ArgumentList.Add(argument);
            using Process trackedProcess = Process.Start(tracked)!;
            trackedProcess.WaitForExit();
            if (trackedProcess.ExitCode != 0) return true;

            var untracked = new ProcessStartInfo {
                FileName = "git",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            foreach (string argument in new[] { "ls-files", "--others", "--exclude-standard" }) untracked.ArgumentList.Add(argument);
            using Process untrackedProcess = Process.Start(untracked)!;
            string output = untrackedProcess.StandardOutput.ReadToEnd();
            untrackedProcess.WaitForExit();
            return untrackedProcess.ExitCode != 0 || !string.IsNullOrWhiteSpace(output);
        } catch {
            return true;
        }
    }
}

internal sealed record OpenDocumentOperationResult(byte[]? PackageBytes, long ObservedCount, long Checksum) {
    internal long OutputBytes => PackageBytes?.LongLength ?? 0;
}

internal sealed record OpenDocumentContentObservation(long ObservedCount, long Checksum, bool ContainsEditMarker);

internal sealed record OpenDocumentEvidenceMeasurement(
    string Format,
    string Operation,
    string Scale,
    int Iteration,
    int ExpectedCount,
    long InputBytes,
    long OutputBytes,
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long PeakWorkingSetBytes,
    long ObservedCount,
    long Checksum);

internal sealed record OpenDocumentEvidenceReport(
    DateTimeOffset MeasuredAtUtc,
    string SourceCommit,
    bool SourceTreeDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    IReadOnlyList<OpenDocumentEvidenceMeasurement> Measurements);
