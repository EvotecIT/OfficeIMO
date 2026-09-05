using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint.Pdf;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.PowerPoint.Benchmarks;

internal static class PowerPointBaselineRunner {
    private static readonly string[] Operations = {
        "CreateSave", "OpenEditSave", "OpenImageExport", "OpenPdfExport"
    };

    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int RunProbe(string[] args) {
        bool createSave = args.Length >= 1 && string.Equals(args[0],
            "CreateSave", StringComparison.OrdinalIgnoreCase);
        if ((args.Length != 2 && args.Length != 3)
            || !Operations.Contains(args[0], StringComparer.OrdinalIgnoreCase)
            || createSave && args.Length != 2
            || !createSave && args.Length != 3) {
            Console.Error.WriteLine(
                "Usage: --probe <CreateSave|OpenEditSave|OpenImageExport|OpenPdfExport> <Small|Normal|Large> [source.pptx]");
            return 2;
        }
        try {
            PowerPointBaselineMeasurement measurement = Measure(args[0],
                args[1], args.Length == 3 ? args[2] : null);
            Console.WriteLine(JsonSerializer.Serialize(measurement, JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    internal static int RunBaseline(string[] args) {
        bool verifyBudgets = args.Any(argument => string.Equals(argument,
            "--verify-budgets", StringComparison.OrdinalIgnoreCase));
        string? scaleFilter = GetOption(args, "--scale");
        string? operationFilter = GetOption(args, "--operation");
        string? jsonPath = GetOption(args, "--json");
        string? corpusDirectory = GetOption(args, "--corpus-dir");
        int repeat = GetPositiveIntOption(args, "--repeat", 1);
        IReadOnlyList<string> scales = string.IsNullOrWhiteSpace(scaleFilter)
            ? PowerPointBenchmarkCorpus.Scales
            : new[] { PowerPointBenchmarkCorpus.Get(scaleFilter!).Scale };
        PowerPointBenchmarkBudgetManifest? manifest = verifyBudgets
            ? PowerPointBenchmarkEvidence.LoadBudgetManifest()
            : null;
        IReadOnlyList<string> operations = manifest != null
            && string.IsNullOrWhiteSpace(operationFilter)
            ? manifest.Budgets.Select(budget => budget.Operation)
                .Distinct(StringComparer.OrdinalIgnoreCase).ToArray()
            : SelectOperations(operationFilter);
        var measurements = new List<PowerPointBaselineMeasurement>();
        var failures = new List<string>();
        foreach (string scale in scales) {
            string? sharedSourcePath = null;
            if (!string.IsNullOrWhiteSpace(corpusDirectory)) {
                string fullCorpusDirectory = Path.GetFullPath(corpusDirectory!);
                Directory.CreateDirectory(fullCorpusDirectory);
                PowerPointBenchmarkFixture fixture =
                    PowerPointBenchmarkCorpus.Get(scale);
                sharedSourcePath = Path.Combine(fullCorpusDirectory,
                    fixture.Scale + ".pptx");
                File.WriteAllBytes(sharedSourcePath,
                    PowerPointBenchmarkCorpus.CreatePackage(fixture));
            }
            foreach (string operation in operations) {
                for (var iteration = 1; iteration <= repeat; iteration++) {
                    PowerPointBaselineMeasurement measurement = RunChildProbe(
                        operation, scale, sharedSourcePath) with { Iteration = iteration };
                    measurements.Add(measurement);
                    if (manifest != null) {
                        PowerPointBenchmarkEvidence.EvaluateBudget(manifest,
                            measurement, failures);
                    }
                    Console.WriteLine(
                        $"{operation,-16} {scale,-6} #{iteration,-2} " +
                        $"{measurement.ElapsedMilliseconds,10:F1} ms " +
                        $"{measurement.AllocatedBytes / 1048576D,10:F1} MiB alloc " +
                        $"{measurement.PeakManagedHeapGrowthBytes / 1048576D,10:F1} MiB managed peak " +
                        $"{measurement.PeakWorkingSetBytes / 1048576D,10:F1} MiB process peak " +
                        $"{measurement.OutputBytes / 1048576D,10:F1} MiB output");
                }
            }
        }

        var report = new PowerPointBaselineReport(
            DateTimeOffset.UtcNow,
            PowerPointBenchmarkEvidence.ResolveCommit(),
            PowerPointBenchmarkEvidence.ResolveSourceTreeDirty(),
            RuntimeInformation.FrameworkDescription,
            RuntimeInformation.OSDescription,
            RuntimeInformation.ProcessArchitecture.ToString(),
            Environment.ProcessorCount,
            measurements,
            failures);
        if (!string.IsNullOrWhiteSpace(jsonPath)) {
            string fullPath = Path.GetFullPath(jsonPath!);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
            File.WriteAllText(fullPath, JsonSerializer.Serialize(report, JsonOptions));
            Console.WriteLine("Wrote " + fullPath);
        }
        foreach (string failure in failures) {
            Console.Error.WriteLine("BUDGET FAILURE: " + failure);
        }
        return failures.Count == 0 ? 0 : 1;
    }

    private static PowerPointBaselineMeasurement Measure(string operation,
        string scale, string? sourcePath) {
        PowerPointBenchmarkFixture fixture = PowerPointBenchmarkCorpus.Get(scale);
        long inputBytes = sourcePath == null
            ? 0L
            : new FileInfo(sourcePath).Length;

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long heapBefore = GC.GetTotalMemory(forceFullCollection: false);
        using var heapSampler = new PowerPointManagedHeapSampler();
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        PowerPointOperationResult result = Execute(operation, fixture,
            sourcePath);
        stopwatch.Stop();
        long peakManagedHeap = heapSampler.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        long peakWorkingSet = process.PeakWorkingSet64;
        ValidateResult(result, fixture, operation);
        return new PowerPointBaselineMeasurement(
            operation,
            fixture.Scale,
            1,
            fixture.SlideCount,
            result.ShapeCount,
            inputBytes,
            result.OutputBytes,
            stopwatch.Elapsed.TotalMilliseconds,
            allocated,
            Math.Max(0, peakManagedHeap - heapBefore),
            peakWorkingSet,
            result.InputAllocatedBytes,
            result.LoadAllocatedBytes,
            result.EditAllocatedBytes,
            result.SaveAllocatedBytes);
    }

    private static PowerPointOperationResult Execute(string operation,
        PowerPointBenchmarkFixture fixture, string? sourcePath) {
        if (string.Equals(operation, "CreateSave", StringComparison.OrdinalIgnoreCase)) {
            byte[] bytes = PowerPointBenchmarkCorpus.CreatePackage(fixture);
            return PowerPointOperationResult.Package(bytes);
        }
        if (sourcePath == null) throw new InvalidOperationException("Benchmark source package is unavailable.");
        long stageStart = GC.GetTotalAllocatedBytes(precise: true);
        byte[] source = File.ReadAllBytes(sourcePath);
        long afterInput = GC.GetTotalAllocatedBytes(precise: true);
        if (string.Equals(operation, "OpenEditSave", StringComparison.OrdinalIgnoreCase)) {
            using var input = new MemoryStream(source, writable: false);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(input);
            long afterLoad = GC.GetTotalAllocatedBytes(precise: true);
            for (int index = 0; index < presentation.Slides.Count; index += 10) {
                PowerPointTextBox edit = presentation.Slides[index].AddTextBoxPoints(
                    "Reviewed", 760, 486, 140, 22);
                edit.FontSize = 9;
                edit.Color = "166534";
            }
            long afterEdit = GC.GetTotalAllocatedBytes(precise: true);
            using var output = new MemoryStream();
            presentation.Save(output);
            byte[] package = output.ToArray();
            long afterSave = GC.GetTotalAllocatedBytes(precise: true);
            return PowerPointOperationResult.Package(package,
                afterInput - stageStart,
                afterLoad - afterInput,
                afterEdit - afterLoad,
                afterSave - afterEdit);
        }
        if (string.Equals(operation, "OpenImageExport", StringComparison.OrdinalIgnoreCase)) {
            using var input = new MemoryStream(source, writable: false);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(input);
            IReadOnlyList<OfficeImageExportResult> images = presentation.ExportImages(
                OfficeImageExportFormat.Png);
            return PowerPointOperationResult.ImageSet(images,
                presentation.Slides.Sum(slide => slide.Shapes.Count));
        }
        if (string.Equals(operation, "OpenPdfExport", StringComparison.OrdinalIgnoreCase)) {
            using var input = new MemoryStream(source, writable: false);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(input);
            byte[] pdf = presentation.ToPdfBytes();
            return PowerPointOperationResult.Pdf(pdf,
                presentation.Slides.Sum(slide => slide.Shapes.Count));
        }
        throw new ArgumentException("Unknown operation: " + operation, nameof(operation));
    }

    private static void ValidateResult(PowerPointOperationResult result,
        PowerPointBenchmarkFixture fixture, string operation) {
        if (result.OutputBytes <= 0) {
            throw new InvalidOperationException(operation + " produced no output.");
        }
        if (result.PackageBytes != null) {
            using PowerPointPresentation presentation = PowerPointPresentation.Load(
                new MemoryStream(result.PackageBytes, writable: false));
            if (presentation.Slides.Count != fixture.SlideCount) {
                throw new InvalidOperationException(
                    $"{operation} produced {presentation.Slides.Count} slides; expected {fixture.SlideCount}.");
            }
            OpenXmlValidator validator = new();
            PowerPointBenchmarkSemanticValidator.Validate(
                presentation.OpenXmlDocument, fixture.SlideCount, operation);
            string[] errors = validator.Validate(presentation.OpenXmlDocument)
                .Select(error => error.Description ?? error.ToString() ?? string.Empty).ToArray();
            if (errors.Length > 0) {
                throw new InvalidOperationException(operation +
                    " produced invalid Open XML: " + string.Join(" | ", errors.Take(5)));
            }
            int shapes = presentation.Slides.Sum(slide => slide.Shapes.Count);
            if (shapes < fixture.ExpectedMinimumShapeCount) {
                throw new InvalidOperationException(
                    $"{operation} produced {shapes} shapes; expected at least {fixture.ExpectedMinimumShapeCount}.");
            }
            result.ShapeCount = shapes;
        }
        if (result.Images != null) {
            if (result.Images.Count != fixture.SlideCount) {
                throw new InvalidOperationException(
                    $"Image export produced {result.Images.Count} images; expected {fixture.SlideCount}.");
            }
            if (result.Images.Any(image => image.EncodedLength <= 0 || image.Width <= 0 || image.Height <= 0)) {
                throw new InvalidOperationException("Image export produced an empty or dimensionless image.");
            }
            for (int index = 0; index < result.Images.Count; index++) {
                OfficeImageExportResult image = result.Images[index];
                OfficeImageExportDiagnostic[] failures = image.Diagnostics
                    .Where(diagnostic =>
                        diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error
                        || (diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Warning
                            && diagnostic.Code != OfficeImageExportDiagnosticCodes.FontSubstituted))
                    .ToArray();
                if (failures.Length > 0) {
                    throw new InvalidOperationException(
                        $"Image export slide {index + 1} reported: "
                        + string.Join(" | ", failures.Select(failure =>
                            failure.Code + ": " + failure.Message)));
                }
                if (!OfficePngReader.TryDecode(image.Bytes,
                        out OfficeRasterImage? raster) || raster == null) {
                    throw new InvalidOperationException(
                        $"Image export slide {index + 1} produced an undecodable PNG.");
                }
                ValidateRenderedCorpusImage(raster, index);
            }
        }
        if (result.PdfBytes != null) {
            ValidateRenderedCorpusPdf(result.PdfBytes, fixture);
        }
    }

    private static void ValidateRenderedCorpusPdf(byte[] pdf,
        PowerPointBenchmarkFixture fixture) {
        PdfCore.PdfReadDocument parsed = PdfCore.PdfReadDocument.Open(pdf);
        if (parsed.Pages.Count != fixture.SlideCount) {
            throw new InvalidOperationException(
                $"PDF export produced {parsed.Pages.Count} pages; expected {fixture.SlideCount}.");
        }
        for (int index = 0; index < parsed.Pages.Count; index++) {
            string text = parsed.Pages[index].ExtractText();
            RequirePdfText(text, $"Operational review {index + 1}", index);
            RequirePdfText(text, "OfficeIMO.PowerPoint performance corpus", index);
            if (index % 3 == 0) {
                foreach (string expected in new[] {
                             "Metric", "Current", "Target", "Quality",
                             "Coverage", "Latency"
                         }) {
                    RequirePdfText(text, expected, index);
                }
            }
            if (index % 5 == 0) {
                foreach (string expected in new[] {
                             "Actual", "Target", "Q1", "Q2", "Q3", "Q4"
                         }) {
                    RequirePdfText(text, expected, index);
                }
            }
        }

        IReadOnlyList<PdfCore.PdfPageRenderResult> rendered =
            PdfCore.PdfDocument.Load(pdf).Render.Pages(options:
                new PdfCore.PdfPageRenderOptions {
                    Dpi = 72D,
                    Format = PdfCore.PdfPageRenderFormat.Png,
                    MaxPages = fixture.SlideCount,
                    ContinueOnError = false,
                    MaxTotalOutputBytes = Math.Max(256L * 1024L * 1024L,
                        fixture.SlideCount * 4L * 1024L * 1024L)
                });
        if (rendered.Count != fixture.SlideCount) {
            throw new InvalidOperationException(
                $"PDF validation rendered {rendered.Count} pages; expected {fixture.SlideCount}.");
        }
        for (int index = 0; index < rendered.Count; index++) {
            PdfCore.PdfPageRenderResult page = rendered[index];
            if (!page.Succeeded || page.Bytes == null) {
                throw new InvalidOperationException(
                    $"PDF export page {index + 1} could not be rendered: "
                    + string.Join(" | ", page.Diagnostics));
            }
            if (!OfficePngReader.TryDecode(page.Bytes,
                    out OfficeRasterImage? raster) || raster == null) {
                throw new InvalidOperationException(
                    $"PDF export page {index + 1} produced an undecodable validation image.");
            }
            ValidateRenderedCorpusImage(raster, index);
        }
    }

    private static void RequirePdfText(string text, string expected,
        int slideIndex) {
        if (text.IndexOf(expected, StringComparison.Ordinal) < 0) {
            throw new InvalidOperationException(
                $"PDF export page {slideIndex + 1} lost expected content '{expected}'.");
        }
    }

    private static void ValidateRenderedCorpusImage(OfficeRasterImage image,
        int slideIndex) {
        OfficeColor background = slideIndex % 2 == 0
            ? OfficeColor.FromRgb(248, 250, 252)
            : OfficeColor.FromRgb(241, 245, 249);
        if (PowerPointBenchmarkVisualValidator.CountPixelsDifferentFrom(
                image, background,
                0, 0, 960, 540) < 1000) {
            throw new InvalidOperationException(
                $"Image export slide {slideIndex + 1} lost its visible corpus content.");
        }
        if (slideIndex % 3 == 0
            && PowerPointBenchmarkVisualValidator.CountPixelsDifferentFrom(
                image, background,
                40, 224, 300, 220) < 500) {
            throw new InvalidOperationException(
                $"Image export slide {slideIndex + 1} lost its rendered table region.");
        }
        if (slideIndex % 5 == 0
            && PowerPointBenchmarkVisualValidator.CountPixelsDifferentFrom(
                image, background,
                390, 214, 500, 260) < 500) {
            throw new InvalidOperationException(
                $"Image export slide {slideIndex + 1} lost its rendered chart region.");
        }
    }

    private static PowerPointBaselineMeasurement RunChildProbe(string operation,
        string scale, string? sharedSourcePath) {
        string? sourcePath = string.Equals(operation, "CreateSave",
            StringComparison.OrdinalIgnoreCase) ? null : sharedSourcePath;
        bool ownsSourcePath = false;
        if (!string.Equals(operation, "CreateSave",
                StringComparison.OrdinalIgnoreCase) && sourcePath == null) {
            sourcePath = Path.Combine(Path.GetTempPath(),
                "OfficeIMO-PowerPoint-Benchmark-" + Guid.NewGuid().ToString("N")
                + ".pptx");
            PowerPointBenchmarkFixture fixture =
                PowerPointBenchmarkCorpus.Get(scale);
            File.WriteAllBytes(sourcePath,
                PowerPointBenchmarkCorpus.CreatePackage(fixture));
            ownsSourcePath = true;
        }
        try {
            string processPath = Environment.ProcessPath
                ?? throw new InvalidOperationException("Unable to resolve benchmark process path.");
            var startInfo = new ProcessStartInfo {
                FileName = processPath,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            if (string.Equals(Path.GetFileNameWithoutExtension(processPath),
                    "dotnet", StringComparison.OrdinalIgnoreCase)) {
                startInfo.ArgumentList.Add(Assembly.GetEntryAssembly()!.Location);
            }
            startInfo.ArgumentList.Add("--probe");
            startInfo.ArgumentList.Add(operation);
            startInfo.ArgumentList.Add(scale);
            if (sourcePath != null) startInfo.ArgumentList.Add(sourcePath);
            using Process child = Process.Start(startInfo)
                ?? throw new InvalidOperationException("Unable to start benchmark probe process.");
            string output = child.StandardOutput.ReadToEnd();
            string error = child.StandardError.ReadToEnd();
            child.WaitForExit();
            if (child.ExitCode != 0) {
                throw new InvalidOperationException(
                    $"Probe {operation}/{scale} failed: {error}");
            }
            return JsonSerializer.Deserialize<PowerPointBaselineMeasurement>(
                    output, JsonOptions)
                ?? throw new InvalidOperationException(
                    $"Probe {operation}/{scale} returned no measurement.");
        } finally {
            if (ownsSourcePath && sourcePath != null && File.Exists(sourcePath)) {
                File.Delete(sourcePath);
            }
        }
    }

    private static string? GetOption(string[] args, string name) {
        int index = Array.FindIndex(args,
            item => string.Equals(item, name, StringComparison.OrdinalIgnoreCase));
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }

    private static IReadOnlyList<string> SelectOperations(string? filter) {
        if (string.IsNullOrWhiteSpace(filter)) return Operations;
        string? operation = Operations.FirstOrDefault(item =>
            string.Equals(item, filter, StringComparison.OrdinalIgnoreCase));
        return operation == null
            ? throw new ArgumentException("Unknown PowerPoint benchmark operation: " + filter)
            : new[] { operation };
    }

    private static int GetPositiveIntOption(string[] args, string name,
        int defaultValue) {
        string? value = GetOption(args, name);
        if (value == null) return defaultValue;
        return int.TryParse(value, out int parsed) && parsed > 0
            ? parsed
            : throw new ArgumentException(name + " must be a positive integer.");
    }
}

internal sealed record PowerPointBaselineMeasurement(
    string Operation,
    string Scale,
    int Iteration,
    int SlideCount,
    int ShapeCount,
    long InputBytes,
    long OutputBytes,
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long PeakManagedHeapGrowthBytes,
    long PeakWorkingSetBytes,
    long InputAllocatedBytes,
    long LoadAllocatedBytes,
    long EditAllocatedBytes,
    long SaveAllocatedBytes);

internal sealed record PowerPointBaselineReport(
    DateTimeOffset MeasuredAtUtc,
    string SourceCommit,
    bool SourceTreeDirty,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    IReadOnlyList<PowerPointBaselineMeasurement> Measurements,
    IReadOnlyList<string> Failures);

internal sealed class PowerPointOperationResult {
    private PowerPointOperationResult() { }

    internal byte[]? PackageBytes { get; private init; }
    internal IReadOnlyList<OfficeImageExportResult>? Images { get; private init; }
    internal byte[]? PdfBytes { get; private init; }
    internal long OutputBytes { get; private init; }
    internal int ShapeCount { get; set; }
    internal long InputAllocatedBytes { get; private init; }
    internal long LoadAllocatedBytes { get; private init; }
    internal long EditAllocatedBytes { get; private init; }
    internal long SaveAllocatedBytes { get; private init; }

    internal static PowerPointOperationResult Package(byte[] bytes,
        long inputAllocatedBytes = 0,
        long loadAllocatedBytes = 0,
        long editAllocatedBytes = 0,
        long saveAllocatedBytes = 0) => new() {
        PackageBytes = bytes,
        OutputBytes = bytes.LongLength,
        InputAllocatedBytes = inputAllocatedBytes,
        LoadAllocatedBytes = loadAllocatedBytes,
        EditAllocatedBytes = editAllocatedBytes,
        SaveAllocatedBytes = saveAllocatedBytes
    };

    internal static PowerPointOperationResult ImageSet(
        IReadOnlyList<OfficeImageExportResult> images, int shapeCount) => new() {
        Images = images,
        OutputBytes = images.Sum(image => image.EncodedLength),
        ShapeCount = shapeCount
    };

    internal static PowerPointOperationResult Pdf(byte[] bytes, int shapeCount) => new() {
        PdfBytes = bytes,
        OutputBytes = bytes.LongLength,
        ShapeCount = shapeCount
    };
}
