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
        if (args.Length != 2 || !Operations.Contains(args[0],
                StringComparer.OrdinalIgnoreCase)) {
            Console.Error.WriteLine(
                "Usage: --probe <CreateSave|OpenEditSave|OpenImageExport|OpenPdfExport> <Small|Normal|Large>");
            return 2;
        }
        try {
            PowerPointBaselineMeasurement measurement = Measure(args[0], args[1]);
            Console.WriteLine(JsonSerializer.Serialize(measurement, JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    internal static int RunBaseline(string[] args) {
        string? scaleFilter = GetOption(args, "--scale");
        string? jsonPath = GetOption(args, "--json");
        IReadOnlyList<string> scales = string.IsNullOrWhiteSpace(scaleFilter)
            ? PowerPointBenchmarkCorpus.Scales
            : new[] { PowerPointBenchmarkCorpus.Get(scaleFilter!).Scale };
        var measurements = new List<PowerPointBaselineMeasurement>();
        foreach (string scale in scales) {
            foreach (string operation in Operations) {
                PowerPointBaselineMeasurement measurement = RunChildProbe(operation, scale);
                measurements.Add(measurement);
                Console.WriteLine(
                    $"{operation,-16} {scale,-6} " +
                    $"{measurement.ElapsedMilliseconds,10:F1} ms " +
                    $"{measurement.AllocatedBytes / 1048576D,10:F1} MiB alloc " +
                    $"{measurement.PeakWorkingSetBytes / 1048576D,10:F1} MiB peak " +
                    $"{measurement.OutputBytes / 1048576D,10:F1} MiB output");
            }
        }

        var report = new PowerPointBaselineReport(
            DateTimeOffset.UtcNow,
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

    private static PowerPointBaselineMeasurement Measure(string operation, string scale) {
        PowerPointBenchmarkFixture fixture = PowerPointBenchmarkCorpus.Get(scale);
        byte[]? source = string.Equals(operation, "CreateSave",
            StringComparison.OrdinalIgnoreCase)
            ? null
            : PowerPointBenchmarkCorpus.CreatePackage(fixture);

        ValidateResult(Execute(operation, fixture, source), fixture, operation);
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        PowerPointOperationResult result = Execute(operation, fixture, source);
        stopwatch.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        ValidateResult(result, fixture, operation);
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        return new PowerPointBaselineMeasurement(
            operation,
            fixture.Scale,
            fixture.SlideCount,
            result.ShapeCount,
            source?.LongLength ?? 0L,
            result.OutputBytes,
            stopwatch.Elapsed.TotalMilliseconds,
            allocated,
            process.PeakWorkingSet64);
    }

    private static PowerPointOperationResult Execute(string operation,
        PowerPointBenchmarkFixture fixture, byte[]? source) {
        if (string.Equals(operation, "CreateSave", StringComparison.OrdinalIgnoreCase)) {
            byte[] bytes = PowerPointBenchmarkCorpus.CreatePackage(fixture);
            return PowerPointOperationResult.Package(bytes);
        }
        if (source == null) throw new InvalidOperationException("Benchmark source package is unavailable.");
        if (string.Equals(operation, "OpenEditSave", StringComparison.OrdinalIgnoreCase)) {
            using var input = new MemoryStream(source, writable: false);
            using PowerPointPresentation presentation = PowerPointPresentation.Load(input);
            for (int index = 0; index < presentation.Slides.Count; index += 10) {
                PowerPointTextBox edit = presentation.Slides[index].AddTextBoxPoints(
                    "Reviewed", 760, 486, 140, 22);
                edit.FontSize = 9;
                edit.Color = "166534";
            }
            using var output = new MemoryStream();
            presentation.Save(output);
            return PowerPointOperationResult.Package(output.ToArray());
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
            byte[] pdf = presentation.ToPdf();
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
        }
        if (result.PdfBytes != null) {
            PdfCore.PdfDocumentInfo info = PdfCore.PdfDocument.Open(result.PdfBytes).Inspect();
            if (info.PageCount != fixture.SlideCount) {
                throw new InvalidOperationException(
                    $"PDF export produced {info.PageCount} pages; expected {fixture.SlideCount}.");
            }
        }
    }

    private static PowerPointBaselineMeasurement RunChildProbe(string operation,
        string scale) {
        string processPath = Environment.ProcessPath
            ?? throw new InvalidOperationException("Unable to resolve benchmark process path.");
        var startInfo = new ProcessStartInfo {
            FileName = processPath,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            CreateNoWindow = true
        };
        if (string.Equals(Path.GetFileNameWithoutExtension(processPath), "dotnet",
                StringComparison.OrdinalIgnoreCase)) {
            startInfo.ArgumentList.Add(Assembly.GetEntryAssembly()!.Location);
        }
        startInfo.ArgumentList.Add("--probe");
        startInfo.ArgumentList.Add(operation);
        startInfo.ArgumentList.Add(scale);
        using Process child = Process.Start(startInfo)
            ?? throw new InvalidOperationException("Unable to start benchmark probe process.");
        string output = child.StandardOutput.ReadToEnd();
        string error = child.StandardError.ReadToEnd();
        child.WaitForExit();
        if (child.ExitCode != 0) {
            throw new InvalidOperationException(
                $"Probe {operation}/{scale} failed: {error}");
        }
        return JsonSerializer.Deserialize<PowerPointBaselineMeasurement>(output, JsonOptions)
            ?? throw new InvalidOperationException(
                $"Probe {operation}/{scale} returned no measurement.");
    }

    private static string? GetOption(string[] args, string name) {
        int index = Array.FindIndex(args,
            item => string.Equals(item, name, StringComparison.OrdinalIgnoreCase));
        return index >= 0 && index + 1 < args.Length ? args[index + 1] : null;
    }
}

internal sealed record PowerPointBaselineMeasurement(
    string Operation,
    string Scale,
    int SlideCount,
    int ShapeCount,
    long InputBytes,
    long OutputBytes,
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long PeakWorkingSetBytes);

internal sealed record PowerPointBaselineReport(
    DateTimeOffset MeasuredAtUtc,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    IReadOnlyList<PowerPointBaselineMeasurement> Measurements);

internal sealed class PowerPointOperationResult {
    private PowerPointOperationResult() { }

    internal byte[]? PackageBytes { get; private init; }
    internal IReadOnlyList<OfficeImageExportResult>? Images { get; private init; }
    internal byte[]? PdfBytes { get; private init; }
    internal long OutputBytes { get; private init; }
    internal int ShapeCount { get; set; }

    internal static PowerPointOperationResult Package(byte[] bytes) => new() {
        PackageBytes = bytes,
        OutputBytes = bytes.LongLength
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
