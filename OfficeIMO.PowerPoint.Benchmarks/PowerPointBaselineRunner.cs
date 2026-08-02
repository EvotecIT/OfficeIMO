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
        string? scaleFilter = GetOption(args, "--scale");
        string? jsonPath = GetOption(args, "--json");
        string? corpusDirectory = GetOption(args, "--corpus-dir");
        IReadOnlyList<string> scales = string.IsNullOrWhiteSpace(scaleFilter)
            ? PowerPointBenchmarkCorpus.Scales
            : new[] { PowerPointBenchmarkCorpus.Get(scaleFilter!).Scale };
        var measurements = new List<PowerPointBaselineMeasurement>();
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
            foreach (string operation in Operations) {
                PowerPointBaselineMeasurement measurement = RunChildProbe(
                    operation, scale, sharedSourcePath);
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

    private static PowerPointBaselineMeasurement Measure(string operation,
        string scale, string? sourcePath) {
        PowerPointBenchmarkFixture fixture = PowerPointBenchmarkCorpus.Get(scale);
        long inputBytes = sourcePath == null
            ? 0L
            : new FileInfo(sourcePath).Length;

        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        PowerPointOperationResult result = Execute(operation, fixture,
            sourcePath);
        stopwatch.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        long peakWorkingSet = process.PeakWorkingSet64;
        ValidateResult(result, fixture, operation);
        return new PowerPointBaselineMeasurement(
            operation,
            fixture.Scale,
            fixture.SlideCount,
            result.ShapeCount,
            inputBytes,
            result.OutputBytes,
            stopwatch.Elapsed.TotalMilliseconds,
            allocated,
            peakWorkingSet);
    }

    private static PowerPointOperationResult Execute(string operation,
        PowerPointBenchmarkFixture fixture, string? sourcePath) {
        if (string.Equals(operation, "CreateSave", StringComparison.OrdinalIgnoreCase)) {
            byte[] bytes = PowerPointBenchmarkCorpus.CreatePackage(fixture);
            return PowerPointOperationResult.Package(bytes);
        }
        if (sourcePath == null) throw new InvalidOperationException("Benchmark source package is unavailable.");
        byte[] source = File.ReadAllBytes(sourcePath);
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
            PdfCore.PdfDocument.Open(pdf).Read.RenderPages(options:
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
        if (CountPixelsDifferentFrom(image, background,
                0, 0, 960, 540) < 1000) {
            throw new InvalidOperationException(
                $"Image export slide {slideIndex + 1} lost its visible corpus content.");
        }
        if (slideIndex % 3 == 0 && CountPixelsDifferentFrom(image, background,
                40, 224, 300, 220) < 500) {
            throw new InvalidOperationException(
                $"Image export slide {slideIndex + 1} lost its rendered table region.");
        }
        if (slideIndex % 5 == 0 && CountPixelsDifferentFrom(image, background,
                390, 214, 500, 260) < 500) {
            throw new InvalidOperationException(
                $"Image export slide {slideIndex + 1} lost its rendered chart region.");
        }
    }

    private static int CountPixelsDifferentFrom(OfficeRasterImage image,
        OfficeColor background, double leftPoints, double topPoints,
        double widthPoints, double heightPoints) {
        int left = Math.Max(0, (int)Math.Floor(leftPoints / 960D * image.Width));
        int top = Math.Max(0, (int)Math.Floor(topPoints / 540D * image.Height));
        int right = Math.Min(image.Width,
            (int)Math.Ceiling((leftPoints + widthPoints) / 960D * image.Width));
        int bottom = Math.Min(image.Height,
            (int)Math.Ceiling((topPoints + heightPoints) / 540D * image.Height));
        int different = 0;
        for (int y = top; y < bottom; y++) {
            for (int x = left; x < right; x++) {
                OfficeColor pixel = image.GetPixel(x, y);
                if (pixel.A > 0 && (Math.Abs(pixel.R - background.R)
                                    + Math.Abs(pixel.G - background.G)
                                    + Math.Abs(pixel.B - background.B)) > 12) {
                    different++;
                }
            }
        }
        return different;
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
