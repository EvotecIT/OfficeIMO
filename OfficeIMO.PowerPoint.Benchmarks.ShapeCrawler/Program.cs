using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.PowerPoint.Benchmarks;
using ShapeCrawler;

return ShapeCrawlerBaselineRunner.Run(args);

internal static class ShapeCrawlerBaselineRunner {
    private static readonly string[] Operations = { "CreateSave", "OpenEditSave" };
    private static readonly string[] Scales = { "Small", "Normal", "Large" };
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true
    };

    internal static int Run(string[] args) {
        if (args.Length > 0 && string.Equals(args[0], "--probe", StringComparison.OrdinalIgnoreCase)) {
            return RunProbe(args.Skip(1).ToArray());
        }

        string? scaleFilter = GetOption(args, "--scale");
        string? operationFilter = GetOption(args, "--operation");
        string? jsonPath = GetOption(args, "--json");
        string? corpusDirectory = GetOption(args, "--corpus-dir");
        int repeat = GetPositiveIntOption(args, "--repeat", 1);
        if (string.IsNullOrWhiteSpace(corpusDirectory)) {
            Console.Error.WriteLine(
                "--corpus-dir is required so OpenEditSave uses the exact same prebuilt input as the OfficeIMO lane.");
            return 2;
        }
        string fullCorpusDirectory = Path.GetFullPath(corpusDirectory!);
        IReadOnlyList<BenchmarkFixture> fixtures = string.IsNullOrWhiteSpace(scaleFilter)
            ? Scales.Select(GetFixture).ToArray()
            : new[] { GetFixture(scaleFilter!) };
        IReadOnlyList<string> operations = SelectOperations(operationFilter);
        var measurements = new List<BaselineMeasurement>();
        foreach (BenchmarkFixture fixture in fixtures) {
            string sourcePath = Path.Combine(fullCorpusDirectory,
                fixture.Scale + ".pptx");
            if (!File.Exists(sourcePath)) {
                Console.Error.WriteLine("Shared benchmark corpus was not found: "
                    + sourcePath);
                return 2;
            }
            foreach (string operation in operations) {
                for (var iteration = 1; iteration <= repeat; iteration++) {
                    BaselineMeasurement measurement = RunChildProbe(operation,
                        fixture.Scale, sourcePath) with { Iteration = iteration };
                    measurements.Add(measurement);
                    Console.WriteLine(
                        $"{operation,-12} {fixture.Scale,-6} #{iteration,-2} " +
                        $"{measurement.ElapsedMilliseconds,10:F1} ms " +
                        $"{measurement.AllocatedBytes / 1048576D,10:F1} MiB alloc " +
                        $"{measurement.PeakWorkingSetBytes / 1048576D,10:F1} MiB peak " +
                        $"{measurement.OutputBytes / 1048576D,10:F1} MiB output");
                }
            }
        }

        var report = new BaselineReport(
            DateTimeOffset.UtcNow,
            "ShapeCrawler",
            typeof(Presentation).Assembly.GetName().Version?.ToString() ?? "unknown",
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

    private static int RunProbe(string[] args) {
        bool createSave = args.Length >= 1 && string.Equals(args[0],
            "CreateSave", StringComparison.OrdinalIgnoreCase);
        if ((args.Length != 2 && args.Length != 3)
            || !Operations.Contains(args[0], StringComparer.OrdinalIgnoreCase)
            || createSave && args.Length != 2
            || !createSave && args.Length != 3) {
            Console.Error.WriteLine(
                "Usage: --probe <CreateSave|OpenEditSave> <Small|Normal|Large> [source.pptx]");
            return 2;
        }
        try {
            Console.WriteLine(JsonSerializer.Serialize(Measure(args[0],
                GetFixture(args[1]), args.Length == 3 ? args[2] : null),
                JsonOptions));
            return 0;
        } catch (Exception exception) {
            Console.Error.WriteLine(exception);
            return 1;
        }
    }

    private static BaselineMeasurement Measure(string operation,
        BenchmarkFixture fixture, string? sourcePath) {
        long inputBytes = sourcePath == null
            ? 0L
            : new FileInfo(sourcePath).Length;
        GC.Collect(2, GCCollectionMode.Forced, blocking: true, compacting: true);
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        byte[] result = Execute(operation, fixture, sourcePath);
        stopwatch.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        using Process process = Process.GetCurrentProcess();
        process.Refresh();
        long peakWorkingSet = process.PeakWorkingSet64;
        int shapeCount = Validate(result, fixture, operation);
        return new BaselineMeasurement(
            operation,
            fixture.Scale,
            1,
            fixture.SlideCount,
            shapeCount,
            inputBytes,
            result.LongLength,
            stopwatch.Elapsed.TotalMilliseconds,
            allocated,
            peakWorkingSet);
    }

    private static byte[] Execute(string operation, BenchmarkFixture fixture,
        string? sourcePath) {
        if (string.Equals(operation, "CreateSave", StringComparison.OrdinalIgnoreCase)) {
            return CreatePackage(fixture);
        }
        if (!string.Equals(operation, "OpenEditSave", StringComparison.OrdinalIgnoreCase) || sourcePath == null) {
            throw new ArgumentException("Unknown operation: " + operation, nameof(operation));
        }
        byte[] source = File.ReadAllBytes(sourcePath);
        using var input = new MemoryStream(source, writable: false);
        using var presentation = new Presentation(input);
        for (int index = 1; index <= presentation.Slides.Count; index += 10) {
            IUserSlideShapeCollection shapes = presentation.Slide(index).Shapes;
            shapes.AddTextBox(760, 486, 140, 22, "Reviewed");
            IShape edit = shapes.Last();
            edit.SetFontSize(9);
            edit.SetFontColor("166534");
        }
        using var output = new MemoryStream();
        presentation.Save(output);
        return output.ToArray();
    }

    private static byte[] CreatePackage(BenchmarkFixture fixture) {
        using var presentation = new Presentation();
        presentation.SlideWidth = 960;
        presentation.SlideHeight = 540;
        while (presentation.Slides.Count < fixture.SlideCount) {
            presentation.Slides.Add(1);
        }
        for (int index = 0; index < fixture.SlideCount; index++) {
            PopulateSlide(presentation.Slide(index + 1), index, fixture.SlideCount);
        }
        using var output = new MemoryStream();
        presentation.Save(output);
        return output.ToArray();
    }

    private static void PopulateSlide(IUserSlide slide, int index, int slideCount) {
        IUserSlideShapeCollection shapes = slide.Shapes;
        slide.Fill.SetColor(index % 2 == 0 ? "F8FAFC" : "F1F5F9");
        shapes.AddTextBox(40, 24, 600, 40, $"Operational review {index + 1}");
        IShape title = shapes.Last();
        title.SetFontSize(24);
        title.SetFontColor("0F172A");
        SetFirstRunBold(title);
        shapes.AddTextBox(40, 72, 700, 28,
            $"Slide {index + 1} of {slideCount} · deterministic benchmark corpus");
        IShape subtitle = shapes.Last();
        subtitle.SetFontSize(12);
        subtitle.SetFontColor("475569");
        string[] colors = { "DBEAFE", "DCFCE7", "FEF3C7", "FCE7F3" };
        for (int card = 0; card < 4; card++) {
            shapes.AddShape(40 + card * 220, 120, 190, 72,
                Geometry.Rectangle);
            IShape panel = shapes.Last();
            panel.Fill?.SetColor(colors[(card + index) % colors.Length]);
            panel.Outline?.SetHexColor("CBD5E1");
            if (panel.Outline != null) panel.Outline.Weight = 1;
        }

        if (index % 3 == 0) {
            shapes.AddTable(40, 224, 3, 4);
            IShape tableShape = shapes.Last();
            tableShape.Width = 300;
            tableShape.Height = 220;
            ITable table = tableShape.Table
                ?? throw new InvalidOperationException("ShapeCrawler did not expose the added table.");
            string[,] values = {
                { "Metric", "Current", "Target" },
                { "Quality", (92 + index % 7).ToString(), "98" },
                { "Coverage", (80 + index % 15).ToString(), "95" },
                { "Latency", (24 + index % 9).ToString(), "20" }
            };
            for (int row = 0; row < 4; row++) {
                for (int column = 0; column < 3; column++) {
                    ITableCell cell = table[row, column];
                    (cell.TextBox
                        ?? throw new InvalidOperationException("ShapeCrawler table cell has no text box."))
                        .SetText(values[row, column]);
                    if (row == 0) {
                        cell.Fill.SetColor("DBEAFE");
                        ITextBox textBox = cell.TextBox!;
                        textBox.Paragraphs.First().Portions.First().Font!.IsBold = true;
                    }
                }
            }
        } else {
            for (int row = 0; row < 3; row++) {
                shapes.AddShape(40, 224 + row * 72, 300, 54,
                    Geometry.Rectangle);
                IShape panel = shapes.Last();
                panel.Fill?.SetColor(row % 2 == 0 ? "FFFFFF" : "F8FAFC");
                panel.Outline?.SetHexColor("CBD5E1");
                shapes.AddTextBox(56, 239 + row * 72, 260, 24,
                    $"Workstream {row + 1}: checkpoint {index + row + 1}");
                IShape detail = shapes.Last();
                detail.SetFontSize(12);
                detail.SetFontColor("334155");
            }
        }

        if (index % 5 == 0) {
            var categories = new List<string> { "Q1", "Q2", "Q3", "Q4" };
            var series = new List<ShapeCrawler.Presentations.DraftChart.SeriesData> {
                new("Actual", new[] {
                    12D + index, 18D + index, 24D + index, 30D + index
                }),
                new("Target", new[] {
                    15D + index, 20D + index, 26D + index, 32D + index
                })
            };
            shapes.AddClusteredBarChart(390, 214, 500, 260, categories,
                series, string.Empty);
        } else {
            shapes.AddTextBox(390, 238, 500, 110,
                "Measured work includes editable text, vector shapes, package serialization, and rendering.");
            IShape narrative = shapes.Last();
            narrative.SetFontSize(18);
            narrative.SetFontColor("1E293B");
        }
        shapes.AddTextBox(40, 500, 420, 20,
            "OfficeIMO.PowerPoint performance corpus");
        IShape footer = shapes.Last();
        footer.SetFontSize(9);
        footer.SetFontColor("64748B");
    }

    private static void SetFirstRunBold(IShape shape) {
        ITextBox textBox = shape.TextBox
            ?? throw new InvalidOperationException("Expected a text-bearing shape.");
        textBox.Paragraphs.First().Portions.First().Font!.IsBold = true;
    }

    private static int Validate(byte[] package, BenchmarkFixture fixture, string operation) {
        if (package.Length == 0) throw new InvalidOperationException(operation + " produced no output.");
        using var input = new MemoryStream(package, writable: false);
        using var presentation = new Presentation(input);
        if (presentation.Slides.Count != fixture.SlideCount) {
            throw new InvalidOperationException(
                $"{operation} produced {presentation.Slides.Count} slides; expected {fixture.SlideCount}.");
        }
        int shapeCount = Enumerable.Range(1, presentation.Slides.Count)
            .Sum(index => presentation.Slide(index).Shapes.Count());
        if (shapeCount < fixture.ExpectedMinimumShapeCount) {
            throw new InvalidOperationException(
                $"{operation} produced {shapeCount} shapes; expected at least {fixture.ExpectedMinimumShapeCount}.");
        }
        using var packageStream = new MemoryStream(package, writable: false);
        using PresentationDocument document = PresentationDocument.Open(
            packageStream, false);
        PowerPointBenchmarkSemanticValidator.Validate(document,
            fixture.SlideCount, operation);
        string[] validationErrors = new OpenXmlValidator().Validate(document)
            .Select(error => error.Description ?? error.ToString()
                ?? string.Empty)
            .Take(5)
            .ToArray();
        if (validationErrors.Length > 0) {
            throw new InvalidOperationException(operation
                + " produced invalid Open XML: "
                + string.Join(" | ", validationErrors));
        }
        return shapeCount;
    }

    private static BaselineMeasurement RunChildProbe(string operation,
        string scale, string sharedSourcePath) {
        string? sourcePath = string.Equals(operation, "CreateSave",
            StringComparison.OrdinalIgnoreCase) ? null : sharedSourcePath;
        if (sourcePath != null && !File.Exists(sourcePath)) {
            throw new FileNotFoundException(
                "Shared benchmark corpus was not found.", sourcePath);
        }
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
        return JsonSerializer.Deserialize<BaselineMeasurement>(output,
                JsonOptions)
            ?? throw new InvalidOperationException(
                $"Probe {operation}/{scale} returned no measurement.");
    }

    private static BenchmarkFixture GetFixture(string scale) => scale.ToLowerInvariant() switch {
        "small" => new BenchmarkFixture("Small", 3, 18),
        "normal" => new BenchmarkFixture("Normal", 30, 180),
        "large" => new BenchmarkFixture("Large", 120, 720),
        _ => throw new ArgumentException("Scale must be Small, Normal, or Large.", nameof(scale))
    };

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
            ? throw new ArgumentException("Unknown ShapeCrawler benchmark operation: " + filter)
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

internal sealed record BenchmarkFixture(string Scale, int SlideCount, int ExpectedMinimumShapeCount);

internal sealed record BaselineMeasurement(
    string Operation,
    string Scale,
    int Iteration,
    int SlideCount,
    int ShapeCount,
    long InputBytes,
    long OutputBytes,
    double ElapsedMilliseconds,
    long AllocatedBytes,
    long PeakWorkingSetBytes);

internal sealed record BaselineReport(
    DateTimeOffset MeasuredAtUtc,
    string Library,
    string LibraryVersion,
    string Framework,
    string OperatingSystem,
    string Architecture,
    int ProcessorCount,
    IReadOnlyList<BaselineMeasurement> Measurements);
