using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using HtmlTinkerX;
using Microsoft.Playwright;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Tests;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Produces reviewable H4 corpus evidence for the three distinct rendering intents:
/// browser screen, browser print, and OfficeIMO screen-to-PDF. PeachPDF remains an
/// isolated managed comparison and never enters an OfficeIMO runtime package.
/// </summary>
internal static partial class HtmlCorpusEvidenceRunner {
    private const int BrowserViewportWidth = 640;
    private const int BrowserViewportHeight = 900;
    private const double PdfRasterDpi = 96D;
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        WriteIndented = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    internal static async Task<int> RunAsync(string[] args) {
        ValidateArguments(args);
        if (args.Any(value => string.Equals(value, "--help", StringComparison.OrdinalIgnoreCase))) {
            WriteHelp();
            return 0;
        }

        string outputDirectory = ResolveOutputDirectory(args);
        using EvidenceOutputReservation outputReservation = EvidenceOutputReservation.Acquire(outputDirectory);
        string repositoryRoot = FindRepositoryRoot();
        bool requireCleanSource = HasFlag(args, "--require-clean-source");
        bool verifyAcceptance = HasFlag(args, "--verify-acceptance");
        bool worktreeDirty = IsGitDirty(repositoryRoot);
        if (requireCleanSource && worktreeDirty) {
            throw new InvalidOperationException("Clean, commit-addressable OfficeIMO source is required for H4 visual acceptance evidence.");
        }
        string? caseFilter = ReadOption(args, "--case");
        string corpusSelection = ReadOption(args, "--corpus") ?? "representative";
        HtmlCorpusEvidenceInputSet corpus = LoadCorpus(corpusSelection);
        HtmlCorpusEvidenceInput[] cases = corpus.Cases
            .Where(item => caseFilter == null || string.Equals(item.Scenario.Id, caseFilter, StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (cases.Length == 0) throw new ArgumentException("Unknown " + corpus.CorpusId + " case: " + caseFilter + ".");

        ExternalPdfRasterizer rasterizer = await ExternalPdfRasterizer.FindAsync().ConfigureAwait(false)
            ?? throw new InvalidOperationException(
                "The H4 corpus comparison requires pdftoppm on PATH so external PDFs are rasterized independently of OfficeIMO.Pdf.");

        await using HtmlBrowserSession browser = await HtmlPdfComparisonRenderers.OpenChromiumSessionAsync().ConfigureAwait(false);
        await browser.Page.SetViewportSizeAsync(BrowserViewportWidth, BrowserViewportHeight).ConfigureAwait(false);
        string chromiumVersion = browser.Browser?.Version ?? "unknown";

        var failures = new List<string>();
        var evidence = new List<HtmlCorpusCaseEvidence>(cases.Length);
        foreach (HtmlCorpusEvidenceInput input in cases) {
            Console.WriteLine("H4_CORPUS_CASE=" + input.Scenario.Id);
            HtmlCorpusCaseEvidence result = await RunCaseAsync(
                input, browser, rasterizer, outputDirectory).ConfigureAwait(false);
            evidence.Add(result);
            failures.AddRange(result.Failures.Select(failure => input.Scenario.Id + ": " + failure));
        }

        HtmlCorpusAcceptanceEvidence? acceptance = null;
        if (corpus.AcceptanceCorpus != null) {
            acceptance = HtmlCorpusAcceptanceEvaluator.Evaluate(corpus.AcceptanceCorpus, evidence);
        }
        if (verifyAcceptance) {
            if (caseFilter != null) {
                throw new InvalidOperationException("The H4 visual acceptance gate requires the complete selected corpus; remove --case.");
            }
            if (acceptance == null) {
                throw new InvalidOperationException("The selected corpus does not define a visual acceptance manifest.");
            }
            failures.AddRange(acceptance.Failures.Select(failure => "acceptance: " + failure));
        }

        var report = new HtmlCorpusEvidenceReport(
            SchemaVersion: 2,
            GeneratedUtc: DateTimeOffset.UtcNow,
            Environment: new HtmlCorpusEvidenceEnvironment(
                RuntimeInformation.OSDescription,
                RuntimeInformation.ProcessArchitecture.ToString(),
                RuntimeInformation.FrameworkDescription,
                AssemblyVersion(typeof(HtmlRenderEngine).Assembly),
                DependencyVersion("PeachPDF", typeof(PeachPDF.PdfGenerator).Assembly),
                DependencyVersion("HtmlTinkerX", typeof(HtmlBrowser).Assembly),
                chromiumVersion,
                rasterizer.Identity),
            Source: new HtmlCorpusEvidenceSource(
                corpus.CorpusId,
                corpus.RelativeRoot,
                corpus.ManifestSha256,
                ReadGit(repositoryRoot, "rev-parse", "HEAD"),
                worktreeDirty,
                cases.Length),
            Cases: evidence,
            Acceptance: acceptance,
            Failures: failures);

        string reportPath = Path.Combine(outputDirectory, "html-corpus-evidence.json");
        await File.WriteAllTextAsync(reportPath, JsonSerializer.Serialize(report, JsonOptions), new UTF8Encoding(false)).ConfigureAwait(false);
        string? acceptanceReportPath = null;
        if (acceptance != null) {
            acceptanceReportPath = Path.Combine(outputDirectory, "html-corpus-acceptance.md");
            await HtmlCorpusAcceptanceReportWriter.WriteAsync(acceptanceReportPath, acceptance, evidence).ConfigureAwait(false);
        }
        Console.WriteLine("HTML_CORPUS_EVIDENCE_REPORT=" + reportPath);
        if (acceptanceReportPath != null) Console.WriteLine("HTML_CORPUS_ACCEPTANCE_REPORT=" + acceptanceReportPath);
        Console.WriteLine("HTML_CORPUS_EVIDENCE_CASES=" + evidence.Count);
        if (acceptance != null) Console.WriteLine("HTML_CORPUS_ACCEPTANCE_STATUS=" + (acceptance.Passed ? "Passed" : "Failed"));
        Console.WriteLine("HTML_CORPUS_EVIDENCE_FAILURES=" + failures.Count);
        foreach (string failure in failures) Console.Error.WriteLine("EVIDENCE FAILURE: " + failure);
        return failures.Count == 0 ? 0 : 1;
    }

    private static async Task<HtmlCorpusCaseEvidence> RunCaseAsync(
        HtmlCorpusEvidenceInput input,
        HtmlBrowserSession browser,
        ExternalPdfRasterizer rasterizer,
        string outputDirectory) {
        HtmlRenderingCorpusCase scenario = input.Scenario;
        string caseDirectory = Path.Combine(outputDirectory, scenario.Id);
        Directory.CreateDirectory(caseDirectory);
        string sourcePath = Path.Combine(caseDirectory, "source.html");
        byte[] sourceBytes = input.SourceBytes;
        await File.WriteAllBytesAsync(sourcePath, sourceBytes).ConfigureAwait(false);
        var failures = new List<string>();

        HtmlCorpusStaticEvidence? officeImo = TryRender(
            () => RenderOfficeImo(input, caseDirectory, failures), "OfficeIMO", failures);
        HtmlCorpusPdfEvidence? peachPdf = await TryRenderAsync(
            () => RenderPeachPdfAsync(scenario, caseDirectory, rasterizer, failures), "PeachPDF", failures).ConfigureAwait(false);
        HtmlCorpusBrowserEvidence? chromium = await TryRenderAsync(
            () => RenderBrowserAsync(scenario, browser, rasterizer, caseDirectory, failures), "Chromium", failures).ConfigureAwait(false);

        HtmlCorpusTextComparison? chromiumTextComparison = officeImo != null && chromium != null
            ? await TryRenderAsync(
                () => CompareTextAsync(chromium.PrintPdf, officeImo.PrintPdf, caseDirectory, rasterizer),
                "OfficeIMO-to-Chromium text comparison", failures).ConfigureAwait(false)
            : null;
        HtmlCorpusTextComparison? peachTextComparison = officeImo != null && peachPdf != null
            ? await TryRenderAsync(
                () => CompareTextAsync(peachPdf.Pdf, officeImo.PrintPdf, caseDirectory, rasterizer),
                "OfficeIMO-to-PeachPDF text comparison", failures).ConfigureAwait(false)
            : null;
        HtmlCorpusGeometryComparison? screenGeometry = officeImo != null && chromium != null
            ? TryRender(() => CompareGeometry(officeImo.ScreenElements, chromium.ScreenElements),
                "screen geometry comparison", failures)
            : null;
        HtmlCorpusPixelComparison? screenPixels = officeImo != null && chromium != null
            ? TryRender(() => ComparePng(
                Path.Combine(caseDirectory, officeImo.ScreenPng.RelativePath),
                Path.Combine(caseDirectory, chromium.ScreenPng.RelativePath),
                caseDirectory,
                "screen-difference.png",
                HtmlCorpusPixelAlignment.TopLeftOverlap), "screen pixel comparison", failures)
            : null;
        HtmlCorpusScreenToPageComparison? screenToPage = officeImo != null
            ? TryRender(() => CompareScreenToPage(
                Path.Combine(caseDirectory, officeImo.ScreenPng.RelativePath),
                officeImo.ScreenToPdf.Pages,
                caseDirectory,
                "screen-to-page-difference.png"), "screen-to-page comparison", failures)
            : null;
        IReadOnlyList<HtmlCorpusPageComparison> chromiumPageComparisons = officeImo != null && chromium != null
            ? TryRender(
                () => ComparePdfPages(officeImo.PrintPdf.Pages, chromium.PrintPdf.Pages, caseDirectory, "print-chromium"),
                "OfficeIMO-to-Chromium page comparison", failures) ?? Array.Empty<HtmlCorpusPageComparison>()
            : Array.Empty<HtmlCorpusPageComparison>();
        IReadOnlyList<HtmlCorpusPageComparison> peachPageComparisons = officeImo != null && peachPdf != null
            ? TryRender(
                () => ComparePdfPages(officeImo.PrintPdf.Pages, peachPdf.Pdf.Pages, caseDirectory, "print-peachpdf"),
                "OfficeIMO-to-PeachPDF page comparison", failures) ?? Array.Empty<HtmlCorpusPageComparison>()
            : Array.Empty<HtmlCorpusPageComparison>();
        HtmlCorpusComparisonEvidence comparisons = new(
            chromiumTextComparison,
            peachTextComparison,
            screenGeometry,
            screenPixels,
            screenToPage,
            chromiumPageComparisons,
            peachPageComparisons);

        return new HtmlCorpusCaseEvidence(
            scenario.Id,
            input.SourceRelativePath,
            sourceBytes.LongLength,
            Sha256(sourceBytes),
            input.Capabilities,
            scenario.TextMarkers,
            officeImo,
            peachPdf,
            chromium,
            comparisons,
            failures);
    }

    private static HtmlCorpusStaticEvidence RenderOfficeImo(
        HtmlCorpusEvidenceInput input,
        string caseDirectory,
        ICollection<string> failures) {
        HtmlRenderingCorpusCase scenario = input.Scenario;
        using var sourceStream = new MemoryStream(input.SourceBytes, writable: false);
        HtmlConversionDocument source = HtmlConversionDocument.Load(sourceStream);
        HtmlToPdfOptions printOptions = new(scenario.CreateOptions());
        HtmlRenderRequest printRequest = HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf, printOptions);

        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        long operationAllocatedBefore = allocatedBefore;
        var operationStopwatch = Stopwatch.StartNew();
        HtmlPdfRenderRequestResult printResult = source.RenderToPdfResult(printRequest);
        byte[] printPdf = printResult.ToBytes();
        operationStopwatch.Stop();
        var printMetrics = new HtmlCorpusOperationMetrics(
            operationStopwatch.Elapsed.TotalMilliseconds,
            GC.GetTotalAllocatedBytes(precise: true) - operationAllocatedBefore,
            printPdf.LongLength);

        HtmlRenderOptions screenOptions = scenario.CreateOptions();
        screenOptions.Mode = HtmlRenderMode.Continuous;
        screenOptions.ViewportWidth = BrowserViewportWidth;
        screenOptions.ViewportHeight = BrowserViewportHeight;
        screenOptions.Margins = HtmlRenderMargins.All(0D);
        screenOptions.Scale = 1D;
        HtmlRenderRequest screenRequest = HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.Png, screenOptions);
        operationAllocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        operationStopwatch.Restart();
        HtmlRenderResult screenResult = HtmlRenderEngine.Execute(source, screenRequest);
        byte[] screenPng = RenderScenePagePng(screenResult.Document.Pages[0]);
        operationStopwatch.Stop();
        var screenMetrics = new HtmlCorpusOperationMetrics(
            operationStopwatch.Elapsed.TotalMilliseconds,
            GC.GetTotalAllocatedBytes(precise: true) - operationAllocatedBefore,
            screenPng.LongLength);

        HtmlToPdfOptions screenToPdfOptions = new(screenOptions) {
            PageSize = new OfficePageSize(
                BrowserViewportWidth / HtmlRenderOptions.CssPixelsPerInch,
                BrowserViewportHeight / HtmlRenderOptions.CssPixelsPerInch),
            Margins = HtmlRenderMargins.All(0D),
            HonorCssPageRules = false
        };
        HtmlRenderRequest screenToPdfRequest = HtmlRenderRequest.Create(
            HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.Pdf, screenToPdfOptions);
        operationAllocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        operationStopwatch.Restart();
        byte[] screenToPdf = source.RenderToPdfBytes(screenToPdfRequest);
        operationStopwatch.Stop();
        var screenToPdfMetrics = new HtmlCorpusOperationMetrics(
            operationStopwatch.Elapsed.TotalMilliseconds,
            GC.GetTotalAllocatedBytes(precise: true) - operationAllocatedBefore,
            screenToPdf.LongLength);
        stopwatch.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;

        string printPdfName = "officeimo-print.pdf";
        string screenPngName = "officeimo-screen.png";
        string screenToPdfName = "officeimo-screen-to-pdf.pdf";
        File.WriteAllBytes(Path.Combine(caseDirectory, printPdfName), printPdf);
        File.WriteAllBytes(Path.Combine(caseDirectory, screenPngName), screenPng);
        File.WriteAllBytes(Path.Combine(caseDirectory, screenToPdfName), screenToPdf);

        HtmlCorpusTextEvidence printText = ReadPdfText(printPdf, scenario.TextMarkers, failures, "OfficeIMO print PDF");
        HtmlCorpusTextEvidence screenText = ObserveText(screenResult.Document.Text, scenario.TextMarkers, failures, "OfficeIMO screen scene");
        HtmlCorpusOutputEvidence printOutput = CreatePdfEvidence(printPdfName, printPdf, caseDirectory, "officeimo-print");
        HtmlCorpusOutputEvidence screenOutput = CreatePngEvidence(screenPngName, screenPng);
        HtmlCorpusOutputEvidence screenToPdfOutput = CreatePdfEvidence(screenToPdfName, screenToPdf, caseDirectory, "officeimo-screen-to-pdf");
        HtmlCorpusSceneEvidence scene = WriteSceneEvidence(printResult.RenderResult.Document, caseDirectory);

        return new HtmlCorpusStaticEvidence(
            printOutput,
            screenOutput,
            screenToPdfOutput,
            scene,
            printText,
            screenText,
            ObserveGeometry(screenResult.Document),
            printMetrics,
            screenMetrics,
            screenToPdfMetrics,
            stopwatch.Elapsed.TotalMilliseconds,
            allocated,
            printRequest.ProfileId,
            screenRequest.ProfileId,
            screenToPdfRequest.ProfileId);
    }

    private static async Task<HtmlCorpusPdfEvidence> RenderPeachPdfAsync(
        HtmlRenderingCorpusCase scenario,
        string caseDirectory,
        ExternalPdfRasterizer rasterizer,
        ICollection<string> failures) {
        long allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
        var stopwatch = Stopwatch.StartNew();
        byte[] pdf = PeachPdfGenerator.Generate(scenario.Html);
        stopwatch.Stop();
        long allocated = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore;
        const string fileName = "peachpdf-print.pdf";
        File.WriteAllBytes(Path.Combine(caseDirectory, fileName), pdf);
        return new HtmlCorpusPdfEvidence(
            await CreateExternalPdfEvidenceAsync(fileName, pdf, caseDirectory, "peachpdf-print", rasterizer).ConfigureAwait(false),
            await ReadExternalPdfTextAsync(
                Path.Combine(caseDirectory, fileName), scenario.TextMarkers, failures, "PeachPDF print PDF", rasterizer,
                failOnMissing: false).ConfigureAwait(false),
            stopwatch.Elapsed.TotalMilliseconds,
            allocated);
    }

    private static async Task<HtmlCorpusBrowserEvidence> RenderBrowserAsync(
        HtmlRenderingCorpusCase scenario,
        HtmlBrowserSession browser,
        ExternalPdfRasterizer rasterizer,
        string caseDirectory,
        ICollection<string> failures) {
        var stopwatch = Stopwatch.StartNew();
        await browser.Page.EmulateMediaAsync(new PageEmulateMediaOptions { Media = Media.Screen }).ConfigureAwait(false);
        await HtmlPdfComparisonRenderers.PrepareChromiumPageAsync(browser, scenario.Html).ConfigureAwait(false);
        byte[] screenshot = await browser.Page.ScreenshotAsync(new PageScreenshotOptions {
            FullPage = true,
            Type = ScreenshotType.Png,
            Animations = ScreenshotAnimations.Disabled,
            Caret = ScreenshotCaret.Hide
        }).ConfigureAwait(false);
        string observationJson = await browser.Page.EvaluateAsync<string>(BrowserObservationScript).ConfigureAwait(false);
        HtmlCorpusBrowserObservation observation = JsonSerializer.Deserialize<HtmlCorpusBrowserObservation>(observationJson, JsonOptions)
            ?? throw new InvalidDataException("Chromium returned an empty page observation.");
        HtmlCorpusTextEvidence screenText = ObserveText(
            observation.Text, scenario.TextMarkers, failures, "Chromium screen observation", failOnMissing: false);

        await browser.Page.EmulateMediaAsync(new PageEmulateMediaOptions { Media = Media.Print }).ConfigureAwait(false);
        byte[] pdf = await HtmlPdfComparisonRenderers.CaptureChromiumPageAsync(browser).ConfigureAwait(false);
        stopwatch.Stop();

        const string pdfName = "chromium-print.pdf";
        const string screenshotName = "chromium-screen.png";
        File.WriteAllBytes(Path.Combine(caseDirectory, pdfName), pdf);
        File.WriteAllBytes(Path.Combine(caseDirectory, screenshotName), screenshot);
        _ = await ReadExternalPdfTextAsync(
            Path.Combine(caseDirectory, pdfName), scenario.TextMarkers, failures, "Chromium print PDF", rasterizer,
            failOnMissing: false).ConfigureAwait(false);
        return new HtmlCorpusBrowserEvidence(
            await CreateExternalPdfEvidenceAsync(pdfName, pdf, caseDirectory, "chromium-print", rasterizer).ConfigureAwait(false),
            CreatePngEvidence(screenshotName, screenshot),
            screenText,
            observation.Elements.Select(element => new HtmlCorpusElementGeometry(
                element.Key, element.Source, element.Index,
                element.X, element.Y, element.Width, element.Height)).ToArray(),
            stopwatch.Elapsed.TotalMilliseconds);
    }

    private static HtmlCorpusOutputEvidence CreatePdfEvidence(
        string relativePath,
        byte[] pdf,
        string caseDirectory,
        string rasterPrefix) {
        PdfCore.PdfDocumentInfo info = PdfCore.PdfDocument.Load(pdf).Inspect();
        IReadOnlyList<PdfCore.PdfPageRenderResult> renders = PdfCore.PdfDocument.Load(pdf).Render.Pages(
            options: new PdfCore.PdfPageRenderOptions {
                Format = PdfCore.PdfPageRenderFormat.Png,
                Dpi = PdfRasterDpi,
                ContinueOnError = false,
                MaxPages = Math.Max(1, info.PageCount)
            });
        var pages = new List<HtmlCorpusPageArtifact>(renders.Count);
        foreach (PdfCore.PdfPageRenderResult render in renders) {
            byte[] bytes = render.Bytes ?? throw new InvalidDataException(rasterPrefix + " page raster is empty.");
            string fileName = rasterPrefix + "-page-" + render.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".png";
            File.WriteAllBytes(Path.Combine(caseDirectory, fileName), bytes);
            pages.Add(new HtmlCorpusPageArtifact(
                render.PageNumber, fileName, render.Width, render.Height,
                bytes.LongLength, Sha256(bytes), render.Diagnostics));
        }
        return new HtmlCorpusOutputEvidence(
            relativePath, "application/pdf", pdf.LongLength, Sha256(pdf), info.PageCount, pages);
    }

    private static async Task<HtmlCorpusOutputEvidence> CreateExternalPdfEvidenceAsync(
        string relativePath,
        byte[] pdf,
        string caseDirectory,
        string rasterPrefix,
        ExternalPdfRasterizer rasterizer) {
        IReadOnlyList<HtmlPdfVisualEvidence> renders = await rasterizer.RenderAllPagesAsync(
            Path.Combine(caseDirectory, relativePath),
            caseDirectory,
            rasterPrefix,
            PdfRasterDpi).ConfigureAwait(false);
        HtmlCorpusPageArtifact[] pages = renders.Select(render => new HtmlCorpusPageArtifact(
            render.PageNumber,
            render.RelativePath,
            render.Width,
            render.Height,
            render.SizeBytes,
            render.Sha256,
            render.Diagnostics)).ToArray();
        return new HtmlCorpusOutputEvidence(
            relativePath, "application/pdf", pdf.LongLength, Sha256(pdf), pages.Length, pages);
    }

    private static T? TryRender<T>(Func<T> render, string owner, ICollection<string> failures) where T : class {
        try {
            return render();
        } catch (Exception exception) {
            failures.Add(owner + " evidence failed: " + exception.GetType().Name + ": " + exception.Message);
            return null;
        }
    }

    private static async Task<T?> TryRenderAsync<T>(
        Func<Task<T>> render,
        string owner,
        ICollection<string> failures) where T : class {
        try {
            return await render().ConfigureAwait(false);
        } catch (Exception exception) {
            failures.Add(owner + " evidence failed: " + exception.GetType().Name + ": " + exception.Message);
            return null;
        }
    }

    private static HtmlCorpusOutputEvidence CreatePngEvidence(string relativePath, byte[] png) {
        OfficeRasterImage image = DecodePng(png, relativePath);
        HtmlCorpusPageArtifact page = new(
            1, relativePath, image.Width, image.Height, png.LongLength, Sha256(png), Array.Empty<string>());
        return new HtmlCorpusOutputEvidence(relativePath, "image/png", png.LongLength, Sha256(png), 1, new[] { page });
    }

    private static HtmlCorpusSceneEvidence WriteSceneEvidence(HtmlRenderDocument document, string caseDirectory) {
        var rasterPages = new List<HtmlCorpusPageArtifact>(document.Pages.Count);
        var svgPages = new List<HtmlCorpusPageArtifact>(document.Pages.Count);
        foreach (HtmlRenderPage page in document.Pages) {
            byte[] png = RenderScenePagePng(page);
            string pngName = "officeimo-print-scene-page-" + page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".png";
            File.WriteAllBytes(Path.Combine(caseDirectory, pngName), png);
            OfficeRasterImage image = DecodePng(png, pngName);
            rasterPages.Add(new HtmlCorpusPageArtifact(
                page.PageNumber, pngName, image.Width, image.Height,
                png.LongLength, Sha256(png), Array.Empty<string>()));

            string svg = OfficeDrawingSvgExporter.ToSvg(page.CreateDrawing(), 1D);
            byte[] svgBytes = Encoding.UTF8.GetBytes(svg);
            string svgName = "officeimo-print-scene-page-" + page.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) + ".svg";
            File.WriteAllBytes(Path.Combine(caseDirectory, svgName), svgBytes);
            svgPages.Add(new HtmlCorpusPageArtifact(
                page.PageNumber, svgName,
                (int)Math.Ceiling(page.Width), (int)Math.Ceiling(page.Height),
                svgBytes.LongLength, Sha256(svgBytes), Array.Empty<string>()));
        }
        return new HtmlCorpusSceneEvidence(
            document.Pages.Count,
            document.Pages.Sum(page => Flatten(page.Scene).Count()),
            document.Diagnostics.Count,
            document.Diagnostics.Select(diagnostic => diagnostic.Code).Distinct(StringComparer.Ordinal).OrderBy(code => code, StringComparer.Ordinal).ToArray(),
            rasterPages,
            svgPages);
    }

    private static byte[] RenderScenePagePng(HtmlRenderPage page) =>
        OfficeDrawingRasterRenderer.ToPng(page.CreateDrawing(), 1D, OfficeColor.White);

    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static string AssemblyVersion(Assembly assembly) =>
        assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
        ?? assembly.GetName().Version?.ToString()
        ?? "unknown";

    private static string DependencyVersion(string packageId, Assembly assembly) {
        string? depsFiles = AppContext.GetData("APP_CONTEXT_DEPS_FILES") as string;
        foreach (string depsFile in (depsFiles ?? string.Empty).Split(Path.PathSeparator, StringSplitOptions.RemoveEmptyEntries)) {
            if (!File.Exists(depsFile)) continue;
            try {
                using JsonDocument document = JsonDocument.Parse(File.ReadAllBytes(depsFile));
                if (!document.RootElement.TryGetProperty("libraries", out JsonElement libraries)) continue;
                string prefix = packageId + "/";
                foreach (JsonProperty library in libraries.EnumerateObject()) {
                    if (library.Name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) {
                        return library.Name[prefix.Length..];
                    }
                }
            } catch (JsonException) {
                // Fall back to assembly metadata when a host provides an unrelated or malformed deps file.
            }
        }
        return AssemblyVersion(assembly);
    }

    private static string ResolveOutputDirectory(string[] args) {
        string? configured = ReadOption(args, "--output");
        if (!string.IsNullOrWhiteSpace(configured)) return Path.GetFullPath(configured);
        string run = DateTime.UtcNow.ToString("yyyyMMdd-HHmmss.fff", System.Globalization.CultureInfo.InvariantCulture)
            + "-" + Environment.ProcessId.ToString(System.Globalization.CultureInfo.InvariantCulture);
        return Path.Combine(Path.GetTempPath(), "OfficeIMO", "HtmlCorpusEvidence", run);
    }

    private static string? ReadOption(string[] args, string option) {
        for (int index = 1; index < args.Length; index++) {
            if (!string.Equals(args[index], option, StringComparison.OrdinalIgnoreCase)) continue;
            if (index == args.Length - 1 || args[index + 1].StartsWith("--", StringComparison.Ordinal)) {
                throw new ArgumentException(option + " requires a value.");
            }
            return args[index + 1];
        }
        return null;
    }

    private static bool HasFlag(string[] args, string option) =>
        args.Any(argument => string.Equals(argument, option, StringComparison.OrdinalIgnoreCase));

    private static void ValidateArguments(string[] args) {
        for (int index = 1; index < args.Length; index++) {
            string argument = args[index];
            if (string.Equals(argument, "--help", StringComparison.OrdinalIgnoreCase)
                || string.Equals(argument, "--require-clean-source", StringComparison.OrdinalIgnoreCase)
                || string.Equals(argument, "--verify-acceptance", StringComparison.OrdinalIgnoreCase)) continue;
            if (string.Equals(argument, "--output", StringComparison.OrdinalIgnoreCase)
                || string.Equals(argument, "--case", StringComparison.OrdinalIgnoreCase)
                || string.Equals(argument, "--corpus", StringComparison.OrdinalIgnoreCase)) {
                if (++index >= args.Length || args[index].StartsWith("--", StringComparison.Ordinal)) {
                    throw new ArgumentException(argument + " requires a value.");
                }
                continue;
            }
            throw new ArgumentException("Unknown html-corpus-evidence option: " + argument);
        }
    }

    private static string FindRepositoryRoot() {
        foreach (string seed in new[] { Directory.GetCurrentDirectory(), AppContext.BaseDirectory }) {
            string? current = Path.GetFullPath(seed);
            while (!string.IsNullOrWhiteSpace(current)) {
                if (File.Exists(Path.Combine(current, "OfficeIMO.sln"))) return current;
                current = Directory.GetParent(current)?.FullName;
            }
        }
        throw new DirectoryNotFoundException("Could not locate the OfficeIMO repository root.");
    }

    private static string? ReadGit(string repositoryRoot, params string[] arguments) {
        try {
            var info = new ProcessStartInfo("git") {
                WorkingDirectory = repositoryRoot,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };
            foreach (string argument in arguments) info.ArgumentList.Add(argument);
            using Process process = Process.Start(info)!;
            string output = process.StandardOutput.ReadToEnd().Trim();
            process.WaitForExit();
            return process.ExitCode == 0 && output.Length > 0 ? output : null;
        } catch {
            return null;
        }
    }

    private static bool IsGitDirty(string repositoryRoot) =>
        !string.IsNullOrWhiteSpace(ReadGit(repositoryRoot, "status", "--porcelain", "--untracked-files=normal"));

    private static void WriteHelp() {
        Console.WriteLine("html-corpus-evidence [--corpus <representative|advanced-held-out>] [--case <id>] [--output <new-directory>] [--verify-acceptance] [--require-clean-source]");
        Console.WriteLine("Captures every H4 source through OfficeIMO print, screen and screen-to-PDF; PeachPDF print; and Chromium screen and print. It writes all-page PDF rasters, OfficeIMO scene PNG/SVG files, text, geometry, pixel comparisons, and the advanced held-out per-capability acceptance result.");
    }

    private static HtmlCorpusEvidenceInputSet LoadCorpus(string selection) {
        if (string.Equals(selection, "representative", StringComparison.OrdinalIgnoreCase)) {
            return new HtmlCorpusEvidenceInputSet(
                "officeimo-html-h4-representative",
                HtmlRenderingRepresentativeCorpus.RelativeRoot,
                null,
                null,
                HtmlRenderingRepresentativeCorpus.All.Select(scenario => new HtmlCorpusEvidenceInput(
                    scenario,
                    scenario.SourceRelativePath,
                    HtmlMarketScenarioCatalog.Get(scenario.Id).Capabilities,
                    Encoding.UTF8.GetBytes(scenario.Html))).ToArray());
        }
        if (string.Equals(selection, "advanced-held-out", StringComparison.OrdinalIgnoreCase)) {
            HtmlRenderingAdvancedHeldOutCorpus corpus = HtmlRenderingAdvancedHeldOutCorpus.Load();
            return new HtmlCorpusEvidenceInputSet(
                corpus.Manifest.CorpusId,
                HtmlRenderingAdvancedHeldOutCorpus.RelativeRoot,
                corpus.ManifestSha256,
                corpus,
                corpus.Cases.Select(item => new HtmlCorpusEvidenceInput(
                    new HtmlRenderingCorpusCase(
                        item.Id,
                        HtmlRenderMode.Paged,
                        item.Html,
                        item.Manifest.TextMarkers,
                        expectedPageCount: item.Manifest.ExpectedPrintPageCount,
                        minimumVisualCount: 1,
                        minimumHeadingCount: 0),
                    item.SourceRelativePath,
                    item.Manifest.Capabilities,
                    item.SourceBytes)).ToArray());
        }
        throw new ArgumentException("Unknown H4 corpus selection: " + selection + ". Use representative or advanced-held-out.");
    }

    private sealed record HtmlCorpusEvidenceInput(
        HtmlRenderingCorpusCase Scenario,
        string SourceRelativePath,
        IReadOnlyList<string> Capabilities,
        byte[] SourceBytes);

    private sealed record HtmlCorpusEvidenceInputSet(
        string CorpusId,
        string RelativeRoot,
        string? ManifestSha256,
        HtmlRenderingAdvancedHeldOutCorpus? AcceptanceCorpus,
        IReadOnlyList<HtmlCorpusEvidenceInput> Cases);

    private const string BrowserObservationScript = """
        () => {
          const selected = [...document.querySelectorAll('h1,h2,table,form,svg,input,select,textarea,button')];
          const counts = new Map();
          const elements = selected.map(element => {
            const source = element.tagName.toLowerCase() + (element.id ? '#' + element.id : '');
            const index = counts.get(source) || 0;
            counts.set(source, index + 1);
            const rect = element.getBoundingClientRect();
            return {
              key: source + ':' + index,
              source,
              index,
              x: rect.x + window.scrollX,
              y: rect.y + window.scrollY,
              width: rect.width,
              height: rect.height
            };
          });
          return JSON.stringify({
            text: document.body ? document.body.innerText : '',
            scrollWidth: document.documentElement.scrollWidth,
            scrollHeight: document.documentElement.scrollHeight,
            elements
          });
        }
        """;
}
