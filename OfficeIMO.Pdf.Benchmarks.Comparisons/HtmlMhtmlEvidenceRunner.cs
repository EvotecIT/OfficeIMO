using System.Security.Cryptography;
using System.Diagnostics;
using System.Text;
using System.Text.Json;
using HtmlTinkerX;
using Microsoft.Playwright;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Mhtml;
using PeachPDF;
using PeachPDF.Network;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Captures one independently sourced page and records distinct PDF intents from
/// the same frozen MHTML bytes. This is an opt-in qualification tool, not a runtime route.
/// </summary>
internal static class HtmlMhtmlEvidenceRunner {
    private const int ViewportWidth = 816;
    private const int ViewportHeight = 900;
    private const int MaximumArchiveBytes = 32 * 1024 * 1024;
    private static readonly JsonSerializerOptions JsonOptions = new() {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        WriteIndented = true
    };

    internal static async Task<int> RunAsync(string[] args) {
        string[] flags = args.Skip(5).ToArray();
        if (args.Length < 5 || (args[1] != "--url" && args[1] != "--mhtml") || args[3] != "--output"
            || flags.Distinct(StringComparer.Ordinal).Count() != flags.Length
            || flags.Any(flag => flag != "--replay-browser" && flag != "--require-clean-source")) {
            Console.Error.WriteLine("html-mhtml-evidence <--url https-url|--mhtml existing-archive> --output <new-directory> [--replay-browser] [--require-clean-source]");
            return 2;
        }
        bool replay = args[1] == "--mhtml";
        bool replayBrowser = replay && flags.Contains("--replay-browser", StringComparer.Ordinal);
        if (!replay && flags.Contains("--replay-browser", StringComparer.Ordinal))
            throw new ArgumentException("--replay-browser requires --mhtml.");
        Uri? url = replay ? null : new Uri(args[2], UriKind.Absolute);
        if (url != null && url.Scheme != Uri.UriSchemeHttps) throw new ArgumentException("Only HTTPS page URLs are supported.");
        string repositoryRoot = FindRepositoryRoot();
        string sourceCommit = GitOutput(repositoryRoot, "rev-parse", "HEAD");
        bool worktreeDirty = GitOutput(repositoryRoot, "status", "--porcelain", "--untracked-files=normal").Length > 0;
        var ownerVersions = new Dictionary<string, string?> {
            ["Core"] = InformationalVersion(typeof(OfficeFontFaceCollection).Assembly),
            ["Email"] = InformationalVersion(typeof(OfficeIMO.Email.EmailDocumentReader).Assembly),
            ["HTML Core"] = InformationalVersion(typeof(OfficeIMO.Html.Dom.HtmlDocument).Assembly),
            ["HTML AngleSharp"] = InformationalVersion(typeof(OfficeIMO.Html.Providers.AngleSharpHtmlParser).Assembly),
            ["HTML"] = InformationalVersion(typeof(HtmlRenderEngine).Assembly),
            ["HTML PDF"] = InformationalVersion(typeof(HtmlToPdfOptions).Assembly),
            ["MHTML"] = InformationalVersion(typeof(MhtmlDocument).Assembly),
            ["MHTML PDF"] = InformationalVersion(typeof(MhtmlPdfConverterExtensions).Assembly),
            ["PDF"] = InformationalVersion(typeof(OfficeIMO.Pdf.PdfDocument).Assembly),
            ["evidence runner"] = InformationalVersion(typeof(HtmlMhtmlEvidenceRunner).Assembly)
        };
        if (flags.Contains("--require-clean-source", StringComparer.Ordinal)
            && (worktreeDirty || ownerVersions.Values.Any(version => version == null
                || !version.EndsWith("+" + sourceCommit, StringComparison.OrdinalIgnoreCase)))) {
            throw new InvalidOperationException("Evidence requires a clean checkout and all owner assemblies rebuilt from its exact HEAD commit.");
        }
        string output = Path.GetFullPath(args[4]);
        if (Directory.Exists(output) || File.Exists(output)) throw new IOException("Output directory already exists.");
        Directory.CreateDirectory(output);

        var results = new List<OperationEvidence>();
        var failures = new List<string>();
        string? finalUrl = null;
        string? chromiumVersion = null;
        byte[] archive;
        if (replay) {
            archive = await File.ReadAllBytesAsync(Path.GetFullPath(args[2])).ConfigureAwait(false);
            if (archive.Length == 0 || archive.Length > MaximumArchiveBytes)
                throw new InvalidDataException("MHTML input is empty or exceeds the 32 MiB input limit.");
            await File.WriteAllBytesAsync(Path.Combine(output, "source.mhtml"), archive).ConfigureAwait(false);
            if (replayBrowser) {
                try {
                    (chromiumVersion, finalUrl) = await CaptureOfflineBrowserOutputsAsync(output, results, failures).ConfigureAwait(false);
                } catch (Exception exception) {
                    failures.Add("chromium-archive-replay: " + exception);
                }
            }
        } else {
            await using HtmlBrowserSession browser = await HtmlPdfComparisonRenderers.OpenChromiumSessionAsync().ConfigureAwait(false);
            await browser.Page.SetViewportSizeAsync(ViewportWidth, ViewportHeight).ConfigureAwait(false);
            chromiumVersion = browser.Browser?.Version;
            IResponse? response = await browser.Page.GotoAsync(url!.ToString(), new PageGotoOptions {
                WaitUntil = WaitUntilState.DOMContentLoaded,
                Timeout = 30000
            }).ConfigureAwait(false);
            if (response == null || !response.Ok) throw new InvalidOperationException("Page navigation failed: " + response?.Status);
            finalUrl = browser.Page.Url;
            await browser.Page.WaitForTimeoutAsync(1000).ConfigureAwait(false);
            ICDPSession session = await browser.Page.Context.NewCDPSessionAsync(browser.Page).ConfigureAwait(false);
            try {
                JsonElement? result = await session.SendAsync("Page.captureSnapshot", new Dictionary<string, object> {
                    ["format"] = "mhtml"
                }).ConfigureAwait(false);
                if (!result.HasValue || !result.Value.TryGetProperty("data", out JsonElement data))
                    throw new InvalidDataException("Chromium did not return MHTML snapshot data.");
                archive = Encoding.UTF8.GetBytes(data.GetString() ?? string.Empty);
            } finally {
                await session.DetachAsync().ConfigureAwait(false);
            }
            if (archive.Length == 0 || archive.Length > MaximumArchiveBytes)
                throw new InvalidDataException("MHTML snapshot is empty or exceeds the 32 MiB input limit.");
            await File.WriteAllBytesAsync(Path.Combine(output, "source.mhtml"), archive).ConfigureAwait(false);

            try {
                (chromiumVersion, _) = await CaptureOfflineBrowserOutputsAsync(output, results, failures).ConfigureAwait(false);
            } catch (Exception exception) {
                failures.Add("chromium-archive-replay: " + exception);
            }
        }

        MhtmlDocument? document = null;
        try {
            using var source = new MemoryStream(archive, writable: false);
            document = MhtmlDocument.Load(source);
        } catch (Exception exception) {
            failures.Add("OfficeIMO MHTML load: " + exception);
        }
        if (document != null) {
            await RunConversionAsync("officeimo-print", () => document.ToPdfDocumentResultAsync(), output, results, failures).ConfigureAwait(false);
            await RunConversionAsync("officeimo-print-zero-margin", () => document.ToPdfDocumentResultAsync(new HtmlToPdfOptions {
                Margins = HtmlRenderMargins.All(0)
            }), output, results, failures).ConfigureAwait(false);
            await RunConversionAsync("officeimo-print-browser-ua", () => {
                var options = new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0) };
                options.UseBrowserUserAgentStyles();
                return document.ToPdfDocumentResultAsync(options);
            }, output, results, failures).ConfigureAwait(false);
            await RunConversionAsync("officeimo-print-zero-margin-local-fonts", () => {
                var options = new HtmlToPdfOptions { Margins = HtmlRenderMargins.All(0) };
                options.ResourcePolicy.AllowDocumentFontEmbedding = true;
                return document.ToPdfDocumentResultAsync(options);
            }, output, results, failures).ConfigureAwait(false);
            await RunConversionAsync("officeimo-screen-media", () => RenderScreenPdfAsync(document, HtmlRenderIntentProfile.ScreenMediaPaged), output, results, failures).ConfigureAwait(false);
            await RunConversionAsync("officeimo-screen-snapshot", () => RenderScreenPdfAsync(document, HtmlRenderIntentProfile.ScreenSnapshotPaged), output, results, failures).ConfigureAwait(false);
        }
        await RunAsync("peachpdf-print", async () => {
            using var source = new MemoryStream(archive, writable: false);
            var generator = new PdfGenerator();
            var configuration = new PdfGenerateConfig {
                PageSize = PeachPDF.PageSize.A4,
                PageOrientation = PageOrientation.Portrait,
                EnableTaggedPdf = true,
                AllowLocalFileAccess = false,
                NetworkLoader = new MimeKitNetworkLoader(source)
            };
            var pdf = await generator.GeneratePdf(null, configuration).ConfigureAwait(false);
            using var outputStream = new MemoryStream();
            pdf.Save(outputStream);
            return outputStream.ToArray();
        }, output, results, failures).ConfigureAwait(false);

        var report = new {
            schemaVersion = 2,
            sourceUrl = url?.ToString() ?? document?.BaseUri.ToString(),
            finalUrl,
            runUtc = DateTimeOffset.UtcNow,
            sourceMode = replay ? "frozen-mhtml-replay" : "live-browser-capture",
            readinessPolicy = replay ? "offline archive navigation at DOMContentLoaded" : "live navigation at DOMContentLoaded plus 1000 ms",
            browserReference = replay && !replayBrowser ? "not-recorded" : "offline-archive-replay",
            sourceCommit,
            worktreeDirty,
            chromiumVersion,
            peachPdfVersion = HtmlCorpusEvidenceRunner.DependencyVersion("PeachPDF", typeof(PdfGenerator).Assembly),
            ownerVersions,
            viewport = new { width = ViewportWidth, height = ViewportHeight },
            archive = new { path = "source.mhtml", bytes = archive.Length, sha256 = Sha256(archive) },
            resourceCount = document?.Resources.Count,
            mimeDiagnostics = document?.MimeDiagnostics.Select(diagnostic => new {
                diagnostic.Code,
                diagnostic.Message,
                diagnostic.Location,
                severity = diagnostic.Severity.ToString()
            }).ToArray(),
            operations = results,
            failures
        };
        string reportPath = Path.Combine(output, "html-mhtml-evidence.json");
        await File.WriteAllTextAsync(reportPath, JsonSerializer.Serialize(report, JsonOptions)).ConfigureAwait(false);
        Console.WriteLine("HTML_MHTML_EVIDENCE_REPORT=" + reportPath);
        foreach (string failure in failures) Console.Error.WriteLine("EVIDENCE FAILURE: " + failure);
        return failures.Count == 0 ? 0 : 1;
    }

    private static Task<PdfCore.PdfDocumentConversionResult> RenderScreenPdfAsync(MhtmlDocument document, HtmlRenderIntentProfile profile) {
        var options = new HtmlToPdfOptions {
            ViewportWidth = ViewportWidth,
            ViewportHeight = ViewportHeight,
            Margins = HtmlRenderMargins.All(0),
            HonorCssPageRules = false
        };
        if (profile == HtmlRenderIntentProfile.ScreenSnapshotPaged) {
            options.PageSize = new OfficePageSize(
                ViewportWidth / HtmlRenderOptions.CssPixelsPerInch,
                ViewportHeight / HtmlRenderOptions.CssPixelsPerInch);
        }
        HtmlRenderRequest request = HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Pdf, options);
        return document.RenderToPdfDocumentResultAsync(request);
    }

    private static async Task<(string? ChromiumVersion, string FinalUrl)> CaptureOfflineBrowserOutputsAsync(
        string output,
        ICollection<OperationEvidence> results,
        ICollection<string> failures) {
        await using HtmlBrowserSession browser = await HtmlPdfComparisonRenderers.OpenChromiumSessionAsync().ConfigureAwait(false);
        await browser.Page.SetViewportSizeAsync(ViewportWidth, ViewportHeight).ConfigureAwait(false);
        await browser.Page.Context.SetOfflineAsync(true).ConfigureAwait(false);
        string archiveUrl = new Uri(Path.Combine(output, "source.mhtml")).AbsoluteUri;
        await NavigateArchiveAsync(browser, archiveUrl).ConfigureAwait(false);

        try {
            await browser.Page.EmulateMediaAsync(new PageEmulateMediaOptions { Media = Media.Screen }).ConfigureAwait(false);
            byte[] screenshot = await browser.Page.ScreenshotAsync(new PageScreenshotOptions {
                FullPage = true, Type = ScreenshotType.Png,
                Animations = ScreenshotAnimations.Disabled
            }).ConfigureAwait(false);
            await File.WriteAllBytesAsync(Path.Combine(output, "chromium-screen.png"), screenshot).ConfigureAwait(false);
        } catch (Exception exception) {
            failures.Add("chromium-screen-screenshot: " + exception);
        }
        await RunAsync("chromium-screen", async () => {
            await browser.Page.EmulateMediaAsync(new PageEmulateMediaOptions { Media = Media.Screen }).ConfigureAwait(false);
            return await HtmlPdfComparisonRenderers.CaptureChromiumPageAsync(browser).ConfigureAwait(false);
        }, output, results, failures).ConfigureAwait(false);
        await RunAsync("chromium-print", async () => {
            await NavigateArchiveAsync(browser, archiveUrl).ConfigureAwait(false);
            await browser.Page.EmulateMediaAsync(new PageEmulateMediaOptions { Media = Media.Print }).ConfigureAwait(false);
            return await HtmlPdfComparisonRenderers.CaptureChromiumPageAsync(browser).ConfigureAwait(false);
        }, output, results, failures).ConfigureAwait(false);
        return (browser.Browser?.Version, browser.Page.Url);
    }

    private static async Task NavigateArchiveAsync(HtmlBrowserSession browser, string archiveUrl) {
        await browser.Page.GotoAsync(archiveUrl, new PageGotoOptions {
            WaitUntil = WaitUntilState.DOMContentLoaded,
            Timeout = 30000
        }).ConfigureAwait(false);
        if (!string.Equals(browser.Page.Url, archiveUrl, StringComparison.Ordinal))
            throw new InvalidDataException("Chromium did not load the archived document URL.");
    }

    private static async Task RunAsync(
        string name,
        Func<Task<byte[]>> render,
        string output,
        ICollection<OperationEvidence> results,
        ICollection<string> failures) {
        try {
            var timer = System.Diagnostics.Stopwatch.StartNew();
            byte[] bytes = await render().ConfigureAwait(false);
            timer.Stop();
            string file = name + ".pdf";
            await File.WriteAllBytesAsync(Path.Combine(output, file), bytes).ConfigureAwait(false);
            int pageCount = OfficeIMO.Pdf.PdfDocument.Load(bytes).Inspect().PageCount;
            results.Add(new OperationEvidence(name, file, bytes.Length, Sha256(bytes), pageCount, timer.Elapsed.TotalMilliseconds));
        } catch (Exception exception) {
            failures.Add(FormatFailure(name, exception));
        }
    }

    private static async Task RunConversionAsync(
        string name,
        Func<Task<PdfCore.PdfDocumentConversionResult>> render,
        string output,
        ICollection<OperationEvidence> results,
        ICollection<string> failures) {
        try {
            var timer = Stopwatch.StartNew();
            PdfCore.PdfDocumentConversionResult result = await render().ConfigureAwait(false);
            byte[] bytes = result.ToBytes();
            timer.Stop();
            string file = name + ".pdf";
            await File.WriteAllBytesAsync(Path.Combine(output, file), bytes).ConfigureAwait(false);
            int pageCount = PdfCore.PdfDocument.Load(bytes).Inspect().PageCount;
            var warnings = result.Report.Warnings.Select(warning => new WarningEvidence(
                warning.Converter,
                warning.Code,
                warning.Source,
                warning.Message,
                warning.Severity.ToString(),
                warning.LossKind.ToString(),
                warning.Details)).ToArray();
            results.Add(new OperationEvidence(name, file, bytes.Length, Sha256(bytes), pageCount,
                timer.Elapsed.TotalMilliseconds,
                new ConversionReportEvidence(result.HasLoss, result.Report.FidelityStatus.ToString(), warnings)));
        } catch (Exception exception) {
            failures.Add(FormatFailure(name, exception));
        }
    }

    private static string FormatFailure(string operation, Exception exception) =>
        exception is HtmlDomLimitException limit
            ? operation + ": " + limit.Code + " (" + limit.LimitSource + ", actual=" + limit.Actual + ", limit=" + limit.Limit + ")"
            : operation + ": " + exception;

    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static string? InformationalVersion(System.Reflection.Assembly assembly) =>
        assembly.GetCustomAttributes(typeof(System.Reflection.AssemblyInformationalVersionAttribute), false)
            .OfType<System.Reflection.AssemblyInformationalVersionAttribute>()
            .FirstOrDefault()?.InformationalVersion;

    private static string FindRepositoryRoot() {
        string? current = Directory.GetCurrentDirectory();
        while (current != null) {
            if (File.Exists(Path.Combine(current, "OfficeIMO.sln"))) return current;
            current = Directory.GetParent(current)?.FullName;
        }
        throw new DirectoryNotFoundException("Could not locate the OfficeIMO repository root.");
    }

    private static string GitOutput(string repositoryRoot, params string[] arguments) {
        var start = new ProcessStartInfo("git") {
            WorkingDirectory = repositoryRoot,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        foreach (string argument in arguments) start.ArgumentList.Add(argument);
        using Process process = Process.Start(start) ?? throw new InvalidOperationException("Could not start git.");
        string output = process.StandardOutput.ReadToEnd().Trim();
        string error = process.StandardError.ReadToEnd().Trim();
        process.WaitForExit();
        if (process.ExitCode != 0) throw new InvalidOperationException("git " + string.Join(" ", arguments) + " failed: " + error);
        return output;
    }

    private sealed record OperationEvidence(string Intent, string File, int Bytes, string Sha256, int PageCount,
        double ElapsedMilliseconds, ConversionReportEvidence? ConversionReport = null);
    private sealed record ConversionReportEvidence(bool HasLoss, string FidelityStatus, WarningEvidence[] Warnings);
    private sealed record WarningEvidence(string Converter, string Code, string Source, string Message,
        string Severity, string LossKind, IReadOnlyDictionary<string, string> Details);
}
