using System.Diagnostics;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Mhtml;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Replays one frozen MHTML archive through one declared static PDF intent.
/// MHTML parsing precedes the conversion measurement used by the H10 budget owner.
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
        if (!TryParse(args, out string? isolatedIntent, out int? maxCssRules, out bool requireClean, out string userAgent)) {
            Console.Error.WriteLine("html-mhtml-evidence --mhtml <existing-archive> --output <new-directory> --isolated-officeimo-intent print|print-zero-margin|screen-media|screen-snapshot [--require-clean-source] [--max-css-rules 10000..20000] [--user-agent document|browser]");
            return 2;
        }
        string repositoryRoot = FindRepositoryRoot();
        GitSourceState sourceState = await SourceProvenanceReader.ReadGitStateAsync(repositoryRoot).ConfigureAwait(false)
            ?? throw new InvalidOperationException("Could not read source provenance for the H10 worker.");
        string sourceCommit = sourceState.Commit;
        bool worktreeDirty = !sourceState.IsClean;
        var ownerVersions = new Dictionary<string, string?> {
            ["Core"] = InformationalVersion(typeof(OfficeFontFaceCollection).Assembly),
            ["Email"] = InformationalVersion(typeof(OfficeIMO.Email.EmailDocumentReader).Assembly),
            ["HTML Core"] = InformationalVersion(typeof(OfficeIMO.Html.Dom.HtmlDocument).Assembly),
            ["HTML AngleSharp"] = InformationalVersion(typeof(OfficeIMO.Html.Providers.AngleSharpHtmlParser).Assembly),
            ["HTML"] = InformationalVersion(typeof(HtmlRenderEngine).Assembly),
            ["MHTML"] = InformationalVersion(typeof(MhtmlDocument).Assembly),
            ["MHTML PDF"] = InformationalVersion(typeof(MhtmlPdfConverterExtensions).Assembly),
            ["PDF"] = InformationalVersion(typeof(PdfCore.PdfDocument).Assembly),
            ["evidence runner"] = InformationalVersion(typeof(HtmlMhtmlEvidenceRunner).Assembly)
        };
        if (requireClean && (worktreeDirty || ownerVersions.Values.Any(version => version == null
            || !version.EndsWith("+" + sourceCommit, StringComparison.OrdinalIgnoreCase)))) {
            throw new InvalidOperationException("Evidence requires a clean checkout and all owner assemblies rebuilt from its exact HEAD commit.");
        }
        string input = Path.GetFullPath(args[2]);
        long inputLength = new FileInfo(input).Length;
        if (inputLength == 0 || inputLength > MaximumArchiveBytes)
            throw new InvalidDataException("MHTML input is empty or exceeds the 32 MiB input limit.");
        byte[] archive = await File.ReadAllBytesAsync(input).ConfigureAwait(false);
        if (archive.Length == 0 || archive.Length > MaximumArchiveBytes)
            throw new InvalidDataException("MHTML input is empty or exceeds the 32 MiB input limit.");
        string output = Path.GetFullPath(args[4]);
        if (Directory.Exists(output) || File.Exists(output)) throw new IOException("Output directory already exists.");
        Directory.CreateDirectory(output);
        await File.WriteAllBytesAsync(Path.Combine(output, "source.mhtml"), archive).ConfigureAwait(false);
        var results = new List<OperationEvidence>();
        var failures = new List<string>();
        MhtmlDocument? document = null;
        try {
            using var source = new MemoryStream(archive, writable: false);
            HtmlConversionDocumentOptions? htmlOptions = null;
            if (maxCssRules.HasValue) {
                htmlOptions = HtmlConversionDocumentOptions.CreateUntrustedProfile();
                HtmlConversionLimits limits = htmlOptions.Limits.Clone();
                limits.MaxCssRules = maxCssRules.Value;
                htmlOptions.Limits = limits;
            }
            document = MhtmlDocument.Load(source, htmlOptions: htmlOptions);
        } catch (Exception exception) {
            failures.Add("OfficeIMO MHTML load: " + exception);
        }
        // Archive parsing is deliberately outside RunConversionAsync's timer and
        // allocation counter; changing this boundary invalidates old ceilings.
        if (document != null) {
            switch (isolatedIntent) {
                case "print":
                    await RunConversionAsync("officeimo-print", () => document.ToPdfDocumentResultAsync(userAgent == "document" ? null : CreatePdfOptions(userAgent)), output, results, failures).ConfigureAwait(false);
                    break;
                case "print-zero-margin":
                    await RunConversionAsync("officeimo-print-zero-margin", () => {
                        HtmlToPdfOptions options = CreatePdfOptions(userAgent);
                        options.Margins = HtmlRenderMargins.All(0);
                        return document.ToPdfDocumentResultAsync(options);
                    }, output, results, failures).ConfigureAwait(false);
                    break;
                case "screen-media":
                    await RunConversionAsync("officeimo-screen-media", () => RenderScreenPdfAsync(document, HtmlRenderIntentProfile.ScreenMediaPaged, userAgent), output, results, failures).ConfigureAwait(false);
                    break;
                case "screen-snapshot":
                    await RunConversionAsync("officeimo-screen-snapshot", () => RenderScreenPdfAsync(document, HtmlRenderIntentProfile.ScreenSnapshotPaged, userAgent), output, results, failures).ConfigureAwait(false);
                    break;
            }
        }
        HtmlToPdfOptions selectedDefaults = CreatePdfOptions(userAgent);
        var report = new {
            schemaVersion = 3,
            sourceUrl = document?.BaseUri.ToString(),
            runUtc = DateTimeOffset.UtcNow,
            sourceMode = "frozen-mhtml-replay",
            readinessPolicy = "preloaded frozen MIME resources; no browser capture or live resource readiness",
            browserReference = "not-recorded",
            sourceCommit,
            worktreeDirty,
            maxCssRules = maxCssRules ?? HtmlConversionLimits.CreateUntrustedProfile().MaxCssRules,
            isolatedOfficeimoIntent = isolatedIntent,
            userAgent,
            userAgentStyles = selectedDefaults.UserAgentStyles.ToString(),
            defaultFontFamily = selectedDefaults.DefaultFontFamily,
            maxSelectorEvaluations = HtmlConversionLimits.CreateUntrustedProfile().MaxSelectorEvaluations,
            ownerVersions,
            ownerBoundary = "HTML PDF option and render APIs are owned by OfficeIMO.Html; no separate forwarder assembly is attributed.",
            conversionScope = "layout and PDF conversion of preloaded MHTML; PDF serialization included, archive load and saved-artifact inspection excluded",
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

    private static bool TryParse(string[] args, out string? intent, out int? maxCssRules, out bool requireClean, out string userAgent) {
        intent = null;
        maxCssRules = null;
        requireClean = false;
        userAgent = "document";
        if (args.Length < 7 || args[1] != "--mhtml" || args[3] != "--output") return false;
        var flags = new HashSet<string>(StringComparer.Ordinal);
        for (int index = 5; index < args.Length; index++) {
            string flag = args[index];
            if (!flags.Add(flag)) return false;
            if (flag == "--require-clean-source") {
                requireClean = true;
            } else if (flag == "--isolated-officeimo-intent" && index + 1 < args.Length
                && args[index + 1] is "print" or "print-zero-margin" or "screen-media" or "screen-snapshot") {
                intent = args[++index];
            } else if (flag == "--max-css-rules" && index + 1 < args.Length
                && int.TryParse(args[++index], out int limit) && limit >= 10_000 && limit <= 20_000) {
                maxCssRules = limit;
            } else if (flag == "--user-agent" && index + 1 < args.Length
                && args[index + 1] is "document" or "browser") {
                userAgent = args[++index];
            } else {
                return false;
            }
        }
        return intent != null;
    }

    private static HtmlToPdfOptions CreatePdfOptions(string userAgent) {
        var options = new HtmlToPdfOptions();
        if (userAgent == "browser") options.UseBrowserUserAgentStyles();
        return options;
    }

    private static Task<PdfCore.PdfDocumentConversionResult> RenderScreenPdfAsync(MhtmlDocument document, HtmlRenderIntentProfile profile, string userAgent, bool allowDocumentFontEmbedding = false) {
        HtmlToPdfOptions options = CreatePdfOptions(userAgent);
        options.ViewportWidth = ViewportWidth;
        options.ViewportHeight = ViewportHeight;
        options.Margins = HtmlRenderMargins.All(0);
        options.HonorCssPageRules = false;
        options.ResourcePolicy.AllowDocumentFontEmbedding = allowDocumentFontEmbedding;
        if (profile == HtmlRenderIntentProfile.ScreenSnapshotPaged) {
            options.PageSize = new OfficePageSize(
                global::OfficeIMO.OfficeLength.Inches(ViewportWidth / HtmlRenderOptions.CssPixelsPerInch),
                global::OfficeIMO.OfficeLength.Inches(ViewportHeight / HtmlRenderOptions.CssPixelsPerInch));
        }
        HtmlRenderRequest request = HtmlRenderRequest.Create(profile, HtmlRenderEncoder.Pdf, options);
        return document.RenderToPdfDocumentResultAsync(request);
    }

    private static async Task RunConversionAsync(
        string name,
        Func<Task<PdfCore.PdfDocumentConversionResult>> render,
        string output,
        ICollection<OperationEvidence> results,
        ICollection<string> failures) {
        try {
            var timer = Stopwatch.StartNew();
            long allocationStart = GC.GetTotalAllocatedBytes(precise: true);
            PdfCore.PdfDocumentConversionResult result = await render().ConfigureAwait(false);
            byte[] bytes = result.ToBytes();
            timer.Stop();
            long allocatedBytes = GC.GetTotalAllocatedBytes(precise: true) - allocationStart;
            string file = name + ".pdf";
            await File.WriteAllBytesAsync(Path.Combine(output, file), bytes).ConfigureAwait(false);
            using var pdfInput = new MemoryStream(bytes, writable: false);
            using PdfCore.PdfDocument pdfDocument = PdfCore.PdfDocument.Load(pdfInput);
            PdfCore.PdfDocumentInfo inspection = pdfDocument.Inspect();
            var warnings = result.Report.Warnings.Select(warning => new WarningEvidence(
                warning.Converter,
                warning.Code,
                warning.Source,
                warning.Message,
                warning.Severity.ToString(),
                warning.LossKind.ToString(),
                warning.Details)).ToArray();
            results.Add(new OperationEvidence(name, file, bytes.Length, Sha256(bytes), inspection.PageCount,
                inspection.Pages.Select(page => new PageSizeEvidence(page.PageNumber, page.Width, page.Height)).ToArray(),
                timer.Elapsed.TotalMilliseconds,
                new ConversionReportEvidence(result.HasLoss, result.Report.FidelityStatus.ToString(), warnings), allocatedBytes));
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

    private sealed record OperationEvidence(string Intent, string File, int Bytes, string Sha256, int PageCount,
        PageSizeEvidence[] PageSizesPoints, double ElapsedMilliseconds, ConversionReportEvidence? ConversionReport = null,
        long? ManagedAllocatedBytes = null);
    private sealed record PageSizeEvidence(int PageNumber, double Width, double Height);
    private sealed record ConversionReportEvidence(bool HasLoss, string FidelityStatus, WarningEvidence[] Warnings);
    private sealed record WarningEvidence(string Converter, string Code, string Source, string Message,
        string Severity, string LossKind, IReadOnlyDictionary<string, string> Details);
}
