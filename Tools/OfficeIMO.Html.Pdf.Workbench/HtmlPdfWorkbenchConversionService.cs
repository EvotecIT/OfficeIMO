using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using HtmlTinkerX;
using OfficeIMO.Drawing;
using OfficeIMO.Html.Pdf.Browser;
using OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf.Workbench;

public sealed class HtmlPdfWorkbenchConversionService {
    public const int MaximumInputCharacters = 2 * 1024 * 1024;
    public const int MaximumPdfBytes = 32 * 1024 * 1024;
    private readonly HtmlBrowserPdfRenderer _browserRenderer;

    public HtmlPdfWorkbenchConversionService(HtmlBrowserPdfRenderer browserRenderer) {
        _browserRenderer = browserRenderer ?? throw new ArgumentNullException(nameof(browserRenderer));
    }

    public async Task<HtmlPdfWorkbenchResult> ConvertAsync(
        HtmlPdfWorkbenchRequest request,
        CancellationToken cancellationToken = default) {
        Validate(request);
        long started = Stopwatch.GetTimestamp();
        ConversionPayload payload = request.Engine == HtmlPdfWorkbenchEngine.Managed
            ? await ConvertManagedAsync(request, cancellationToken).ConfigureAwait(false)
            : await ConvertChromiumAsync(request, cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();

        long elapsedMilliseconds = (long)Stopwatch.GetElapsedTime(started).TotalMilliseconds;
        PdfDocumentInfo pdfInfo = PdfDocument.Open(payload.PdfBytes).Inspect();
        var evidence = new HtmlPdfWorkbenchEvidence(
            "officeimo.html-pdf-workbench/v1",
            DateTimeOffset.UtcNow,
            request.Engine.ToString(),
            payload.RendererVersion,
            Sha256(Encoding.UTF8.GetBytes(request.Html + "\u001e" + request.Css)),
            Sha256(payload.PdfBytes),
            elapsedMilliseconds,
            payload.PdfBytes.Length,
            pdfInfo.PageCount,
            payload.HasLoss,
            request.Settings.Clone(),
            payload.Diagnostics,
            payload.BrowserEvidence);
        byte[] evidenceBytes = JsonSerializer.SerializeToUtf8Bytes(evidence, WorkbenchJsonContext.Default.HtmlPdfWorkbenchEvidence);
        return new HtmlPdfWorkbenchResult(payload.PdfBytes, evidenceBytes, evidence);
    }

    private static async Task<ConversionPayload> ConvertManagedAsync(
        HtmlPdfWorkbenchRequest request,
        CancellationToken cancellationToken) {
        HtmlConversionDocument document = HtmlConversionDocument.Parse(request.Html);
        HtmlPdfSaveOptions options = CreateManagedOptions(request.Settings);
        if (!string.IsNullOrWhiteSpace(request.Css)) options.AdditionalStylesheets.Add(request.Css);
        PdfDocumentConversionResult conversion = await document
            .ToPdfDocumentResultAsync(options, cancellationToken)
            .ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        byte[] pdfBytes = conversion.ToBytes();
        cancellationToken.ThrowIfCancellationRequested();
        if (pdfBytes.Length > MaximumPdfBytes) {
            throw new InvalidOperationException($"Generated PDF exceeds the {MaximumPdfBytes / 1024 / 1024} MiB workbench limit.");
        }
        IReadOnlyList<HtmlPdfWorkbenchDiagnostic> diagnostics = conversion.Warnings
            .Select(warning => new HtmlPdfWorkbenchDiagnostic(
                warning.Severity.ToString(),
                warning.Code,
                warning.Source,
                warning.Message))
            .ToArray();
        return new ConversionPayload(
            pdfBytes,
            GetVersion(typeof(HtmlPdfConverterExtensions).Assembly),
            conversion.HasLoss,
            diagnostics,
            null);
    }

    private async Task<ConversionPayload> ConvertChromiumAsync(
        HtmlPdfWorkbenchRequest request,
        CancellationToken cancellationToken) {
        HtmlPdfWorkbenchSettings settings = request.Settings;
        string margin = settings.MarginMillimeters.ToString("0.###", CultureInfo.InvariantCulture) + "mm";
        var pdfOptions = new HtmlBrowserPdfOptions(
            landscape: settings.Landscape,
            printBackground: settings.PrintBackground,
            format: ResolveBrowserFormat(settings.PageSize),
            marginTop: margin,
            marginRight: margin,
            marginBottom: margin,
            marginLeft: margin,
            preferCssPageSize: settings.HonorCssPageSize,
            outline: settings.Outline,
            tagged: settings.TaggedPdf);
        var browserRequest = new HtmlBrowserPdfRequest(
            HtmlBrowserPdfSource.FromHtml(HtmlPdfPreviewComposer.ComposeForCapture(request.Html, request.Css)),
            pdfOptions,
            readiness: new HtmlBrowserPdfReadiness(
                loadState: HtmlBrowserLoadState.Load,
                stable: true,
                stableMilliseconds: 250,
                timeout: 15000),
            maximumPdfBytes: MaximumPdfBytes);
        HtmlBrowserPdfResult capture = await _browserRenderer.CaptureAsync(browserRequest, cancellationToken).ConfigureAwait(false);
        PdfDocumentConversionResult bridge = capture.ToPdfDocumentResult();
        HtmlBrowserPdfCaptureReport report = bridge.SourceConversionReports.OfType<HtmlBrowserPdfCaptureReport>().Single();
        if (settings.StrictFidelity) report.RequireNoLoss();

        var diagnostics = new List<HtmlPdfWorkbenchDiagnostic>();
        foreach (string warning in capture.Diagnostics.Warnings) {
            diagnostics.Add(new HtmlPdfWorkbenchDiagnostic("Warning", "BrowserCaptureWarning", "chromium", warning));
        }
        foreach (string blocked in capture.Diagnostics.BlockedRequests) {
            diagnostics.Add(new HtmlPdfWorkbenchDiagnostic("Warning", "BrowserRequestBlocked", "network-policy", blocked));
        }
        BrowserCaptureEvidence browserEvidence = new(
            capture.Diagnostics.BrowserVersion,
            capture.Diagnostics.BrowserReused,
            capture.Diagnostics.RetriedAfterBrowserFailure,
            capture.Diagnostics.BlockedRequestCount,
            (long)capture.Diagnostics.QueueDuration.TotalMilliseconds,
            (long)capture.Diagnostics.NavigationDuration.TotalMilliseconds,
            (long)capture.Diagnostics.ReadinessDuration.TotalMilliseconds,
            (long)capture.Diagnostics.PdfDuration.TotalMilliseconds);
        return new ConversionPayload(
            capture.PdfBytes,
            GetVersion(typeof(HtmlBrowserPdfRenderer).Assembly),
            report.HasLoss,
            diagnostics.AsReadOnly(),
            browserEvidence);
    }

    private static HtmlPdfSaveOptions CreateManagedOptions(HtmlPdfWorkbenchSettings settings) {
        OfficePageSize pageSize = settings.PageSize switch {
            "Letter" => OfficePageSizes.Letter,
            "Legal" => OfficePageSizes.Legal,
            "A3" => OfficePageSizes.A3,
            "A5" => OfficePageSizes.A5,
            _ => OfficePageSizes.A4
        };
        if (settings.Landscape) pageSize = pageSize.Landscape();
        double marginPixels = settings.MarginMillimeters / 25.4D * HtmlRenderOptions.CssPixelsPerInch;
        var options = new HtmlPdfSaveOptions {
            PageSize = pageSize,
            Margins = HtmlRenderMargins.All(marginPixels),
            HonorCssPageRules = settings.HonorCssPageSize,
            InteractiveFormControls = settings.InteractiveForms,
            FidelityPolicy = settings.StrictFidelity
                ? HtmlRenderFidelityPolicy.RequireNoLoss
                : HtmlRenderFidelityPolicy.AllowDiagnosedLoss,
            MaxInputCharacters = MaximumInputCharacters
        };
        options.PdfOptions.TaggedStructureMode = settings.TaggedPdf
            ? PdfTaggedStructureMode.CatalogMarkers
            : PdfTaggedStructureMode.None;
        options.PdfOptions.Language = string.IsNullOrWhiteSpace(settings.Language)
            ? null
            : settings.Language.Trim();
        return options;
    }

    private static PdfPageFormat ResolveBrowserFormat(string pageSize) => pageSize switch {
        "Letter" => PdfPageFormat.Letter,
        "Legal" => PdfPageFormat.Legal,
        "A3" => PdfPageFormat.A3,
        "A5" => PdfPageFormat.A5,
        _ => PdfPageFormat.A4
    };

    private static void Validate(HtmlPdfWorkbenchRequest request) {
        ArgumentNullException.ThrowIfNull(request);
        ArgumentNullException.ThrowIfNull(request.Settings);
        if (string.IsNullOrWhiteSpace(request.Html)) throw new ArgumentException("HTML input cannot be empty.", nameof(request));
        if ((long)request.Html.Length + request.Css.Length > MaximumInputCharacters) {
            throw new ArgumentException($"HTML and CSS exceed the {MaximumInputCharacters:N0}-character workbench limit.", nameof(request));
        }
        if (request.Settings.MarginMillimeters is < 0 or > 80
            || double.IsNaN(request.Settings.MarginMillimeters)
            || double.IsInfinity(request.Settings.MarginMillimeters)) {
            throw new ArgumentOutOfRangeException(nameof(request), "Margins must be between 0 and 80 millimeters.");
        }
    }

    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static string GetVersion(Assembly assembly) =>
        assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion
        ?? assembly.GetName().Version?.ToString()
        ?? "unknown";

    private sealed record ConversionPayload(
        byte[] PdfBytes,
        string RendererVersion,
        bool HasLoss,
        IReadOnlyList<HtmlPdfWorkbenchDiagnostic> Diagnostics,
        BrowserCaptureEvidence? BrowserEvidence);
}
