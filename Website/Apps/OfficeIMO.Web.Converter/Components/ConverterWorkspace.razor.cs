using System.Diagnostics;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.JSInterop;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;

namespace OfficeIMO.Web.Converter.Components;

public partial class ConverterWorkspace {
    [Inject] private BrowserDocumentSession Session { get; set; } = null!;
    [Parameter] public int Revision { get; set; }
    private int _sessionRevision = -1;
    private bool _disposed;
    internal const long MaxUploadBytes = BrowserConversionService.MaxPackageBytes;

    private const string DefaultMarkdown = """
# OfficeIMO conversion sample

This **Markdown** becomes a browser preview or an editable Word document.

| Route | Execution |
| --- | --- |
| Markdown to HTML | Browser-local |
| Markdown to DOCX | Browser-local |

- No account
- No server upload
- Downloadable output
""";

    private const string DefaultHtml = """
<article>
  <h1>OfficeIMO HTML sample</h1>
  <p>A short report with <strong>headings and a list</strong>.</p>
  <ul><li>Headings</li><li>Lists</li><li>Links</li></ul>
</article>
""";

    [Inject] private HttpClient Http { get; set; } = null!;
    [Inject] private IJSRuntime JS { get; set; } = null!;
    [Inject] private BrowserConversionService ConversionService { get; set; } = null!;

    private ConverterInterop? _interop;
    private DotNetObjectReference<ConverterWorkspace>? _webMcpReference;
    private ConversionRoute ActiveRoute { get; set; } = ConversionRouteCatalog.Default;
    private SelectedDocument? SelectedFile { get; set; }
    private ConversionResult? Output { get; set; }
    private string? OutputUrl { get; set; }
    private string? OutputReportUrl { get; set; }
    private string? OutputOverlayUrl { get; set; }
    private string? OutputSupportUrl { get; set; }
    private string OutputFileName { get; set; } = "officeimo-output";
    private string? OutputReportFileName { get; set; }
    private string? OutputOverlayFileName { get; set; }
    private string? OutputSupportFileName { get; set; }
    private string TextInput { get; set; } = DefaultMarkdown;
    private bool LimitExcelRows { get; set; }
    private bool GenerateDebugOverlay { get; set; }
    private bool IncludeDocumentContentInSupportBundle { get; set; }
    private string SelectedProfileId { get; set; } = BrowserPdfProfileCatalog.Faithful.Id;
    private string SelectedPowerPointImportProfileId { get; set; } = BrowserPowerPointImportProfileCatalog.Editable.Id;
    private bool PreviewOutput { get; set; } = true;
    private bool IsBusy { get; set; }
    private long ElapsedMilliseconds { get; set; }
    private List<ConversionDiagnostic> Diagnostics { get; } = [];

    private static IReadOnlyList<ConversionRoute> Routes => ConversionRouteCatalog.All;
    private static IReadOnlyList<BrowserPdfProfile> PdfProfiles => BrowserPdfProfileCatalog.All;
    private BrowserPdfProfile SelectedProfile => BrowserPdfProfileCatalog.Find(SelectedProfileId);
    private static IReadOnlyList<BrowserPowerPointImportProfile> PowerPointImportProfiles => BrowserPowerPointImportProfileCatalog.All;
    private BrowserPowerPointImportProfile SelectedPowerPointImportProfile => BrowserPowerPointImportProfileCatalog.Find(SelectedPowerPointImportProfileId);
    private bool IsPdfRoute => string.Equals(ActiveRoute.Target, "PDF", StringComparison.OrdinalIgnoreCase);
    private bool IsPowerPointImportRoute => string.Equals(ActiveRoute.Id, "pdf-pptx", StringComparison.Ordinal);
    private IReadOnlyList<ConversionWarningView> ReviewWarnings =>
        Output?.StructuredWarnings
            .Where(static warning =>
                !string.Equals(warning.Severity, "Information", StringComparison.OrdinalIgnoreCase))
            .ToArray()
        ?? [];
    private IEnumerable<IGrouping<string, ConversionWarningView>> WarningGroups =>
        GroupWarnings(ReviewWarnings);
    private IEnumerable<IGrouping<string, ConversionWarningView>> InformationGroups =>
        GroupWarnings(
            Output?.StructuredWarnings
                .Where(static warning =>
                    string.Equals(warning.Severity, "Information", StringComparison.OrdinalIgnoreCase))
            ?? Enumerable.Empty<ConversionWarningView>());
    private bool CanConvert => !IsBusy && (ActiveRoute.InputKind == ConversionInputKind.File ? SelectedFile is not null : !string.IsNullOrWhiteSpace(TextInput));
    private string FileInputId => $"conversion-file-input-{ActiveRoute.Id}";
    private string OutputHeading => Output?.FileName ?? $"{ActiveRoute.Target} output";
    private string ElapsedLabel => ElapsedMilliseconds < 1000 ? $"{ElapsedMilliseconds} ms" : $"{ElapsedMilliseconds / 1000d:0.0} s";

    protected override void OnInitialized() {
        _interop = new ConverterInterop(JS);
        ActiveRoute = ConversionRouteCatalog.Find(RouteId);
        TextInput = IsHtmlInputRoute(ActiveRoute) ? DefaultHtml : DefaultMarkdown;
    }

    [Parameter] public string? RouteId { get; set; }

    protected override async Task OnParametersSetAsync() {
        var route = ConversionRouteCatalog.Find(RouteId);
        bool changed = ActiveRoute.Id != route.Id;
        bool hadWorkingFile = SelectedFile is not null;
        await SelectRouteAsync(route);
        if (changed || _sessionRevision != Session.Revision) {
            _sessionRevision = Session.Revision;
            await ResetOutputAsync();
            SelectedFile = Session.Current.FirstOrDefault(file => route.Accept.Split(',').Contains(file.Extension, StringComparer.OrdinalIgnoreCase));
            if (route.InputKind == ConversionInputKind.Text && SelectedFile is null && (Session.Current.Count > 0 || hadWorkingFile)) TextInput = string.Empty;
            if (route.InputKind == ConversionInputKind.Text && SelectedFile is not null) {
                string text = System.Text.Encoding.UTF8.GetString(SelectedFile.Bytes);
                if (text.Length <= BrowserConversionService.MaxTextInputChars) TextInput = text;
                else { TextInput = string.Empty; Diagnostics.Add(new("Text too large", "The working file exceeds this text tool's limit.", "ocx-dot--warn")); }
            }
        }
    }

    protected override async Task OnAfterRenderAsync(bool firstRender) {
        if (!firstRender || _interop is null) {
            return;
        }

        _webMcpReference = DotNetObjectReference.Create(this);
        await _interop.RegisterWebMcpToolAsync(_webMcpReference);
    }

    private async Task SelectRouteAsync(ConversionRoute route) {
        if (ActiveRoute.Id == route.Id) {
            return;
        }
        await ResetOutputAsync();
        ActiveRoute = route;
        SelectedFile = null;
        TextInput = IsHtmlInputRoute(route) ? DefaultHtml : DefaultMarkdown;
        GenerateDebugOverlay = false;
        IncludeDocumentContentInSupportBundle = false;
        Diagnostics.Clear();
    }

    private async Task HandleFileSelectedAsync(InputFileChangeEventArgs args) {
        int revision = Session.Revision;
        string routeId = ActiveRoute.Id;
        await ResetOutputAsync();
        Diagnostics.Clear();
        IBrowserFile file = args.File;
        string extension = Path.GetExtension(file.Name).ToLowerInvariant();
        if (!ActiveRoute.Accept.Split(',').Contains(extension, StringComparer.OrdinalIgnoreCase)) {
            SelectedFile = null;
            Diagnostics.Add(new("Unsupported file", $"Choose a {ActiveRoute.Source} file for this route.", "ocx-dot--bad"));
            return;
        }

        try {
            await using Stream source = file.OpenReadStream(MaxUploadBytes);
            using var buffer = new MemoryStream();
            await source.CopyToAsync(buffer);
            if (_disposed || revision != Session.Revision || routeId != ActiveRoute.Id) return;
            SelectedFile = new(file.Name, extension, ActiveRoute.Source, file.Size, buffer.ToArray());
            Session.Open([SelectedFile]); _sessionRevision = Session.Revision;
            Diagnostics.Add(new("Ready", $"{file.Name} is loaded in this browser tab.", "ocx-dot--good"));
        } catch (IOException) {
            SelectedFile = null;
            Diagnostics.Add(new("File too large", $"The browser demo accepts files up to {FormatBytes(MaxUploadBytes)}.", "ocx-dot--bad"));
        } catch (Exception ex) {
            SelectedFile = null;
            Diagnostics.Add(new("Could not read file", DescribeFailure(ex), "ocx-dot--bad"));
        }
    }

    private async Task LoadSampleAsync() {
        int revision = Session.Revision;
        string routeId = ActiveRoute.Id;
        SampleDocument sample = ActiveRoute.Id switch {
            "pdf-docx" or "pdf-xlsx" or "pdf-pptx" or "pdf-html" or "pdf-png" => new("Sample PDF", "samples/showcase-dashboard.pdf", "OfficeIMO-Showcase.pdf", ".pdf"),
            "xlsx-pdf" => new("Sample XLSX", "samples/basic.xlsx", "OfficeIMO-Table.xlsx", ".xlsx"),
            "pptx-pdf" => new("Sample PPTX", "samples/conversion-proof.pptx", "OfficeIMO-Conversion-Proof.pptx", ".pptx"),
            _ => new("Sample DOCX", "samples/business-summary.docx", "OfficeIMO-Monthly-Operations.docx", ".docx")
        };
        await ResetOutputAsync();
        Diagnostics.Clear();
        try {
            byte[] bytes = await Http.GetByteArrayAsync(sample.Path);
            if (_disposed || revision != Session.Revision || routeId != ActiveRoute.Id) return;
            SelectedFile = new(sample.FileName, sample.Extension, ActiveRoute.Source, bytes.LongLength, bytes);
            Session.Open([SelectedFile]); _sessionRevision = Session.Revision;
            Diagnostics.Add(new("Sample ready", $"{sample.FileName} is loaded locally.", "ocx-dot--good"));
        } catch (Exception ex) {
            Diagnostics.Add(new("Could not load sample", DescribeFailure(ex), "ocx-dot--bad"));
        }
    }

    private async Task LoadTextSampleAsync() {
        await ResetOutputAsync();
        TextInput = IsHtmlInputRoute(ActiveRoute) ? DefaultHtml : DefaultMarkdown;
        Diagnostics.Clear();
        Diagnostics.Add(new("Sample ready", $"Sample {ActiveRoute.Source} is ready.", "ocx-dot--good"));
    }

    private async Task ConvertAsync() {
        if (!CanConvert || _interop is null) {
            return;
        }

        int sourceRevision = Session.Revision;
        string sourceRoute = ActiveRoute.Id;
        Session.ClearResult();
        IsBusy = true;
        await ResetOutputAsync();
        Diagnostics.Clear();
        await InvokeAsync(StateHasChanged);
        await Task.Yield();
        var stopwatch = Stopwatch.StartNew();

        try {
            if (_disposed || sourceRevision != Session.Revision || sourceRoute != ActiveRoute.Id) return;
            Output = ActiveRoute.InputKind == ConversionInputKind.File
                ? ConversionService.ConvertFile(
                    ActiveRoute,
                    SelectedFile!,
                    LimitExcelRows,
                    SelectedProfile,
                    GenerateDebugOverlay,
                    SelectedPowerPointImportProfile.Mode)
                : ConversionService.ConvertText(
                    ActiveRoute,
                    TextInput,
                    SelectedProfile,
                    GenerateDebugOverlay);
            stopwatch.Stop();
            ElapsedMilliseconds = stopwatch.ElapsedMilliseconds;
            OutputFileName = Output.FileName;
            OutputUrl = await _interop.CreateObjectUrlAsync(Output.Bytes, Output.ContentType);
            if (Output.CompanionReport is not null) {
                OutputReportFileName = Output.CompanionReport.FileName;
                OutputReportUrl = await _interop.CreateObjectUrlAsync(
                    Output.CompanionReport.Bytes,
                    Output.CompanionReport.ContentType);
            }
            if (Output.DebugOverlay is not null) {
                OutputOverlayFileName = Output.DebugOverlay.FileName;
                OutputOverlayUrl = await _interop.CreateObjectUrlAsync(
                    Output.DebugOverlay.Bytes,
                    Output.DebugOverlay.ContentType);
            }
            string fidelity = Output.FidelityStatus ?? "Complete";
            if (!_disposed && sourceRoute == ActiveRoute.Id) Session.SetResult(Output.Bytes, Output.FileName, sourceRevision);
            string tone = fidelity is "Complete" or "Reconstructed" ? "ocx-dot--good" : "ocx-dot--warn";
            Diagnostics.Add(new($"{fidelity} conversion", $"Created {Output.FileName} locally in {ElapsedLabel}. {Output.ProvenanceSummary}", tone));
        } catch (Exception ex) {
            stopwatch.Stop();
            ElapsedMilliseconds = stopwatch.ElapsedMilliseconds;
            Output = null;
            Diagnostics.Add(new("Conversion failed", DescribeFailure(ex), "ocx-dot--bad"));
        } finally {
            IsBusy = false;
        }
    }

    [JSInvokable]
    public async Task<WebMcpConversionResult> ConvertSelectedDocumentForWebMcpAsync() {
        if (ActiveRoute.InputKind != ConversionInputKind.File) {
            return new(false, ActiveRoute.Id, ActiveRoute.Source, ActiveRoute.Target, null, 0, 0, 0,
                "This tool converts a file already selected in the visible workspace; choose a file-based route first.");
        }
        if (SelectedFile is null) {
            Diagnostics.Clear();
            Diagnostics.Add(new("Select a document", "Choose or load a sample document in the visible workspace before using the Website Tool.", "ocx-dot--warn"));
            await InvokeAsync(StateHasChanged);
            return new(false, ActiveRoute.Id, ActiveRoute.Source, ActiveRoute.Target, null, 0, 0, 0,
                "No document is selected. Choose or load a sample document in the visible workspace, then try again.");
        }
        if (IsBusy) {
            return new(false, ActiveRoute.Id, ActiveRoute.Source, ActiveRoute.Target, null, 0, 0, 0,
                "A conversion is already running in this browser tab.");
        }

        await ConvertAsync();
        await InvokeAsync(StateHasChanged);
        if (Output is null) {
            return new(false, ActiveRoute.Id, ActiveRoute.Source, ActiveRoute.Target, null, 0, 0, ElapsedMilliseconds,
                BoundWebMcpText(Diagnostics.LastOrDefault()?.Message ?? "The browser-local conversion did not produce an output.", 300));
        }

        return new(
            true,
            ActiveRoute.Id,
            ActiveRoute.Source,
            ActiveRoute.Target,
            BoundWebMcpText(Output.FileName, 180),
            Output.Bytes.LongLength,
            ReviewWarnings.Count,
            ElapsedMilliseconds,
            "Conversion completed locally. Review the visible preview, warnings, and download action before saving the result.");
    }

    private static string BoundWebMcpText(string? value, int maximumCharacters) {
        string text = value?.Trim() ?? string.Empty;
        if (text.Length <= maximumCharacters) {
            return text;
        }

        int length = maximumCharacters;
        if (length > 0 &&
            char.IsHighSurrogate(text[length - 1]) &&
            length < text.Length &&
            char.IsLowSurrogate(text[length])) {
            length--;
        }
        return text[..length];
    }

    private static bool IsHtmlInputRoute(ConversionRoute route) =>
        route.Id is "html-markdown" or "html-pdf";

    private async Task HandlePowerPointImportProfileChangedAsync() {
        await ResetOutputAsync();
        Diagnostics.Clear();
        if (SelectedFile is not null) {
            Diagnostics.Add(new(
                "Mode changed",
                $"{SelectedPowerPointImportProfile.Label} is selected. Convert again to create a matching PPTX and report.",
                "ocx-dot--good"));
        }
    }

    private async Task PrepareSupportBundleAsync() {
        if (_interop is null || Output is null) {
            return;
        }

        if (!string.IsNullOrWhiteSpace(OutputSupportUrl)) {
            await _interop.RevokeObjectUrlAsync(OutputSupportUrl);
        }
        BrowserConversionArtifact supportBundle = ConversionService.CreateSupportBundle(
            Output,
            IncludeDocumentContentInSupportBundle);
        OutputSupportFileName = supportBundle.FileName;
        OutputSupportUrl = await _interop.CreateObjectUrlAsync(
            supportBundle.Bytes,
            supportBundle.ContentType);
        Diagnostics.Add(new(
            "Support bundle ready",
            IncludeDocumentContentInSupportBundle
                ? "The bundle includes source and PDF bytes because you opted in."
                : "The bundle contains fingerprints and diagnostics only; document content is excluded.",
            "ocx-dot--good"));
    }

    private async Task InvalidateSupportBundleAsync() {
        if (_interop is not null && !string.IsNullOrWhiteSpace(OutputSupportUrl)) {
            await _interop.RevokeObjectUrlAsync(OutputSupportUrl);
        }
        OutputSupportUrl = null;
        OutputSupportFileName = null;
    }

    private async Task ResetOutputAsync() {
        Session.ClearResult();
        if (_interop is not null && !string.IsNullOrWhiteSpace(OutputUrl)) {
            await _interop.RevokeObjectUrlAsync(OutputUrl);
        }
        if (_interop is not null && !string.IsNullOrWhiteSpace(OutputReportUrl)) {
            await _interop.RevokeObjectUrlAsync(OutputReportUrl);
        }
        if (_interop is not null && !string.IsNullOrWhiteSpace(OutputOverlayUrl)) {
            await _interop.RevokeObjectUrlAsync(OutputOverlayUrl);
        }
        if (_interop is not null && !string.IsNullOrWhiteSpace(OutputSupportUrl)) {
            await _interop.RevokeObjectUrlAsync(OutputSupportUrl);
        }
        OutputUrl = null;
        OutputReportUrl = null;
        OutputOverlayUrl = null;
        OutputSupportUrl = null;
        OutputReportFileName = null;
        OutputOverlayFileName = null;
        OutputSupportFileName = null;
        Output = null;
        ElapsedMilliseconds = 0;
    }


    internal static string FormatBytes(long bytes) {
        string[] units = ["B", "KB", "MB", "GB"];
        double value = bytes;
        int unit = 0;
        while (value >= 1024 && unit < units.Length - 1) { value /= 1024; unit++; }
        return unit == 0 ? $"{bytes} B" : $"{value:0.##} {units[unit]}";
    }

    private static string DescribeFailure(Exception ex) =>
        ex.GetType().Name.Contains("PdfTextEncodingPreflightException", StringComparison.Ordinal)
            ? "This document uses text that needs an embedded font not available to the browser conversion. " + ex.Message
            : ex.Message;

    private static IEnumerable<IGrouping<string, ConversionWarningView>> GroupWarnings(
        IEnumerable<ConversionWarningView> warnings) =>
        warnings
            .GroupBy(
                static warning => warning.PageNumber.HasValue
                    ? $"Page {warning.PageNumber.Value} · {warning.Construct}"
                    : $"Document · {warning.Construct}",
                StringComparer.Ordinal)
            .OrderBy(static group => group.Key, StringComparer.Ordinal);

    private static string WarningTitle(ConversionWarningView warning) =>
        string.Equals(warning.Construct, warning.Code, StringComparison.OrdinalIgnoreCase)
            ? warning.Code
            : $"{warning.Construct} · {warning.Code}";

    public async ValueTask DisposeAsync() {
        _disposed = true;
        if (_interop is not null) {
            await _interop.UnregisterWebMcpToolAsync();
            await _interop.RevokeObjectUrlAsync(OutputUrl);
            await _interop.RevokeObjectUrlAsync(OutputReportUrl);
            await _interop.RevokeObjectUrlAsync(OutputOverlayUrl);
            await _interop.RevokeObjectUrlAsync(OutputSupportUrl);
            await _interop.DisposeAsync();
        }
        _webMcpReference?.Dispose();
    }
}
