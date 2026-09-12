using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.JSInterop;
using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;

namespace OfficeIMO.Web.Converter.Components;

public partial class PdfWorkbench {
    [Inject] private BrowserDocumentSession Session { get; set; } = null!;
    [Parameter] public int Revision { get; set; }
    private int _sessionRevision = -1;
    private bool _disposed;
    private int _outputGeneration;
    [Inject] private HttpClient Http { get; set; } = null!;
    [Inject] private IJSRuntime JS { get; set; } = null!;
    [Inject] private BrowserPdfToolService PdfTools { get; set; } = null!;

    private ConverterInterop? _interop;
    private PdfToolDefinition ActiveTool { get; set; } = PdfToolCatalog.Default;
    private List<SelectedDocument> Files { get; } = [];
    private PdfToolResult? Result { get; set; }
    private string? ArtifactUrl { get; set; }
    private string? ReportUrl { get; set; }
    private string PageSelection { get; set; } = "all";
    private int PagesPerDocument { get; set; } = 1;
    private int RotationDegrees { get; set; } = 90;
    private PdfOptimizationProfile OptimizationProfile { get; set; } = PdfOptimizationProfile.Balanced;
    private string UserPassword { get; set; } = string.Empty;
    private string OwnerPassword { get; set; } = string.Empty;
    private string RedactionText { get; set; } = string.Empty;
    private bool DestructiveActionConfirmed { get; set; }
    private bool IsBusy { get; set; }
    private List<ConversionDiagnostic> Diagnostics { get; } = [];

    private bool CanRun => !IsBusy && HasRequiredFiles && HasRequiredSettings;
    private bool HasRequiredFiles => ActiveTool.InputMode switch {
        PdfToolInputMode.Single => Files.Count == 1,
        PdfToolInputMode.Pair => Files.Count == 2,
        PdfToolInputMode.Multiple => Files.Count is >= 2 and <= BrowserPdfToolService.MaxPdfFiles,
        _ => false
    };
    private bool HasRequiredSettings =>
        (!ActiveTool.RequiresPageSelection || !string.IsNullOrWhiteSpace(PageSelection)) &&
        (!ActiveTool.RequiresPagesPerDocument || PagesPerDocument > 0) &&
        (!ActiveTool.RequiresUserPassword || !string.IsNullOrWhiteSpace(UserPassword)) &&
        (!ActiveTool.RequiresOwnerPassword || !string.IsNullOrWhiteSpace(OwnerPassword)) &&
        (!ActiveTool.RequiresRedactionText || !string.IsNullOrWhiteSpace(RedactionText)) &&
        (!ActiveTool.RequiresDestructiveConfirmation || DestructiveActionConfirmed);
    private string InputSummary => Files.Count == 0
        ? "Choose PDFs or load the sample to begin."
        : $"{Files.Count} file{(Files.Count == 1 ? string.Empty : "s")} · {ConverterWorkspace.FormatBytes(Files.Sum(static file => file.Size))}";

    protected override void OnInitialized() {
        _interop = new ConverterInterop(JS);
        ActiveTool = PdfToolCatalog.Find(ToolId);
    }

    [Parameter] public string? ToolId { get; set; }

    protected override Task OnParametersSetAsync() {
        if (_disposed) return Task.CompletedTask;
        var tool = PdfToolCatalog.Find(ToolId);
        bool changed = ActiveTool.Id != tool.Id;
        if (!changed && _sessionRevision == Session.Revision) return Task.CompletedTask;
        _sessionRevision = Session.Revision;
        ActiveTool = tool;
        if (changed) { ResetSettings(); Diagnostics.Clear(); }
        Files.Clear();
        var current = Session.Current.Where(file => file.Extension.Equals(".pdf", StringComparison.OrdinalIgnoreCase));
        Files.AddRange(tool.InputMode == PdfToolInputMode.Single ? current.Take(1) : tool.InputMode == PdfToolInputMode.Pair ? current.Take(2) : current);
        return ResetResultAsync();
    }

    private async Task HandleFilesSelectedAsync(InputFileChangeEventArgs args) {
        int revision = Session.Revision;
        Task reset = ResetResultAsync();
        int generation = _outputGeneration;
        bool IsCurrent() => !_disposed && revision == Session.Revision && generation == _outputGeneration;
        await reset;
        if (!IsCurrent()) return;
        Diagnostics.Clear();
        try {
            IReadOnlyList<IBrowserFile> selected = ActiveTool.InputMode == PdfToolInputMode.Single
            ? [args.File]
            : args.GetMultipleFiles(ActiveTool.InputMode == PdfToolInputMode.Pair ? 2 : BrowserPdfToolService.MaxPdfFiles);
            var loaded = new List<SelectedDocument>(selected.Count);
            foreach (IBrowserFile file in selected) {
                string extension = Path.GetExtension(file.Name).ToLowerInvariant();
                if (!string.Equals(extension, ".pdf", StringComparison.OrdinalIgnoreCase)) {
                    throw new InvalidDataException($"{file.Name} is not a PDF file.");
                }
                await using Stream source = file.OpenReadStream(BrowserConversionService.MaxPackageBytes);
                using var buffer = new MemoryStream();
                await source.CopyToAsync(buffer);
                byte[] bytes = buffer.ToArray();
                loaded.Add(new SelectedDocument(file.Name, extension, "PDF", bytes.LongLength, bytes));
            }
            long aggregate = loaded.Sum(static file => file.Size);
            if (aggregate > BrowserPdfToolService.MaxAggregatePdfBytes) {
                throw new InvalidDataException($"Selected PDFs exceed the {ConverterWorkspace.FormatBytes(BrowserPdfToolService.MaxAggregatePdfBytes)} combined limit.");
            }
            if (!IsCurrent()) return;
            Files.Clear();
            Files.AddRange(loaded);
            Session.Open(Files); _sessionRevision = Session.Revision;
            Diagnostics.Add(new ConversionDiagnostic("Ready", $"{Files.Count} PDF file{(Files.Count == 1 ? string.Empty : "s")} loaded in this tab.", "ocx-dot--good"));
        } catch (Exception ex) {
            if (!IsCurrent()) return;
            Diagnostics.Add(new ConversionDiagnostic("Could not load PDFs", DescribeFailure(ex), "ocx-dot--bad"));
        }
    }

    private async Task LoadSampleAsync() {
        int revision = Session.Revision;
        Task reset = ResetResultAsync();
        int generation = _outputGeneration;
        bool IsCurrent() => !_disposed && revision == Session.Revision && generation == _outputGeneration;
        await reset;
        if (!IsCurrent()) return;
        Diagnostics.Clear();
        try {
            byte[] bytes = await Http.GetByteArrayAsync("samples/showcase-dashboard.pdf");
            if (!IsCurrent()) return;
            Files.Clear();
            Files.Add(CreateSample(bytes, ActiveTool.InputMode == PdfToolInputMode.Pair ? "expected" : "showcase"));
            if (ActiveTool.InputMode != PdfToolInputMode.Single) {
                Files.Add(CreateSample((byte[])bytes.Clone(), ActiveTool.InputMode == PdfToolInputMode.Pair ? "actual" : "showcase-copy"));
            }
            if (ActiveTool.Kind == PdfToolKind.Redact) {
                RedactionText = "Critical blockers";
            }
            Session.Open(Files); _sessionRevision = Session.Revision;
            Diagnostics.Add(new ConversionDiagnostic("Sample ready", $"{Files.Count} product PDF file{(Files.Count == 1 ? string.Empty : "s")} loaded locally.", "ocx-dot--good"));
        } catch (Exception ex) {
            if (!IsCurrent()) return;
            Diagnostics.Add(new ConversionDiagnostic("Could not load sample", DescribeFailure(ex), "ocx-dot--bad"));
        }
    }

    private static SelectedDocument CreateSample(byte[] bytes, string suffix) =>
        new($"officeimo-{suffix}.pdf", ".pdf", "PDF", bytes.LongLength, bytes);

    private Task RemoveFileAsync(int index) {
        if (index < 0 || index >= Files.Count) return Task.CompletedTask;
        Files.RemoveAt(index);
        Session.SelectCurrent(Files); _sessionRevision = Session.Revision;
        Diagnostics.Clear();
        return ResetResultAsync();
    }

    private Task MoveFileAsync(PdfFileMoveRequest request) {
        int target = request.Index + request.Offset;
        if (request.Index < 0 || request.Index >= Files.Count || target < 0 || target >= Files.Count) return Task.CompletedTask;
        SelectedDocument file = Files[request.Index];
        Files.RemoveAt(request.Index);
        Files.Insert(target, file);
        Session.SelectCurrent(Files); _sessionRevision = Session.Revision;
        return ResetResultAsync();
    }

    private Task ClearFilesAsync() {
        Files.Clear();
        Session.SelectCurrent(Files); _sessionRevision = Session.Revision;
        Diagnostics.Clear();
        return ResetResultAsync();
    }

    private async Task RunAsync() {
        if (!CanRun || _interop is null) return;
        int sourceRevision = Session.Revision;
        string sourceTool = ActiveTool.Id;
        Session.ClearResult();
        IsBusy = true;
        Task reset = ResetResultAsync();
        int generation = _outputGeneration;
        bool IsCurrent() => !_disposed && generation == _outputGeneration && sourceRevision == Session.Revision && sourceTool == ActiveTool.Id;
        await using var urls = new ConverterObjectUrlBatch(_interop, IsCurrent);
        try {
            await reset;
            if (!IsCurrent()) return;
            Diagnostics.Clear();
            await InvokeAsync(StateHasChanged);
            await Task.Yield();
            if (!IsCurrent()) return;
            var result = PdfTools.Execute(new PdfToolRequest(
                ActiveTool,
                Files.ToArray(),
                PageSelection,
                PagesPerDocument,
                RotationDegrees,
                OptimizationProfile,
                UserPassword,
                OwnerPassword,
                RedactionText,
                DestructiveActionConfirmed));
            string artifactUrl = await urls.CreateAsync(result.Artifact.Bytes, result.Artifact.ContentType);
            string? reportUrl = null;
            if (result.Report is not null) {
                reportUrl = await urls.CreateAsync(result.Report.Bytes, result.Report.ContentType);
            }
            urls.Commit();
            Result = result;
            ArtifactUrl = artifactUrl;
            ReportUrl = reportUrl;
            Diagnostics.Add(new ConversionDiagnostic("Operation complete", Result.Summary, "ocx-dot--good"));
            if (!_disposed && sourceTool == ActiveTool.Id) Session.SetResult(Result.Artifact.Bytes, Result.Artifact.FileName, sourceRevision);
        } catch (Exception ex) {
            if (!IsCurrent()) return;
            Result = null;
            Diagnostics.Add(new ConversionDiagnostic("PDF operation failed", DescribeFailure(ex), "ocx-dot--bad"));
        } finally {
            UserPassword = string.Empty;
            OwnerPassword = string.Empty;
            IsBusy = false;
        }
    }

    private void ResetSettings() {
        PageSelection = "all";
        PagesPerDocument = 1;
        RotationDegrees = 90;
        OptimizationProfile = PdfOptimizationProfile.Balanced;
        UserPassword = string.Empty;
        OwnerPassword = string.Empty;
        RedactionText = string.Empty;
        DestructiveActionConfirmed = false;
    }

    private async Task HandleSettingsChangedAsync() {
        Diagnostics.Clear();
        await ResetResultAsync();
    }

    private async Task ResetResultAsync() {
        _outputGeneration++;
        Session.ClearResult();
        string? artifactUrl = ArtifactUrl, reportUrl = ReportUrl;
        ArtifactUrl = null;
        ReportUrl = null;
        Result = null;
        if (_interop is not null) {
            await _interop.RevokeObjectUrlAsync(artifactUrl);
            await _interop.RevokeObjectUrlAsync(reportUrl);
        }
    }


    private static string DescribeFailure(Exception ex) => ex switch {
        IOException when ex is not InvalidDataException => "The browser workbench accepts PDFs up to 25 MB each.",
        _ => ex.Message
    };

    public async ValueTask DisposeAsync() {
        _disposed = true;
        if (_interop is null) return;
        await _interop.RevokeObjectUrlAsync(ArtifactUrl);
        await _interop.RevokeObjectUrlAsync(ReportUrl);
        await _interop.DisposeAsync();
    }
}
