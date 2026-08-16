using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.JSInterop;
using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;

namespace OfficeIMO.Web.Converter.Components;

public partial class PdfWorkbench {
    [Inject] private HttpClient Http { get; set; } = null!;
    [Inject] private IJSRuntime JS { get; set; } = null!;
    [Inject] private NavigationManager Navigation { get; set; } = null!;
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
        ActiveTool = PdfToolCatalog.Find(GetQueryValue("tool"));
    }

    private async Task SelectToolAsync(PdfToolDefinition tool) {
        if (ActiveTool.Id == tool.Id) return;
        await ResetResultAsync();
        ActiveTool = tool;
        Files.Clear();
        ResetSettings();
        Diagnostics.Clear();
        var values = new Dictionary<string, object?> {
            ["workspace"] = "pdf",
            ["tool"] = tool.Id,
            ["route"] = null
        };
        Navigation.NavigateTo(Navigation.GetUriWithQueryParameters(values), replace: true);
    }

    private async Task HandleFilesSelectedAsync(InputFileChangeEventArgs args) {
        await ResetResultAsync();
        Diagnostics.Clear();
        IReadOnlyList<IBrowserFile> selected = ActiveTool.InputMode == PdfToolInputMode.Single
            ? [args.File]
            : args.GetMultipleFiles(ActiveTool.InputMode == PdfToolInputMode.Pair ? 2 : BrowserPdfToolService.MaxPdfFiles);
        var loaded = new List<SelectedDocument>(selected.Count);
        try {
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
            Files.Clear();
            Files.AddRange(loaded);
            Diagnostics.Add(new ConversionDiagnostic("Ready", $"{Files.Count} PDF file{(Files.Count == 1 ? string.Empty : "s")} loaded in this tab.", "ocx-dot--good"));
        } catch (Exception ex) {
            Files.Clear();
            Diagnostics.Add(new ConversionDiagnostic("Could not load PDFs", DescribeFailure(ex), "ocx-dot--bad"));
        }
    }

    private async Task LoadSampleAsync() {
        await ResetResultAsync();
        Diagnostics.Clear();
        try {
            byte[] bytes = await Http.GetByteArrayAsync("samples/showcase-dashboard.pdf");
            Files.Clear();
            Files.Add(CreateSample(bytes, ActiveTool.InputMode == PdfToolInputMode.Pair ? "expected" : "showcase"));
            if (ActiveTool.InputMode != PdfToolInputMode.Single) {
                Files.Add(CreateSample((byte[])bytes.Clone(), ActiveTool.InputMode == PdfToolInputMode.Pair ? "actual" : "showcase-copy"));
            }
            if (ActiveTool.Kind == PdfToolKind.Redact) {
                RedactionText = "Critical blockers";
            }
            Diagnostics.Add(new ConversionDiagnostic("Sample ready", $"{Files.Count} product PDF file{(Files.Count == 1 ? string.Empty : "s")} loaded locally.", "ocx-dot--good"));
        } catch (Exception ex) {
            Diagnostics.Add(new ConversionDiagnostic("Could not load sample", DescribeFailure(ex), "ocx-dot--bad"));
        }
    }

    private static SelectedDocument CreateSample(byte[] bytes, string suffix) =>
        new($"officeimo-{suffix}.pdf", ".pdf", "PDF", bytes.LongLength, bytes);

    private async Task RemoveFileAsync(int index) {
        if (index < 0 || index >= Files.Count) return;
        await ResetResultAsync();
        Files.RemoveAt(index);
        Diagnostics.Clear();
    }

    private async Task MoveFileAsync(PdfFileMoveRequest request) {
        int target = request.Index + request.Offset;
        if (request.Index < 0 || request.Index >= Files.Count || target < 0 || target >= Files.Count) return;
        await ResetResultAsync();
        SelectedDocument file = Files[request.Index];
        Files.RemoveAt(request.Index);
        Files.Insert(target, file);
    }

    private async Task ClearFilesAsync() {
        await ResetResultAsync();
        Files.Clear();
        Diagnostics.Clear();
    }

    private async Task RunAsync() {
        if (!CanRun || _interop is null) return;
        IsBusy = true;
        await ResetResultAsync();
        Diagnostics.Clear();
        await InvokeAsync(StateHasChanged);
        await Task.Yield();
        try {
            Result = PdfTools.Execute(new PdfToolRequest(
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
            ArtifactUrl = await _interop.CreateObjectUrlAsync(Result.Artifact.Bytes, Result.Artifact.ContentType);
            if (Result.Report is not null) {
                ReportUrl = await _interop.CreateObjectUrlAsync(Result.Report.Bytes, Result.Report.ContentType);
            }
            Diagnostics.Add(new ConversionDiagnostic("Operation complete", Result.Summary, "ocx-dot--good"));
        } catch (Exception ex) {
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

    private async Task ResetResultAsync() {
        if (_interop is not null) {
            await _interop.RevokeObjectUrlAsync(ArtifactUrl);
            await _interop.RevokeObjectUrlAsync(ReportUrl);
        }
        ArtifactUrl = null;
        ReportUrl = null;
        Result = null;
    }

    private string GetQueryValue(string name) {
        string query = new Uri(Navigation.Uri).Query;
        if (query.Length <= 1) return string.Empty;
        foreach (string pair in query[1..].Split('&', StringSplitOptions.RemoveEmptyEntries)) {
            string[] parts = pair.Split('=', 2);
            if (string.Equals(Uri.UnescapeDataString(parts[0]), name, StringComparison.OrdinalIgnoreCase)) {
                return parts.Length == 2 ? Uri.UnescapeDataString(parts[1].Replace("+", " ")) : string.Empty;
            }
        }
        return string.Empty;
    }

    private static string DescribeFailure(Exception ex) => ex switch {
        IOException => "The browser workbench accepts PDFs up to 25 MB each.",
        _ => ex.Message
    };

    public async ValueTask DisposeAsync() {
        if (_interop is null) return;
        await _interop.RevokeObjectUrlAsync(ArtifactUrl);
        await _interop.RevokeObjectUrlAsync(ReportUrl);
        await _interop.DisposeAsync();
    }
}
