using System.Diagnostics;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.JSInterop;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;

namespace OfficeIMO.Web.Converter.Components;

public partial class ConverterWorkspace {
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
  <p>This HTML becomes <strong>portable Markdown</strong>.</p>
  <ul><li>Headings</li><li>Lists</li><li>Links</li></ul>
</article>
""";

    [Inject] private HttpClient Http { get; set; } = null!;
    [Inject] private IJSRuntime JS { get; set; } = null!;
    [Inject] private NavigationManager Navigation { get; set; } = null!;
    [Inject] private BrowserConversionService ConversionService { get; set; } = null!;

    private ConverterInterop? _interop;
    private ConversionRoute ActiveRoute { get; set; } = ConversionRouteCatalog.Default;
    private SelectedDocument? SelectedFile { get; set; }
    private ConversionResult? Output { get; set; }
    private string? OutputUrl { get; set; }
    private string OutputFileName { get; set; } = "officeimo-output";
    private string TextInput { get; set; } = DefaultMarkdown;
    private bool FastPreview { get; set; }
    private bool PreviewOutput { get; set; } = true;
    private bool IsBusy { get; set; }
    private long ElapsedMilliseconds { get; set; }
    private List<ConversionDiagnostic> Diagnostics { get; } = [ReadyDiagnostic()];

    private static IReadOnlyList<ConversionRoute> Routes => ConversionRouteCatalog.All;
    private bool CanConvert => !IsBusy && (ActiveRoute.InputKind == ConversionInputKind.File ? SelectedFile is not null : !string.IsNullOrWhiteSpace(TextInput));
    private string OutputHeading => Output?.FileName ?? $"{ActiveRoute.Target} output";
    private string ElapsedLabel => ElapsedMilliseconds < 1000 ? $"{ElapsedMilliseconds} ms" : $"{ElapsedMilliseconds / 1000d:0.0} s";

    protected override void OnInitialized() {
        _interop = new ConverterInterop(JS);
        ActiveRoute = ConversionRouteCatalog.Find(GetQueryValue("route"));
        TextInput = ActiveRoute.Id == "html-markdown" ? DefaultHtml : DefaultMarkdown;
    }

    private async Task SelectRouteAsync(ConversionRoute route) {
        if (ActiveRoute.Id == route.Id) {
            return;
        }
        await ResetOutputAsync();
        ActiveRoute = route;
        SelectedFile = null;
        TextInput = route.Id == "html-markdown" ? DefaultHtml : DefaultMarkdown;
        Diagnostics.Clear();
        Diagnostics.Add(ReadyDiagnostic());
        Navigation.NavigateTo($"?route={Uri.EscapeDataString(route.Id)}", replace: true);
    }

    private async Task HandleFileSelectedAsync(InputFileChangeEventArgs args) {
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
            SelectedFile = new(file.Name, extension, ActiveRoute.Source, file.Size, buffer.ToArray());
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
        SampleDocument sample = ActiveRoute.Id switch {
            "xlsx-pdf" => new("Sample XLSX", "samples/basic.xlsx", "OfficeIMO-Basic.xlsx", ".xlsx"),
            "pptx-pdf" => new("Sample PPTX", "samples/basic.pptx", "OfficeIMO-Basic.pptx", ".pptx"),
            _ => new("Sample DOCX", "samples/basic.docx", "OfficeIMO-Basic.docx", ".docx")
        };
        await ResetOutputAsync();
        Diagnostics.Clear();
        try {
            byte[] bytes = await Http.GetByteArrayAsync(sample.Path);
            SelectedFile = new(sample.FileName, sample.Extension, ActiveRoute.Source, bytes.LongLength, bytes);
            Diagnostics.Add(new("Sample ready", $"{sample.FileName} is loaded locally.", "ocx-dot--good"));
        } catch (Exception ex) {
            Diagnostics.Add(new("Could not load sample", DescribeFailure(ex), "ocx-dot--bad"));
        }
    }

    private async Task LoadTextSampleAsync() {
        await ResetOutputAsync();
        TextInput = ActiveRoute.Id == "html-markdown" ? DefaultHtml : DefaultMarkdown;
        Diagnostics.Clear();
        Diagnostics.Add(new("Sample ready", $"Sample {ActiveRoute.Source} is ready.", "ocx-dot--good"));
    }

    private async Task ConvertAsync() {
        if (!CanConvert || _interop is null) {
            return;
        }

        IsBusy = true;
        await ResetOutputAsync();
        Diagnostics.Clear();
        await InvokeAsync(StateHasChanged);
        await Task.Yield();
        var stopwatch = Stopwatch.StartNew();

        try {
            Output = ActiveRoute.InputKind == ConversionInputKind.File
                ? ConversionService.ConvertFile(ActiveRoute, SelectedFile!, FastPreview)
                : ConversionService.ConvertText(ActiveRoute, TextInput);
            stopwatch.Stop();
            ElapsedMilliseconds = stopwatch.ElapsedMilliseconds;
            OutputFileName = Output.FileName;
            OutputUrl = await _interop.CreateObjectUrlAsync(Output.Bytes, Output.ContentType);
            Diagnostics.Add(new("Conversion complete", $"Created {Output.FileName} locally in {ElapsedLabel}.", "ocx-dot--good"));
            foreach (string warning in Output.Warnings.Take(8)) {
                Diagnostics.Add(new("Review this result", warning, "ocx-dot--warn"));
            }
        } catch (Exception ex) {
            stopwatch.Stop();
            Output = null;
            Diagnostics.Add(new("Conversion failed", DescribeFailure(ex), "ocx-dot--bad"));
        } finally {
            IsBusy = false;
        }
    }

    private async Task ResetOutputAsync() {
        if (_interop is not null && !string.IsNullOrWhiteSpace(OutputUrl)) {
            await _interop.RevokeObjectUrlAsync(OutputUrl);
        }
        OutputUrl = null;
        Output = null;
        ElapsedMilliseconds = 0;
    }

    private string GetQueryValue(string name) {
        string query = new Uri(Navigation.Uri).Query;
        if (query.Length <= 1) {
            return string.Empty;
        }
        foreach (string pair in query[1..].Split('&', StringSplitOptions.RemoveEmptyEntries)) {
            string[] parts = pair.Split('=', 2);
            if (string.Equals(Uri.UnescapeDataString(parts[0]), name, StringComparison.OrdinalIgnoreCase)) {
                return parts.Length == 2 ? Uri.UnescapeDataString(parts[1].Replace("+", " ")) : string.Empty;
            }
        }
        return string.Empty;
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

    private static ConversionDiagnostic ReadyDiagnostic() =>
        new("Ready", "Choose a source and run the conversion. Nothing is uploaded.", "ocx-dot--good");

    public async ValueTask DisposeAsync() {
        if (_interop is not null) {
            await _interop.RevokeObjectUrlAsync(OutputUrl);
            await _interop.DisposeAsync();
        }
    }
}
