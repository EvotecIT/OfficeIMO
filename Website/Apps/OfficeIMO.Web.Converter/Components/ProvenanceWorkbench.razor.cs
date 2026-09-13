using System.Text.Json;
using OfficeIMO.Drawing;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.JSInterop;
using OfficeIMO.Provenance;
using OfficeIMO.Workflows;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;

namespace OfficeIMO.Web.Converter.Components;

public partial class ProvenanceWorkbench {
    [Inject] private BrowserDocumentSession Session { get; set; } = null!;
    [Inject] private HttpClient Http { get; set; } = null!;
    [Inject] private IJSRuntime JS { get; set; } = null!;
    [Parameter] public int Revision { get; set; }
    private ConverterInterop? _interop;
    private SelectedDocument? _file;
    private OfficeProvenanceReport? _report;
    private OfficeProvenanceRemovalResult? _removal;
    private byte[]? _outputBytes;
    private string? _outputUrl, _outputName, _reportUrl, _previewUrl;
    private string _message = "Choose a file or reuse a compatible working file from this session.";
    private bool _busy, _disposed;
    private bool _manifests = true, _references = true, _declarations = true;
    private int _revision = -1, _generation;

    protected override void OnInitialized() => _interop = new(JS);
    protected override async Task OnParametersSetAsync() {
        if (_revision == Session.Revision) return;
        _revision = Session.Revision;
        int generation = ++_generation;
        await ClearResultAsync();
        if (_disposed || generation != _generation || _revision != Session.Revision) return;
        _report = null;
        _file = Session.Current.FirstOrDefault(file => OfficeProvenanceBufferWorkflow.SupportedExtensions.Contains(file.Extension, StringComparer.OrdinalIgnoreCase));
        _message = _file is null ? "Choose a supported file to inspect." : "Working file ready for inspection.";
        _busy = false;
    }
    private async Task OpenAsync(InputFileChangeEventArgs args) {
        int generation = ++_generation;
        int sourceRevision = Session.Revision;
        bool IsCurrent() => !_disposed && generation == _generation && sourceRevision == Session.Revision;
        _busy = true;
        try {
            var file = args.File;
            string extension = Path.GetExtension(file.Name).ToLowerInvariant();
            if (!OfficeProvenanceBufferWorkflow.SupportedExtensions.Contains(extension)) throw new NotSupportedException("Choose JPEG, PNG, WebP, PDF, DOCX, XLSX, or PPTX.");
            await using var input = file.OpenReadStream(BrowserConversionService.MaxPackageBytes);
            using var memory = new MemoryStream(); await input.CopyToAsync(memory);
            if (!IsCurrent()) return;
            await ClearResultAsync();
            if (!IsCurrent()) return;
            _report = null;
            _file = new(file.Name, extension, extension.TrimStart('.').ToUpperInvariant(), memory.Length, memory.ToArray());
            Session.Open([_file]); _revision = Session.Revision;
            _message = "File ready. Inspect it to see supported provenance.";
        } catch (Exception ex) when (ex is not OutOfMemoryException) {
            if (IsCurrent()) _message = "Could not open file: " + ex.Message;
        } finally { if (generation == _generation) _busy = false; }
    }
    private async Task LoadSampleAsync() {
        if (_busy) return;
        int generation = ++_generation;
        int sourceRevision = Session.Revision;
        bool IsCurrent() => !_disposed && generation == _generation && sourceRevision == Session.Revision;
        _busy = true;
        try {
            byte[] bytes = await Http.GetByteArrayAsync("samples/provenance-demo.png");
            if (!IsCurrent()) return;
            await ClearResultAsync();
            if (!IsCurrent()) return;
            _report = null;
            _file = new("provenance-demo.png", ".png", "PNG", bytes.LongLength, bytes);
            Session.Open([_file]); _revision = Session.Revision;
            _message = "Sample ready. Its AI source declaration is test metadata added for this demonstration.";
        } catch (Exception ex) { if (IsCurrent()) _message = "Could not load sample: " + ex.Message; }
        finally { if (generation == _generation) _busy = false; }
    }
    private Task InspectAsync() => RunAsync(false);
    private Task RemoveAsync() => RunAsync(true);
    private async Task RunAsync(bool remove) {
        if (_file is null || _busy || _interop is null) return;
        var file = _file;
        int generation = ++_generation;
        int sourceRevision = Session.Revision;
        Session.ClearResult();
        _busy = true;
        bool IsCurrent() => !_disposed && generation == _generation && sourceRevision == Session.Revision;
        await using var urls = new ConverterObjectUrlBatch(_interop, IsCurrent);
        try {
            await ClearResultAsync();
            await Task.Yield();
            if (!IsCurrent()) return;
            OfficeProvenanceRemovalResult? removal = null;
            OfficeProvenanceReport report;
            byte[]? output = null;
            string? outputName = null, outputUrl = null, previewUrl = null;
            string message;
            if (remove) {
                removal = OfficeProvenanceBufferWorkflow.Remove(file.Bytes, file.Name, BrowserProvenancePolicy.Removal(_manifests, _references, _declarations));
                report = removal.After; output = removal.ToArray();
                outputName = Path.GetFileNameWithoutExtension(file.Name) + "-provenance-cleaned" + file.Extension;
                outputUrl = await urls.CreateAsync(output, ContentType(file.Extension));
                message = removal.WasChanged ? "A separate copy is ready. Review remaining findings before downloading." : "No selected carriers were removed. The copy retains the original data; review the findings and diagnostics.";
            } else {
                report = OfficeProvenanceBufferWorkflow.Inspect(file.Bytes, file.Name, BrowserProvenancePolicy.Limits());
                message = report.Evidence.Count == 0 ? "No supported provenance carriers were found. This is not proof of origin." : "Inspection complete. Choose what to remove from a copy.";
            }
            if (!IsCurrent()) return;
            byte[] reportBytes = JsonSerializer.SerializeToUtf8Bytes(new {
                schemaVersion = 1, operation = remove ? "remove" : "inspect", fileName = file.Name,
                before = removal?.Before ?? report, after = removal?.After, changes = removal?.Changes,
                structuralOnly = true, externalReferencesFetched = false
            }, new JsonSerializerOptions { WriteIndented = true });
            string reportUrl = await urls.CreateAsync(reportBytes, "application/json");
            if (!IsCurrent()) return;
            if (file.Extension is ".jpg" or ".jpeg" or ".png" or ".webp") {
                try {
                    var info = OfficeImageReader.Identify(output ?? file.Bytes, file.Name);
                    if (info.Width > 0 && info.Height > 0 && info.Width <= 8192 && info.Height <= 8192 && (long)info.Width * info.Height <= 16_000_000)
                        previewUrl = await urls.CreateAsync(output ?? file.Bytes, ContentType(file.Extension));
                    else message += " Image preview omitted because its dimensions exceed the browser preview budget.";
                } catch (Exception ex) when (ex is not OutOfMemoryException) {
                    message += " Image preview is unavailable; the inspection and download remain available.";
                }
            }
            if (!IsCurrent()) return;
            urls.Commit();
            _report = report; _removal = removal; _outputBytes = output;
            _outputName = outputName; _outputUrl = outputUrl; _reportUrl = reportUrl; _previewUrl = previewUrl;
            _message = message;
            if (output is not null) Session.SetResult(output, outputName!, sourceRevision);
        } catch (Exception ex) when (ex is not OutOfMemoryException) {
            if (IsCurrent()) { _report = null; _message = "Operation could not complete: " + ex.Message; }
        } finally {
            if (generation == _generation) _busy = false;
        }
    }
    private async Task OptionsChangedAsync() {
        int generation = ++_generation;
        int sourceRevision = Session.Revision;
        _report = _removal?.Before ?? _report; Session.ClearResult();
        await ClearResultAsync();
        if (!_disposed && generation == _generation && sourceRevision == Session.Revision)
            _message = "Removal options changed. Create a new copy to apply them.";
    }
    private async Task ClearResultAsync() {
        var urls = new[] { _outputUrl, _reportUrl, _previewUrl };
        _outputUrl = _reportUrl = _previewUrl = _outputName = null; _outputBytes = null; _removal = null;
        if (_interop is not null)
            foreach (var url in urls) await _interop.RevokeObjectUrlAsync(url);
    }
    private static string ContentType(string extension) => extension switch {
        ".jpg" or ".jpeg" => "image/jpeg", ".png" => "image/png", ".webp" => "image/webp", ".pdf" => "application/pdf", _ => "application/octet-stream"
    };
    public async ValueTask DisposeAsync() {
        _disposed = true; ++_generation;
        await ClearResultAsync();
        if (_interop is not null) await _interop.DisposeAsync();
    }
}
