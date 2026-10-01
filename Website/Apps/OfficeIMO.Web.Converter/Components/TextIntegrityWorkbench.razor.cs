using System.Text;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;
using Microsoft.JSInterop;
using OfficeIMO.Provenance;
using OfficeIMO.Workflows;
using OfficeIMO.Web.Converter.Services;

namespace OfficeIMO.Web.Converter.Components;

public partial class TextIntegrityWorkbench {
    [Inject] private IJSRuntime JS { get; set; } = null!;
    private ConverterInterop? _interop;
    private OfficeTextIntegrityReview? _review;
    private readonly HashSet<int> _selected = [];
    private string _text = "", _sourceName = "text.txt", _message = "Inspect text to review exact Unicode findings.";
    private byte[]? _sourceBytes;
    private string? _cleaned, _outputUrl, _reportUrl;
    private bool _busy, _disposed;
    private int _generation;
    private static OfficeTextIntegrityOptions Limits() => new() { MaxEncodedBytes = 1024 * 1024, MaxCharacters = 262144, MaxFindings = 512 };
    private string OutputName => Path.GetFileNameWithoutExtension(_sourceName) + "-text-cleaned" + Path.GetExtension(_sourceName);
    private string MarkedText {
        get {
            if (_review == null) return "";
            var result = new StringBuilder(); int offset = 0;
            foreach (OfficeTextIntegrityFinding finding in _review.Report.Findings) {
                result.Append(_text, offset, finding.TextOffset - offset);
                result.Append('⟦').Append(finding.UnicodeNotation).Append(' ').Append(finding.Kind).Append('⟧');
                offset = finding.TextOffset + finding.TextLength;
            }
            result.Append(_text, offset, _text.Length - offset); return result.ToString();
        }
    }
    protected override void OnInitialized() => _interop = new(JS);
    private async Task TextChangedAsync(ChangeEventArgs args) {
        if (_busy || _disposed) return;
        _text = args.Value?.ToString() ?? ""; _sourceBytes = null; _sourceName = "text.txt";
        await InvalidateAsync();
    }
    private async Task LoadSampleAsync() {
        if (_busy || _disposed) return;
        _text = "Invoice\u202E123\u202C\nLanguage joiner: a\u200Db\nEmoji: ❤️\nNon-breaking space: hello\u00A0world";
        _sourceBytes = null; _sourceName = "sample.txt"; await InvalidateAsync();
        _message = "Sample loaded. Review controls in context; nothing is selected automatically.";
    }
    private async Task OpenAsync(InputFileChangeEventArgs args) {
        if (_busy || _disposed) return;
        int generation = ++_generation; _busy = true;
        try {
            await using var input = args.File.OpenReadStream(1024 * 1024);
            using var buffer = new MemoryStream(); await input.CopyToAsync(buffer);
            byte[] bytes = buffer.ToArray();
            var review = OfficeTextIntegrityReview.Inspect(bytes, Limits(), args.File.Name);
            if (_disposed || generation != _generation) return;
            _review = null; _selected.Clear(); _cleaned = null;
            await ClearUrlsAsync();
            if (_disposed || generation != _generation) return;
            _text = review.Text; _sourceBytes = bytes; _sourceName = Path.GetFileName(args.File.Name);
            _message = "Text file decoded. Inspect it to select exact occurrences.";
        } catch (Exception error) when (error is not OutOfMemoryException) {
            if (!_disposed && generation == _generation) _message = "Could not open text: " + error.Message;
        } finally { if (!_disposed) _busy = false; }
    }
    private async Task InspectAsync() {
        if (_busy || _disposed || _interop == null) return;
        int generation = ++_generation; _busy = true;
        try {
            await ClearUrlsAsync(); await Task.Yield();
            if (_disposed || generation != _generation) return;
            var review = _sourceBytes == null ? OfficeTextIntegrityReview.Inspect(_text, Limits(), _sourceName) :
                OfficeTextIntegrityReview.Inspect(_sourceBytes, Limits(), _sourceName);
            await using var urls = new ConverterObjectUrlBatch(_interop, () => !_disposed && generation == _generation);
            string url = await urls.CreateAsync(Encoding.UTF8.GetBytes(OfficeTextIntegrityReportSerializer.Serialize(review, _text, _sourceName)), "application/json");
            urls.Commit(); _review = review; _selected.Clear(); _cleaned = null; _reportUrl = url;
            _message = review.Report.Findings.Count == 0 ? "Text inspection completed with no findings under the selected Unicode policy." : "Review each occurrence before selecting removal. Your original stays unchanged.";
        } catch (Exception error) when (error is not OutOfMemoryException) {
            if (!_disposed && generation == _generation) { _review = null; _message = "Text inspection failed: " + error.Message; }
        } finally { if (!_disposed && generation == _generation) _busy = false; }
    }
    private async Task SelectionChangedAsync(int index, ChangeEventArgs args) {
        if (_busy || _disposed || _review == null || index < 0 || index >= _review.Report.Findings.Count) return;
        if (args.Value is true) _selected.Add(index); else _selected.Remove(index);
        _generation++; _cleaned = null; await ClearUrlsAsync();
        _message = "Selection changed. Create a new copy and report for these occurrences.";
    }
    private async Task CreateCopyAsync() {
        if (_review == null || _selected.Count == 0 || _busy || _interop == null || _disposed) return;
        int generation = ++_generation; _busy = true;
        try {
            await ClearUrlsAsync();
            string cleaned = _review.RemoveSelected(_text, _selected);
            byte[] output = _review.ExportSelected(_text, _selected);
            await using var urls = new ConverterObjectUrlBatch(_interop, () => !_disposed && generation == _generation);
            string outputUrl = await urls.CreateAsync(output, "application/octet-stream");
            string reportUrl = await urls.CreateAsync(Encoding.UTF8.GetBytes(OfficeTextIntegrityReportSerializer.Serialize(_review, _text, _sourceName, _selected)), "application/json");
            urls.Commit(); _outputUrl = outputUrl; _reportUrl = reportUrl; _cleaned = cleaned;
            _message = "A separate text copy is ready. Review it before downloading.";
        } catch (Exception error) when (error is not OutOfMemoryException) {
            if (!_disposed && generation == _generation) _message = "Could not create text copy: " + error.Message;
        } finally { if (!_disposed && generation == _generation) _busy = false; }
    }
    private async Task InvalidateAsync() { _generation++; _review = null; _selected.Clear(); _cleaned = null; _message = "Text changed. Inspect the current text before selecting removal."; await ClearUrlsAsync(); }
    private async Task ClearUrlsAsync() {
        string? output = _outputUrl, report = _reportUrl; _outputUrl = _reportUrl = null;
        if (_interop != null) { await _interop.RevokeObjectUrlAsync(output); await _interop.RevokeObjectUrlAsync(report); }
    }
    public async ValueTask DisposeAsync() { _disposed = true; _generation++; await ClearUrlsAsync(); if (_interop != null) await _interop.DisposeAsync(); }
}
