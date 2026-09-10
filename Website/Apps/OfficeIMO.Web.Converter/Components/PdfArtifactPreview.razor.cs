using Microsoft.AspNetCore.Components;
using OfficeIMO.Pdf;
using OfficeIMO.Web.Converter.Services;

namespace OfficeIMO.Web.Converter.Components;

public partial class PdfArtifactPreview {
    [Parameter, EditorRequired] public byte[] Bytes { get; set; } = [];
    [Parameter] public string FileName { get; set; } = "document.pdf";

    private byte[]? _source;
    private BrowserPdfPreview? _preview;
    private string? _image;
    private int _page = 1;
    private int _generation;
    private bool _busy;
    private bool _hasWarnings;

    protected override async Task OnParametersSetAsync() {
        if (ReferenceEquals(_source, Bytes)) return;
        _source = Bytes;
        _preview = null;
        _page = 1;
        await RenderAsync();
    }

    private Task PreviousAsync() => ChangePageAsync(-1);
    private Task NextAsync() => ChangePageAsync(1);

    private async Task ChangePageAsync(int delta) {
        if (_busy || _preview is null) return;
        _page = Math.Clamp(_page + delta, 1, _preview.PageCount);
        await RenderAsync();
    }

    private async Task RenderAsync() {
        int generation = ++_generation;
        _image = null;
        _hasWarnings = false;
        _busy = true;
        // Yield once so the loading state is rendered before the bounded synchronous page render.
        await Task.Yield();
        if (generation != _generation) return;
        try {
            _preview ??= new BrowserPdfPreview(_source!);
            PdfPageRenderResult result = _preview.Render(_page);
            byte[]? image = result.Bytes;
            if (result.Succeeded && image is not null) _image = "data:image/png;base64," + Convert.ToBase64String(image);
            _hasWarnings = result.Diagnostics.Count > 0;
        } catch (Exception ex) when (ex is not OutOfMemoryException) {
            // A preview failure must not hide a successfully generated downloadable artifact.
            _image = null;
        } finally {
            if (generation == _generation) _busy = false;
        }
    }

    public void Dispose() {
        ++_generation;
        _source = null;
        _preview = null;
        _image = null;
    }
}
