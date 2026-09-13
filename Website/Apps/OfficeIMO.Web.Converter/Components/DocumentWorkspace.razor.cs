using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Routing;
using Microsoft.AspNetCore.Components.Web;
using Microsoft.JSInterop;
using OfficeIMO.Web.Converter.Models;
using OfficeIMO.Web.Converter.Services;

namespace OfficeIMO.Web.Converter.Components;

public partial class DocumentWorkspace {
    [Inject] private BrowserDocumentSession Session { get; set; } = default!;
    [Inject] private NavigationManager Navigation { get; set; } = default!;
    [Inject] private IJSRuntime JS { get; set; } = default!;
    private IJSObjectReference? _module;
    private DotNetObjectReference<DocumentWorkspace>? _reference;
    private bool _disposed;
    private bool _notifyLocation = true;
    private bool _initialLocation = true;
    private string? _selectionKey;
    private bool IsMenuOpen { get; set; }
    private bool IsPdfWorkspace { get; set; }
    private bool IsProvenanceWorkspace { get; set; }
    private bool IsFocused { get; set; }
    private ConversionRoute ActiveRoute { get; set; } = ConversionRouteCatalog.Default;
    private PdfToolDefinition ActiveTool { get; set; } = PdfToolCatalog.Default;
    private string WorkspaceId => IsProvenanceWorkspace ? "provenance" : IsPdfWorkspace ? "pdf" : "convert";
    private string LibraryId => IsPdfWorkspace ? "pdf" : ActiveRoute.Source switch {
        "DOCX" => "word", "XLSX" => "excel", "PPTX" => "powerpoint", "PDF" => "pdf", "MD" => "markdown", _ => "html"
    };
    private string LibraryLabel => LibraryId switch {
        "word" => "Word library", "excel" => "Excel library", "powerpoint" => "PowerPoint library",
        "pdf" => "PDF library", "markdown" => "Markdown library", _ => "HTML library"
    };
    private string LibraryUrl => $"/products/{LibraryId}/";
    private string GuideUrl => IsPdfWorkspace
        ? $"/pdf/{ActiveTool.Id switch { "extract" => "extract-pages", "delete" => "delete-pages", "reorder" => "reorder-pages", "rotate" => "rotate-pages", _ => ActiveTool.Id }}/"
        : ActiveRoute.Id switch {
            "docx-pdf" => "/convert/word-to-pdf/", "xlsx-pdf" => "/convert/excel-to-pdf/",
            "pdf-html" => "/convert/pdf-to-html/", "pdf-docx" => "/convert/pdf-to-word/",
            "pdf-pptx" => "/convert/pdf-to-powerpoint/", "pdf-xlsx" => "/convert/pdf-tables-to-excel/",
            _ => "/convert/guides/"
        };

    protected override void OnInitialized() {
        Session.Changed += SessionChanged;
        ReadLocation();
        Navigation.LocationChanged += LocationChanged;
    }
    private void ReadLocation() {
        var query = System.Web.HttpUtility.ParseQueryString(new Uri(Navigation.Uri).Query);
        IsPdfWorkspace = string.Equals(query["workspace"], "pdf", StringComparison.OrdinalIgnoreCase);
        IsProvenanceWorkspace = string.Equals(query["workspace"], "provenance", StringComparison.OrdinalIgnoreCase);
        ActiveRoute = ConversionRouteCatalog.Find(query["route"]);
        ActiveTool = PdfToolCatalog.Find(query["tool"]);
        string key = WorkspaceId + ":" + (IsPdfWorkspace ? ActiveTool.Id : IsProvenanceWorkspace ? "" : ActiveRoute.Id);
        bool changed = _selectionKey is not null && _selectionKey != key;
        _selectionKey = key;
        if (changed) Session.ChangeTool();
    }
    private void LocationChanged(object? sender, LocationChangedEventArgs args) {
        if (_disposed) return;
        ReadLocation();
        IsFocused = false;
        IsMenuOpen = false;
        _notifyLocation = true;
        _ = InvokeAsync(StateHasChanged);
    }
    private string SelectionUrl(bool pdf, string id) => Navigation.GetUriWithQueryParameters(new Dictionary<string, object?> {
        ["workspace"] = pdf ? "pdf" : null, ["tool"] = pdf ? id : null, ["route"] = pdf ? null : id
    });
    private void Select(bool pdf, string id) {
        IsMenuOpen = false;
        Navigation.NavigateTo(SelectionUrl(pdf, id), replace: true);
    }
    /// <summary>Restores the host page's selection after browser history navigation.</summary>
    [JSInvokable]
    public Task RestoreSelection(string? workspace, string? route, string? tool) {
        if (!_disposed) {
            if (string.Equals(workspace, "provenance", StringComparison.OrdinalIgnoreCase)) {
                IsMenuOpen = false;
                Navigation.NavigateTo(Navigation.GetUriWithQueryParameters(new Dictionary<string, object?> {
                    ["workspace"] = "provenance", ["route"] = null, ["tool"] = null
                }), replace: true);
                return Task.CompletedTask;
            }
            bool pdf = string.Equals(workspace, "pdf", StringComparison.OrdinalIgnoreCase);
            Select(pdf, pdf ? PdfToolCatalog.Find(tool).Id : ConversionRouteCatalog.Find(route).Id);
        }
        return Task.CompletedTask;
    }
    protected override async Task OnAfterRenderAsync(bool firstRender) {
        if (firstRender) {
            _module = await JS.InvokeAsync<IJSObjectReference>("import", "./Components/DocumentWorkspace.razor.js");
            _reference = DotNetObjectReference.Create(this);
            await _module.InvokeVoidAsync("connect", _reference);
        }
        if (_module is not null && _notifyLocation) {
            _notifyLocation = false;
            await _module.InvokeVoidAsync("publishSelection", WorkspaceId, IsPdfWorkspace || IsProvenanceWorkspace ? null : ActiveRoute.Id, IsPdfWorkspace ? ActiveTool.Id : null, _initialLocation);
            _initialLocation = false;
        }
    }
    private async Task ToggleMenuAsync() {
        IsMenuOpen = !IsMenuOpen;
        if (_module is not null) await _module.InvokeVoidAsync("focusMenu", IsMenuOpen);
    }
    private async Task HandleKeyDown(KeyboardEventArgs args) {
        if (args.Key == "Escape" && IsMenuOpen) await ToggleMenuAsync();
    }
    public async ValueTask DisposeAsync() {
        Session.Changed -= SessionChanged;
        _disposed = true;
        Navigation.LocationChanged -= LocationChanged;
        if (_module is not null) {
            try { await _module.InvokeVoidAsync("disconnect"); await _module.DisposeAsync(); }
            catch (JSDisconnectedException) { }
        }
        _reference?.Dispose();
    }
    private void SessionChanged() { if (!_disposed) _ = InvokeAsync(StateHasChanged); }
}
