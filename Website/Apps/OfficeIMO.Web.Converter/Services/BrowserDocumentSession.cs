using OfficeIMO.Web.Converter.Models;

namespace OfficeIMO.Web.Converter.Services;

/// <summary>Tab-local source and result selection. No document data is persisted to browser storage.</summary>
public sealed class BrowserDocumentSession {
    public IReadOnlyList<SelectedDocument> Originals { get; private set; } = [];
    public IReadOnlyList<SelectedDocument> Current { get; private set; } = [];
    public SelectedDocument? LatestResult { get; private set; }
    public int Revision { get; private set; }
    public event Action? Changed;

    public void Open(IReadOnlyList<SelectedDocument> files) {
        ValidateSelection(files);
        Originals = files.ToArray(); Current = Originals; LatestResult = null; NotifySelection();
    }
    private static void ValidateSelection(IReadOnlyList<SelectedDocument> files) {
        if (files.Count > BrowserPdfToolService.MaxPdfFiles ||
            files.Any(file => file.Bytes.LongLength > BrowserConversionService.MaxPackageBytes) ||
            files.Sum(file => file.Bytes.LongLength) > BrowserPdfToolService.MaxAggregatePdfBytes)
            throw new InvalidDataException("The selected files exceed this browser workspace's limits.");
    }
    public void SelectCurrent(IReadOnlyList<SelectedDocument> files) {
        ValidateSelection(files);
        Current = files.ToArray(); LatestResult = null; NotifySelection();
    }
    public void ChangeTool() { LatestResult = null; NotifySelection(); }
    public void ClearResult() { if (LatestResult is null) return; LatestResult = null; Changed?.Invoke(); }
    public void SetResult(byte[] bytes, string name, int sourceRevision) {
        if (sourceRevision != Revision) return;
        // Larger outputs remain downloadable in the result panel but cannot become a new input.
        LatestResult = bytes.LongLength <= BrowserConversionService.MaxPackageBytes &&
            new[] { ".pdf", ".docx", ".xlsx", ".pptx", ".jpg", ".jpeg", ".png", ".webp", ".html", ".htm", ".md", ".markdown", ".txt" }.Contains(Path.GetExtension(name).ToLowerInvariant())
            ? new(name, Path.GetExtension(name).ToLowerInvariant(), Path.GetExtension(name).TrimStart('.').ToUpperInvariant(), bytes.LongLength, bytes)
            : null;
        Changed?.Invoke();
    }
    public void UseResult() {
        if (LatestResult is null) return;
        Current = [LatestResult]; LatestResult = null; NotifySelection();
    }
    public void RestoreOriginals() { Current = Originals; LatestResult = null; NotifySelection(); }
    public void Clear() { Originals = []; Current = []; LatestResult = null; NotifySelection(); }
    private void NotifySelection() { Revision++; Changed?.Invoke(); }
}
