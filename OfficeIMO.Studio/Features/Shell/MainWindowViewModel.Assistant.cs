using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Assistant;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private DocumentAssistantViewModel? _assistant;
    internal DocumentAssistantViewModel Assistant => _assistant ??= new(_services.AiConnections, CaptureAssistantSource,
        page => { if (page >= 1 && page <= Pages.Count) { SelectedPage = Pages[page - 1]; WorkspaceMode = StudioWorkspaceMode.PdfWorkspace; } }, _localizer) {
            OpenOcr = OpenAssistantOcrAsync
        };
    [ObservableProperty] private bool _isAssistantVisible;
    [RelayCommand] private void ToggleAssistant() {
        _services.AiConnections.OpenUri = _openUri;
        IsAssistantVisible = !IsAssistantVisible;
    }
    partial void OnIsAssistantVisibleChanged(bool value) {
        if (!value) _assistant?.Deactivate();
        else if (Assistant.CanPrepare) _ = Assistant.PrepareEvidenceCommand.ExecuteAsync(null);
    }
    internal void DeactivateAssistant() {
        if (IsAssistantVisible) IsAssistantVisible = false;
        else _assistant?.Deactivate();
    }
    private async Task OpenAssistantOcrAsync(CancellationToken token) {
        if (OcrWorkbench.IsBusy)
            throw new AssistantSourceUnavailableException(_localizer.GetOrDefault("Assistant.OcrBusy", "Finish or cancel the current OCR operation before opening another source."));
        if (_workspace is null || IsWorkspaceBusy || IsOpening || HasFormDrafts || IsDirty)
            throw new AssistantSourceUnavailableException(_localizer.GetOrDefault("Assistant.OcrSaveFirst", "Save your current changes and apply form drafts before opening OCR. OCR reads the saved source and creates a separate output."));
        var workspace = _workspace;
        string? path = DocumentPath;
        string? expectedHash = Assistant.PreparedDocument?.SourceHash;
        if (workspace.HasEncryption)
            throw new AssistantSourceUnavailableException(_localizer.GetOrDefault("Assistant.OcrProtected", "OCR cannot preserve this PDF's encryption. Use a separately authorized unprotected copy in the OCR workspace."));
        if (path is null || expectedHash is null) return;
        long revision = workspace.Revision;
        string actualHash = await _services.Storage.FingerprintAsync(path, token);
        token.ThrowIfCancellationRequested();
        if (_disposed || !ReferenceEquals(_workspace, workspace) || workspace.Revision != revision || IsDirty || HasFormDrafts
            || !string.Equals(expectedHash, actualHash, StringComparison.OrdinalIgnoreCase))
            throw new AssistantSourceUnavailableException(_localizer.GetOrDefault("Assistant.OcrSourceChanged", "The saved source changed after evidence preparation. Reopen the current file before using OCR."));
        OcrWorkbench.UseDocument(path, returnToAssistant: true);
        WorkspaceMode = StudioWorkspaceMode.Ocr;
        IsAssistantVisible = false;
    }
    private AssistantSource CaptureAssistantSource(bool currentPageOnly) {
        var workspace = _workspace;
        if (workspace is null || !HasDocument)
            throw new AssistantSourceUnavailableException(_localizer.GetOrDefault("Assistant.OpenPdf", "Open a PDF before asking a document question."));
        if (IsWorkspaceBusy || IsOpening || HasFormDrafts || !workspace.ViewInfo.CanExtractContent)
            throw new AssistantSourceUnavailableException(_localizer.GetOrDefault("Assistant.Unavailable", "Finish the current operation and apply form drafts first. The PDF must allow content extraction."));
        if (workspace.FileSize > new OfficeIMO.AI.OfficeAiLimits().MaxInputBytes)
            throw new AssistantSourceUnavailableException(_localizer.GetOrDefault("Assistant.TooLarge", "The assistant accepts PDF snapshots up to 32 MiB."));
        long revision = workspace.Revision;
        int? page = currentPageOnly ? SelectedPage?.PageNumber : null;
        if (currentPageOnly && page is null) throw new AssistantSourceUnavailableException(_localizer.GetOrDefault("Assistant.SelectPage", "Select a page first."));
        return new AssistantSource(workspace.CopyBytes(), workspace.FileName, page,
            () => !_disposed && ReferenceEquals(_workspace, workspace) && workspace.Revision == revision && !HasFormDrafts
                && (!currentPageOnly || SelectedPage?.PageNumber == page), workspace.CreateReaderOptions(page));
    }
}
