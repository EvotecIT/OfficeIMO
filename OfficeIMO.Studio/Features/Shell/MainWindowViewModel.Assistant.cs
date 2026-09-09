using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Assistant;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private DocumentAssistantViewModel? _assistant;
    internal DocumentAssistantViewModel Assistant => _assistant ??= new(_services.AiConnections, CaptureAssistantSource,
        page => { if (page >= 1 && page <= Pages.Count) { SelectedPage = Pages[page - 1]; WorkspaceMode = StudioWorkspaceMode.PdfWorkspace; } }, _localizer);
    [ObservableProperty] private bool _isAssistantVisible;
    [RelayCommand] private void ToggleAssistant() {
        _services.AiConnections.OpenUri = _openUri;
        IsAssistantVisible = !IsAssistantVisible;
    }
    partial void OnIsAssistantVisibleChanged(bool value) { if (!value) _assistant?.Deactivate(); }
    internal void DeactivateAssistant() => _assistant?.Deactivate();
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
                && (!currentPageOnly || SelectedPage?.PageNumber == page), workspace.CreateReaderOptions());
    }
}
