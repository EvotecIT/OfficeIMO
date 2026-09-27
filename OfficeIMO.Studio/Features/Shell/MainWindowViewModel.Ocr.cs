using CommunityToolkit.Mvvm.Input;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Suggests OCR when the reader lands on a scanned page, and starts it with one click.</summary>
public sealed partial class MainWindowViewModel {
    private string? _ocrPromptDismissedFor;

    public bool ShowOcrPrompt => HasDocument && IsPdfWorkspaceMode && !IsComparisonOpen &&
        SelectedPage?.IsImageOnly == true && !string.Equals(_ocrPromptDismissedFor, DocumentPath, StringComparison.Ordinal);

    [RelayCommand]
    private async Task MakeSearchableAsync() {
        if (_workspace is null || IsWorkspaceBusy || IsOpening) return;
        if (IsDirty || HasFormDrafts) {
            ErrorMessage = _localizer.GetOrDefault("Assistant.OcrSaveFirst",
                "Save your current changes and apply form drafts before opening OCR. OCR reads the saved source and creates a separate output.");
            return;
        }
        ShowOcr();
        if (OcrWorkbench.RunCommand.CanExecute(null)) await OcrWorkbench.RunCommand.ExecuteAsync(null).ConfigureAwait(true);
    }

    [RelayCommand]
    private void DismissOcrPrompt() {
        _ocrPromptDismissedFor = DocumentPath;
        OnPropertyChanged(nameof(ShowOcrPrompt));
    }
}
