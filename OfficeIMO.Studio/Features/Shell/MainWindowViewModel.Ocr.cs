using CommunityToolkit.Mvvm.Input;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Suggests OCR when the reader lands on a scanned page, and starts it with one click.</summary>
public sealed partial class MainWindowViewModel {
    private string? _ocrPromptDismissedFor;
    private object? _ocrPaneOwner;
    private int _ocrPanePageNumber;
    private bool? _ocrPaneImageOnly;

    public bool ShowOcrPrompt => HasDocument && IsPdfWorkspaceMode && !IsComparisonOpen &&
        !IsReplacingReaderPresentation && ActiveReaderPageIsImageOnly &&
        !string.Equals(_ocrPromptDismissedFor, DocumentPath, StringComparison.Ordinal);

    // The pane reports only the active page's fact. Unknown means its current scene has not rendered.
    // The token identifies presentation ownership without retaining the pane or any page scene.
    private bool ActiveReaderPageIsImageOnly => _ocrPaneOwner is null ? SelectedPage?.IsImageOnly == true :
        _ocrPanePageNumber == SelectedPage?.PageNumber && _ocrPaneImageOnly == true;

    internal void SetPaneOcrPage(object token, int pageNumber, bool? isImageOnly) {
        if (ReferenceEquals(_ocrPaneOwner, token) && _ocrPanePageNumber == pageNumber && _ocrPaneImageOnly == isImageOnly) return;
        _ocrPaneOwner = token;
        _ocrPanePageNumber = pageNumber;
        _ocrPaneImageOnly = isImageOnly;
        OnPropertyChanged(nameof(ShowOcrPrompt));
    }

    internal void ClearPaneOcrPage(object token) {
        if (!ReferenceEquals(_ocrPaneOwner, token)) return;
        _ocrPaneOwner = null;
        _ocrPanePageNumber = 0;
        _ocrPaneImageOnly = null;
        OnPropertyChanged(nameof(ShowOcrPrompt));
    }

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
