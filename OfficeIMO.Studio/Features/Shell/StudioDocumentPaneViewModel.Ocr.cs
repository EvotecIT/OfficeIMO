using System.ComponentModel;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class StudioDocumentPaneViewModel {
    private readonly object _ocrPresentationToken = new();

    partial void OnSelectedPageChanged(PdfPageViewModel? oldValue, PdfPageViewModel? newValue) {
        if (oldValue is not null) oldValue.PropertyChanged -= OnSelectedSceneChanged;
        if (newValue is not null) newValue.PropertyChanged += OnSelectedSceneChanged;
        RefreshOcrPage();
    }

    private void OnSelectedSceneChanged(object? sender, PropertyChangedEventArgs args) {
        if (args.PropertyName == nameof(PdfPageViewModel.Scene)) RefreshOcrPage();
    }

    private void RefreshOcrPage() {
        if (_disposed || !IsActive) {
            Document.ClearPaneOcrPage(_ocrPresentationToken);
            return;
        }
        Document.SetPaneOcrPage(_ocrPresentationToken, SelectedPage?.PageNumber ?? 0,
            SelectedPage?.Scene is { } scene ? scene.IsImageOnly : null);
    }
}
