using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Infrastructure.Diagnostics;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Studio.Features.Reader;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private StudioDocumentViewState _documentViewState = new();
    [ObservableProperty] private bool _isFocusReading;

    internal StudioDocumentViewState DocumentViewState => _documentViewState;

    public IReadOnlyList<StudioCommandItem> DocumentModeCommands =>
        [Commands["Read"], Commands["Comment"], Commands["Edit"], Commands["Pages"], Commands["Forms"], Commands["Protect"]];

    public StudioCommandItem SelectedDocumentModeCommand {
        get => Commands[DocumentMode switch {
            StudioDocumentMode.Annotate => "Comment", StudioDocumentMode.Edit => "Edit",
            StudioDocumentMode.Pages => "Pages", StudioDocumentMode.Forms => "Forms",
            StudioDocumentMode.Protect => "Protect", _ => "Read"
        }];
        set {
            if (value is not null && !ReferenceEquals(value, SelectedDocumentModeCommand)) value.Execute(null);
            OnPropertyChanged();
        }
    }

    [RelayCommand]
    private void ToggleFocusReading() {
        if (!HasDocument) return;
        IsFocusReading = !IsFocusReading;
        SaveDocumentViewState();
    }

    partial void OnIsFocusReadingChanged(bool value) {
        if (value) {
            DocumentMode = StudioDocumentMode.View;
            WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;
        }
    }

    internal void UpdatePanePreferences(double navigationWidth, double inspectorWidth, StudioPanePreference? preference = null) {
        var panes = new Dictionary<StudioDocumentMode, StudioPanePreference>(_documentViewState.Panes);
        if (preference is not null) panes[DocumentMode] = preference;
        _documentViewState = (_documentViewState with {
            NavigationWidth = navigationWidth,
            InspectorWidth = inspectorWidth,
            Panes = panes
        }).Normalize();
        SaveDocumentViewState();
    }

    internal void SaveDocumentViewState() {
        if (!HasDocument || DocumentPath is not { } path) return;
        _documentViewState = (_documentViewState with {
            PageNumber = SelectedPage?.PageNumber ?? 1,
            Zoom = Zoom,
            ZoomMode = _zoomMode,
            ReaderLayout = ReaderLayout,
            FocusReading = IsFocusReading
        }).Normalize();
        if (!_persistDocumentViews) return;
        try {
            _services.DocumentViews.Put(path, _documentViewState);
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            _services.Diagnostics.Write(StudioDiagnosticLevel.Warning, "Preferences", "DocumentViewSaveFailed", exception);
        }
    }

    private void RestoreDocumentViewState() {
        _documentViewState = _persistDocumentViews && DocumentPath is { } path ? _services.DocumentViews.Get(path) : new();
        SelectedReaderLayoutChoice = ReaderLayoutChoices.Single(choice => choice.Mode == _documentViewState.ReaderLayout);
        SelectedPage = Pages.Count == 0 ? null : Pages[Math.Clamp(_documentViewState.PageNumber - 1, 0, Pages.Count - 1)];
        _zoomMode = _documentViewState.ZoomMode;
        if (_zoomMode == ViewerZoomMode.Custom) ApplyZoom(_documentViewState.Zoom);
        else ApplyFitZoom();
        IsFocusReading = _documentViewState.FocusReading;
        OnPropertyChanged(nameof(DocumentViewState));
    }
}
