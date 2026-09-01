using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Workflows;
using System.ComponentModel;

namespace OfficeIMO.Studio.Features.Shell;

public enum StudioWorkspaceMode {
    Home,
    PdfWorkspace,
    Convert,
    DocumentHealth
}

public sealed partial class MainWindowViewModel {
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsHomeMode))]
    [NotifyPropertyChangedFor(nameof(IsPdfWorkspaceMode))]
    [NotifyPropertyChangedFor(nameof(IsConversionMode))]
    [NotifyPropertyChangedFor(nameof(IsDocumentHealthMode))]
    [NotifyPropertyChangedFor(nameof(ShowPdfDocumentControls))]
    private StudioWorkspaceMode _workspaceMode;

    public ConversionWorkbenchViewModel ConversionWorkbench { get; private set; } = null!;

    public DocumentHealthViewModel DocumentHealth { get; private set; } = null!;

    public bool IsHomeMode => WorkspaceMode == StudioWorkspaceMode.Home;
    public bool IsPdfWorkspaceMode => WorkspaceMode == StudioWorkspaceMode.PdfWorkspace;
    public bool ShowPdfDocumentControls => IsPdfWorkspaceMode && HasDocument;
    public bool IsConversionMode => WorkspaceMode == StudioWorkspaceMode.Convert;
    public bool IsDocumentHealthMode => WorkspaceMode == StudioWorkspaceMode.DocumentHealth;

    [RelayCommand]
    private void ShowHome() => WorkspaceMode = StudioWorkspaceMode.Home;

    [RelayCommand]
    private void ShowPdfWorkspace() => WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;

    [RelayCommand]
    private void ShowConversionWorkbench() => WorkspaceMode = StudioWorkspaceMode.Convert;

    [RelayCommand]
    private void ShowDocumentHealth() => WorkspaceMode = StudioWorkspaceMode.DocumentHealth;

    private void OnWorkflowPropertyChanged(object? sender, PropertyChangedEventArgs e) {
        if (e.PropertyName == nameof(ConversionWorkbenchViewModel.IsBusy) ||
            e.PropertyName == nameof(DocumentHealthViewModel.IsBusy)) {
            OnPropertyChanged(nameof(CanCancelOperation));
        }
    }
}
