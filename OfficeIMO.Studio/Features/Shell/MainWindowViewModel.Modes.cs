using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Workflows;
using OfficeIMO.Workflows;
using System.ComponentModel;

namespace OfficeIMO.Studio.Features.Shell;

public enum StudioWorkspaceMode {
    Home,
    PdfWorkspace,
    Tools,
    Convert,
    DocumentHealth
}

public enum StudioDocumentMode {
    View,
    Annotate,
    Edit,
    Pages,
    Forms,
    Protect
}

public sealed partial class MainWindowViewModel {
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsHomeMode))]
    [NotifyPropertyChangedFor(nameof(IsPdfWorkspaceMode))]
    [NotifyPropertyChangedFor(nameof(IsToolsMode))]
    [NotifyPropertyChangedFor(nameof(IsConversionMode))]
    [NotifyPropertyChangedFor(nameof(IsDocumentHealthMode))]
    [NotifyPropertyChangedFor(nameof(ShowPdfDocumentControls))]
    private StudioWorkspaceMode _workspaceMode;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsViewDocumentMode))]
    [NotifyPropertyChangedFor(nameof(IsAnnotateDocumentMode))]
    [NotifyPropertyChangedFor(nameof(IsEditDocumentMode))]
    [NotifyPropertyChangedFor(nameof(IsPagesDocumentMode))]
    [NotifyPropertyChangedFor(nameof(IsFormsDocumentMode))]
    [NotifyPropertyChangedFor(nameof(IsProtectDocumentMode))]
    private StudioDocumentMode _documentMode = StudioDocumentMode.View;

    public ConversionWorkbenchViewModel ConversionWorkbench { get; private set; } = null!;

    public DocumentHealthViewModel DocumentHealth { get; private set; } = null!;

    public bool IsHomeMode => WorkspaceMode == StudioWorkspaceMode.Home;
    public bool IsPdfWorkspaceMode => WorkspaceMode == StudioWorkspaceMode.PdfWorkspace;
    public bool IsToolsMode => WorkspaceMode == StudioWorkspaceMode.Tools;
    public bool ShowPdfDocumentControls => IsPdfWorkspaceMode && HasDocument;
    public bool IsConversionMode => WorkspaceMode == StudioWorkspaceMode.Convert;
    public bool IsDocumentHealthMode => WorkspaceMode == StudioWorkspaceMode.DocumentHealth;
    public bool IsJobsMode => IsConversionMode;
    public bool IsViewDocumentMode => DocumentMode == StudioDocumentMode.View;
    public bool IsAnnotateDocumentMode => DocumentMode == StudioDocumentMode.Annotate;
    public bool IsEditDocumentMode => DocumentMode == StudioDocumentMode.Edit;
    public bool IsPagesDocumentMode => DocumentMode == StudioDocumentMode.Pages;
    public bool IsFormsDocumentMode => DocumentMode == StudioDocumentMode.Forms;
    public bool IsProtectDocumentMode => DocumentMode == StudioDocumentMode.Protect;

    partial void OnDocumentModeChanged(StudioDocumentMode value) {
        if (value is not StudioDocumentMode.Annotate and not StudioDocumentMode.Edit &&
            SelectedEditorToolChoice.Tool != PdfEditorTool.Select) {
            SelectedEditorToolChoice = EditorTools[0];
        }
    }

    [RelayCommand]
    private void ShowHome() => WorkspaceMode = StudioWorkspaceMode.Home;

    [RelayCommand]
    private void ShowPdfWorkspace() => WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;

    [RelayCommand]
    private void ShowTools() => WorkspaceMode = StudioWorkspaceMode.Tools;

    [RelayCommand]
    private void ShowJobs() => WorkspaceMode = StudioWorkspaceMode.Convert;

    [RelayCommand]
    private void ShowConversionWorkbench() => WorkspaceMode = StudioWorkspaceMode.Convert;

    [RelayCommand]
    private void ShowDocumentHealth() {
        DocumentHealth.PrepareRepairWorkflow();
        WorkspaceMode = StudioWorkspaceMode.DocumentHealth;
    }

    [RelayCommand]
    private void ShowInspect() {
        DocumentHealth.PrepareWorkflow(OfficeWorkflowOperation.Inspect);
        WorkspaceMode = StudioWorkspaceMode.DocumentHealth;
    }

    [RelayCommand]
    private void ShowOptimize() {
        DocumentHealth.PrepareWorkflow(OfficeWorkflowOperation.Optimize);
        WorkspaceMode = StudioWorkspaceMode.DocumentHealth;
    }

    [RelayCommand]
    private void ShowCompare() {
        DocumentHealth.PrepareWorkflow(OfficeWorkflowOperation.Compare);
        WorkspaceMode = StudioWorkspaceMode.DocumentHealth;
    }

    [RelayCommand]
    private void ShowRepairPlan() {
        DocumentHealth.PrepareWorkflow(OfficeWorkflowOperation.RepairPlan);
        WorkspaceMode = StudioWorkspaceMode.DocumentHealth;
    }

    [RelayCommand]
    private void ShowSanitize() {
        DocumentHealth.PrepareWorkflow(OfficeWorkflowOperation.Sanitize);
        WorkspaceMode = StudioWorkspaceMode.DocumentHealth;
    }

    [RelayCommand]
    private void ShowProtect() {
        DocumentMode = StudioDocumentMode.Protect;
        WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;
    }

    [RelayCommand]
    private void ShowPageTools() {
        DocumentMode = StudioDocumentMode.Pages;
        WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;
    }

    [RelayCommand]
    private void ShowViewMode() => DocumentMode = StudioDocumentMode.View;

    [RelayCommand]
    private void ShowAnnotateMode() => DocumentMode = StudioDocumentMode.Annotate;

    [RelayCommand]
    private void ShowEditMode() => DocumentMode = StudioDocumentMode.Edit;

    [RelayCommand]
    private void ShowPagesMode() => DocumentMode = StudioDocumentMode.Pages;

    [RelayCommand]
    private void ShowFormsMode() => DocumentMode = StudioDocumentMode.Forms;

    [RelayCommand]
    private void ShowProtectMode() => DocumentMode = StudioDocumentMode.Protect;

    [RelayCommand]
    private void BeginRedaction() {
        DocumentMode = StudioDocumentMode.Protect;
        SelectedEditorToolChoice = EditorTools.Single(choice => choice.Tool == PdfEditorTool.Redact);
    }

    [RelayCommand]
    private void BeginSignatureAppearance() {
        DocumentMode = StudioDocumentMode.Edit;
        SelectedEditorToolChoice = EditorTools.Single(choice => choice.Tool == PdfEditorTool.SignatureAppearance);
    }

    private void OnWorkflowPropertyChanged(object? sender, PropertyChangedEventArgs e) {
        if (e.PropertyName == nameof(ConversionWorkbenchViewModel.IsBusy) ||
            e.PropertyName == nameof(DocumentHealthViewModel.IsBusy)) {
            OnPropertyChanged(nameof(CanCancelOperation));
        }
    }
}
