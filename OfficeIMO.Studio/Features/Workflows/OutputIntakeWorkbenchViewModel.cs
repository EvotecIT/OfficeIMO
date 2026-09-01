using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace OfficeIMO.Studio.Features.Workflows;

public enum OutputWorkbenchSection {
    PrintPreview,
    ExportPages,
    AssemblePdf
}

public sealed partial class OutputIntakeWorkbenchViewModel : ObservableObject, IDisposable {
    private string? _activeDocumentPath;

    public OutputIntakeWorkbenchViewModel(
        Func<CancellationToken, Task<string?>> pickPdf,
        Func<CancellationToken, Task<string?>> pickOutputFolder,
        Func<CancellationToken, Task<IReadOnlyList<string>>> pickAssemblyFiles,
        Func<CancellationToken, Task<string?>> pickAssemblyFolder,
        Func<CancellationToken, Task<string?>> pickOutputPdf) {
        PrintPreview = new PrintPreviewViewModel(pickPdf);
        PageExport = new PageImageExportViewModel(pickPdf, pickOutputFolder);
        Assembly = new PdfAssemblyViewModel(pickAssemblyFiles, pickAssemblyFolder, pickOutputPdf);
        PrintPreview.PropertyChanged += OnChildPropertyChanged;
        PageExport.PropertyChanged += OnChildPropertyChanged;
        Assembly.PropertyChanged += OnChildPropertyChanged;
    }

    public PrintPreviewViewModel PrintPreview { get; }
    public PageImageExportViewModel PageExport { get; }
    public PdfAssemblyViewModel Assembly { get; }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsPrintPreview))]
    [NotifyPropertyChangedFor(nameof(IsPageExport))]
    [NotifyPropertyChangedFor(nameof(IsAssembly))]
    private OutputWorkbenchSection _selectedSection = OutputWorkbenchSection.PrintPreview;

    public bool IsPrintPreview => SelectedSection == OutputWorkbenchSection.PrintPreview;
    public bool IsPageExport => SelectedSection == OutputWorkbenchSection.ExportPages;
    public bool IsAssembly => SelectedSection == OutputWorkbenchSection.AssemblePdf;
    public bool IsBusy => PrintPreview.IsBusy || PageExport.IsBusy || Assembly.IsBusy;
    public bool CanCancel => IsBusy;

    internal void Prepare(OutputWorkbenchSection section, string? activeDocumentPath) {
        _activeDocumentPath = activeDocumentPath;
        SelectedSection = section;
        if (section == OutputWorkbenchSection.PrintPreview) PrintPreview.UseDocument(activeDocumentPath);
        if (section == OutputWorkbenchSection.ExportPages) PageExport.UseDocument(activeDocumentPath);
        if (section == OutputWorkbenchSection.AssemblePdf) Assembly.UseDocument(activeDocumentPath);
    }

    [RelayCommand]
    private void ShowPrintPreview() {
        SelectedSection = OutputWorkbenchSection.PrintPreview;
        PrintPreview.UseDocument(_activeDocumentPath);
    }

    [RelayCommand]
    private void ShowPageExport() {
        SelectedSection = OutputWorkbenchSection.ExportPages;
        PageExport.UseDocument(_activeDocumentPath);
    }

    [RelayCommand]
    private void ShowAssembly() {
        SelectedSection = OutputWorkbenchSection.AssemblePdf;
        Assembly.UseDocument(_activeDocumentPath);
    }

    [RelayCommand]
    private void Cancel() {
        if (PrintPreview.CanCancel) PrintPreview.CancelCommand.Execute(null);
        if (PageExport.CanCancel) PageExport.CancelCommand.Execute(null);
        if (Assembly.CanCancel) Assembly.CancelCommand.Execute(null);
    }

    private void OnChildPropertyChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e) {
        if (e.PropertyName == nameof(PrintPreviewViewModel.IsBusy) ||
            e.PropertyName == nameof(PageImageExportViewModel.IsBusy) ||
            e.PropertyName == nameof(PdfAssemblyViewModel.IsBusy)) {
            OnPropertyChanged(nameof(IsBusy));
            OnPropertyChanged(nameof(CanCancel));
        }
    }

    public void Dispose() {
        PrintPreview.PropertyChanged -= OnChildPropertyChanged;
        PageExport.PropertyChanged -= OnChildPropertyChanged;
        Assembly.PropertyChanged -= OnChildPropertyChanged;
        PrintPreview.Dispose();
        PageExport.Dispose();
        Assembly.Dispose();
    }
}
