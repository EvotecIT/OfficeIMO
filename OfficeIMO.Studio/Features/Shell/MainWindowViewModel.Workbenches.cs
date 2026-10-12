namespace OfficeIMO.Studio.Features.Shell;

/// <summary>
/// Application workbenches (Convert, OCR, Jobs, Settings, invoices, books, output, health and provenance) belong to
/// the window, not to a document tab. The desktop window creates one document-less owner and every tab borrows its
/// workbenches, so queues, settings and results stay the same whichever document is selected.
/// </summary>
public sealed partial class MainWindowViewModel {
    private bool _ownsWorkbenches = true;

    /// <summary>True when this view model created the workbenches and answers for their work, output and lifetime.</summary>
    internal bool OwnsWorkbenches => _ownsWorkbenches;

    /// <summary>Finds an open document for a path so shared workbenches read its live state instead of the file on disk.</summary>
    internal Func<string, MainWindowViewModel?>? FindOpenDocument { get; set; }

    private void UseSharedWorkbenches(MainWindowViewModel owner) {
        if (!owner._ownsWorkbenches) throw new ArgumentException("Workbenches can only be shared from the view model that owns them.", nameof(owner));
        Jobs = owner.Jobs;
        BookWorkbench = owner.BookWorkbench;
        ConversionWorkbench = owner.ConversionWorkbench;
        InvoiceWorkbench = owner.InvoiceWorkbench;
        OutputWorkbench = owner.OutputWorkbench;
        ProvenanceWorkbench = owner.ProvenanceWorkbench;
        DocumentHealth = owner.DocumentHealth;
        OcrWorkbench = owner.OcrWorkbench;
        Settings = owner.Settings;
        OcrSession = owner.OcrSession;
        OcrReadiness = owner.OcrReadiness;
    }

    private void CancelWorkbenchOperations() {
        if (ConversionWorkbench.CanCancel) ConversionWorkbench.CancelCommand.Execute(null);
        if (ConversionWorkbench.BatchExport.IsBusy) ConversionWorkbench.BatchExport.CancelCommand.Execute(null);
        if (InvoiceWorkbench.CanCancel) InvoiceWorkbench.CancelCommand.Execute(null);
        if (BookWorkbench.IsBusy) BookWorkbench.CancelCommand.Execute(null);
        if (OutputWorkbench.CanCancel) OutputWorkbench.CancelCommand.Execute(null);
        if (ProvenanceWorkbench.CanCancel) ProvenanceWorkbench.CancelCommand.Execute(null);
        if (DocumentHealth.CanCancel) DocumentHealth.CancelCommand.Execute(null);
        if (OcrWorkbench.CanCancel) OcrWorkbench.CancelCommand.Execute(null);
        if (OcrSession.IsBusy) OcrSession.CancelCommand.Execute(null);
    }

    private void DisposeWorkbenches() {
        ConversionWorkbench.PropertyChanged -= OnWorkflowPropertyChanged;
        ConversionWorkbench.BatchExport.PropertyChanged -= OnWorkflowPropertyChanged;
        InvoiceWorkbench.PropertyChanged -= OnWorkflowPropertyChanged;
        BookWorkbench.PropertyChanged -= OnWorkflowPropertyChanged;
        OutputWorkbench.PropertyChanged -= OnWorkflowPropertyChanged;
        DocumentHealth.PropertyChanged -= OnWorkflowPropertyChanged;
        ProvenanceWorkbench.PropertyChanged -= OnWorkflowPropertyChanged;
        OcrWorkbench.PropertyChanged -= OnWorkflowPropertyChanged;
        OcrSession.PropertyChanged -= OnWorkflowPropertyChanged;
        ConversionWorkbench.Dispose();
        InvoiceWorkbench.Dispose();
        BookWorkbench.Dispose();
        Jobs.Dispose();
        OutputWorkbench.Dispose();
        DocumentHealth.Dispose();
        ProvenanceWorkbench.Dispose();
        OcrWorkbench.Dispose();
        OcrSession.Dispose();
        OcrReadiness.Dispose();
        Settings.Dispose();
    }
}
