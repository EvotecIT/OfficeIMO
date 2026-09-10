using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed record PrintPaperSourceChoice(string? Id, string Label) {
    public override string ToString() => Label;
}

public sealed partial class PrintPreviewViewModel {
    private CancellationTokenSource? _paperSourceCancellation;
    [ObservableProperty] private IReadOnlyList<PrintPaperSourceChoice> _paperSourceChoices = [];
    [ObservableProperty] private PrintPaperSourceChoice? _selectedPaperSource;
    [ObservableProperty] private bool _isDiscoveringPaperSources;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasPaperSourceError))]
    private string _paperSourceError = string.Empty;
    public bool HasPaperSourceError => !string.IsNullOrEmpty(PaperSourceError);
    internal Task PaperSourceDiscovery { get; private set; } = Task.CompletedTask;

    partial void OnSelectedPrinterChanged(PdfPrinterInfo? value) => PaperSourceDiscovery = RefreshPaperSourcesAsync(value);

    private async Task RefreshPaperSourcesAsync(PdfPrinterInfo? printer) {
        _paperSourceCancellation?.Cancel();
        using var operation = new CancellationTokenSource();
        _paperSourceCancellation = operation;
        PaperSourceError = string.Empty;
        var defaultChoice = new PrintPaperSourceChoice(null, T("PaperSource.Default", "Printer default"));
        PaperSourceChoices = [defaultChoice];
        SelectedPaperSource = defaultChoice;
        IsDiscoveringPaperSources = printer is not null && !_disposed;
        try {
            if (printer is null || _disposed) return;
            var sources = await _printers.GetPaperSourcesAsync(printer.Name, operation.Token).ConfigureAwait(true);
            operation.Token.ThrowIfCancellationRequested();
            if (_disposed || !ReferenceEquals(_paperSourceCancellation, operation)) return;
            PaperSourceChoices = [defaultChoice, .. sources.Select(source => new PrintPaperSourceChoice(source.Id, source.Name))];
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) { }
        catch (Exception error) {
            if (!_disposed && ReferenceEquals(_paperSourceCancellation, operation))
                PaperSourceError = T("PaperSource.Unavailable", "Paper-source discovery failed. Printing will use the printer default.") + " " + error.Message;
        } finally {
            if (ReferenceEquals(_paperSourceCancellation, operation)) {
                _paperSourceCancellation = null;
                IsDiscoveringPaperSources = false;
            }
        }
    }
}
