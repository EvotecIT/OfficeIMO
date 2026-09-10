using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Ocr.Tesseract;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class SearchablePdfOcrViewModel {
    private readonly IScanTextRecognitionService _textRecognition;
    private bool _isExtractingText;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasExtractedText))]
    private string _extractedText = string.Empty;

    public bool HasExtractedText => !string.IsNullOrEmpty(ExtractedText);
    private bool CanExtractText => !_disposed && !IsBusy && !Scan.IsBusy && !string.IsNullOrWhiteSpace(InputPath)
        && Languages.Any(choice => choice.IsSelected);

    [RelayCommand(CanExecute = nameof(CanExtractText))]
    private async Task ExtractTextAsync() {
        if (!CanExtractText) return;
        using var operation = new CancellationTokenSource();
        _cancellation = operation;
        _isExtractingText = true;
        IsBusy = true;
        ExtractedText = string.Empty;
        ErrorMessage = null;
        Status = T("Text.Preparing", "Recognizing text on the selected scan page…");
        try {
            var languages = Languages.Where(choice => choice.IsSelected)
                .Aggregate((TesseractOcrLanguage)0, (current, choice) => current | choice.Value);
            var settings = Scan.ApplyTo(new PdfOcrMergeOptions {
                ReadOptions = new PdfReadOptions { PageSelection = PdfPageSelection.From(Scan.PageNumber) },
                Dpi = RenderDpi, MinimumConfidence = MinimumConfidencePercent / 100D
            });
            var options = new SearchablePdfOcrOptions(languages, ProvisionMissingLanguageData,
                OfficeConversionFileConflictPolicy.FailIfExists, settings);
            byte[] source = await ReadScanSourceAsync(operation.Token).ConfigureAwait(true);
            using IDisposable? execution = _jobHistory == null ? null : await _jobHistory.EnterAsync(operation.Token).ConfigureAwait(true);
            var evidence = await _textRecognition.PrepareAsync(source, options, operation.Token).ConfigureAwait(true);
            operation.Token.ThrowIfCancellationRequested();
            var selected = await ReviewWordsAsync(evidence, operation.Token, textOnly: true).ConfigureAwait(true);
            ExtractedText = evidence.ExtractText(selected, operation.Token);
            Status = HasExtractedText ? T("Text.Ready", "Reviewed text is ready to copy.") : T("Text.Empty", "No eligible words were selected.");
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) {
            Status = T("Status.Cancelled", "OCR cancelled");
        } catch (Exception error) {
            Status = T("Status.Failed", "OCR could not finish");
            ErrorMessage = error.Message;
        } finally {
            Review?.Dispose(); Review = null;
            if (ReferenceEquals(_cancellation, operation)) _cancellation = null;
            _isExtractingText = false;
            IsBusy = false;
        }
    }
}
