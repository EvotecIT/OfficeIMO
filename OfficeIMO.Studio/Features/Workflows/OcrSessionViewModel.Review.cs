using Avalonia.Threading;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class OcrSessionViewModel {
    // Background completions belong to this application, even after its dispatcher shuts down.
    private readonly Avalonia.Threading.Dispatcher _uiDispatcher = Avalonia.Threading.Dispatcher.UIThread;
    private async Task<IReadOnlyDictionary<PdfRecognizedWord, string>> ReviewPdfAsync(OcrSessionItem item, PdfSearchableOcrReview evidence,
        CancellationTokenSource operation, CancellationToken token) {
        var review = await _uiDispatcher.InvokeAsync(() => {
            token.ThrowIfCancellationRequested();
            SelectedItem = item;
            return PdfReview = new OcrReviewViewModel(evidence, _localizer, operation.Cancel);
        });
        try { return await review.Completion.WaitAsync(token).ConfigureAwait(false); }
        finally {
            await _uiDispatcher.InvokeAsync(() => {
                if (ReferenceEquals(PdfReview, review)) PdfReview = null;
                review.Dispose();
            });
        }
    }
    private async Task<string> ReviewImageAsync(OcrSessionItem item, ImageOcrWorkflowReview evidence,
        CancellationTokenSource operation, CancellationToken token) {
        var review = await _uiDispatcher.InvokeAsync(() => {
            token.ThrowIfCancellationRequested();
            SelectedItem = item;
            return ImageReview = new ImageOcrReviewViewModel(evidence, _localizer, operation.Cancel);
        });
        try { return await review.Completion.WaitAsync(token).ConfigureAwait(false); }
        finally {
            await _uiDispatcher.InvokeAsync(() => {
                if (ReferenceEquals(ImageReview, review)) ImageReview = null;
                review.Dispose();
            });
        }
    }
}
