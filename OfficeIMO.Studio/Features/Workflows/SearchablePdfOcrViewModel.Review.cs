using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class SearchablePdfOcrViewModel {
    // Background completions belong to this application, even after its dispatcher shuts down.
    private readonly Avalonia.Threading.Dispatcher _uiDispatcher = Avalonia.Threading.Dispatcher.UIThread;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasReview))]
    private OcrReviewViewModel? _review;

    public bool HasReview => Review is not null;

    private Task<IReadOnlyDictionary<PdfRecognizedWord, string>> ReviewWordsAsync(PdfSearchableOcrReview evidence, CancellationToken cancellationToken) =>
        ReviewWordsAsync(evidence, cancellationToken, false);

    private async Task<IReadOnlyDictionary<PdfRecognizedWord, string>> ReviewWordsAsync(PdfSearchableOcrReview evidence, CancellationToken cancellationToken, bool textOnly) {
        var model = await _uiDispatcher.InvokeAsync(() => {
            cancellationToken.ThrowIfCancellationRequested();
            var operation = _cancellation;
            var pending = new OcrReviewViewModel(evidence, _localizer, () => operation?.Cancel(), textOnly);
            Review = pending;
            Status = textOnly ? T("Text.Review", "Review the recognized words to extract.")
                : T("Review.Status", "Review the recognized words before creating the searchable PDF.");
            return pending;
        });
        try { return await model.Completion.WaitAsync(cancellationToken).ConfigureAwait(false); }
        finally {
            await _uiDispatcher.InvokeAsync(() => {
                if (ReferenceEquals(Review, model)) Review = null;
                model.Dispose();
            });
        }
    }
}
