using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Pdf.Ocr;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed partial class SearchablePdfOcrViewModel {
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasReview))]
    private OcrReviewViewModel? _review;

    public bool HasReview => Review is not null;

    private Task<IReadOnlyList<PdfRecognizedWord>> ReviewWordsAsync(PdfSearchableOcrReview evidence, CancellationToken cancellationToken) =>
        ReviewWordsAsync(evidence, cancellationToken, false);

    private async Task<IReadOnlyList<PdfRecognizedWord>> ReviewWordsAsync(PdfSearchableOcrReview evidence, CancellationToken cancellationToken, bool textOnly) {
        var model = await Dispatcher.UIThread.InvokeAsync(() => {
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
            await Dispatcher.UIThread.InvokeAsync(() => {
                if (ReferenceEquals(Review, model)) Review = null;
                model.Dispose();
            });
        }
    }
}
