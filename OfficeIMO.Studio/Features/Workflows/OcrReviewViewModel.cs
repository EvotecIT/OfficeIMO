using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf.Ocr;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Workflows;

public sealed record OcrReviewPageChoice(int Number, string Label);

public sealed partial class OcrReviewWord : ObservableObject {
    private readonly Action<OcrReviewWord> _changed;
    internal OcrReviewWord(PdfOcrWordEvidence evidence, bool included, string reason, Action<OcrReviewWord> changed) {
        Evidence = evidence; _isIncluded = included; Reason = reason; _changed = changed;
    }
    internal PdfOcrWordEvidence Evidence { get; }
    public string Text => Evidence.Word.Text;
    public bool IsEligible => Evidence.Disposition == PdfOcrWordDisposition.Accepted;
    public string Reason { get; }
    public string Confidence => Evidence.Word.Confidence.ToString("P0");
    public string Geometry => $"X {Evidence.Word.X:0.#}, Y {Evidence.Word.Y:0.#}, {Evidence.Word.Width:0.#} × {Evidence.Word.Height:0.#} pt";
    public double Left => Evidence.Word.X;
    public double Top => Evidence.Word.Y;
    public double Width => Evidence.Word.Width;
    public double Height => Evidence.Word.Height;
    [ObservableProperty] private bool _isIncluded;
    partial void OnIsIncludedChanged(bool value) {
        if (value && !IsEligible) { IsIncluded = false; return; }
        _changed(this);
    }
}

/// <summary>Review choices over canonical, source-bound OCR evidence.</summary>
public sealed partial class OcrReviewViewModel : ObservableObject, IDisposable {
    private readonly PdfSearchableOcrReview _review;
    private readonly IStudioLocalizer _localizer;
    private readonly Action _cancel;
    private readonly HashSet<PdfRecognizedWord> _excluded = [];
    private readonly TaskCompletionSource<IReadOnlyList<PdfRecognizedWord>> _completion = new(TaskCreationOptions.RunContinuationsAsynchronously);
    private CancellationTokenSource? _previewCancellation;
    private bool _disposed;
    internal Task PreviewTask { get; private set; } = Task.CompletedTask;
    internal Task<IReadOnlyList<PdfRecognizedWord>> Completion => _completion.Task;

    internal OcrReviewViewModel(PdfSearchableOcrReview review, IStudioLocalizer localizer, Action cancel, bool textOnly = false) {
        _review = review; _localizer = localizer; _cancel = cancel;
        CommitLabel = textOnly ? localizer.GetOrDefault("Ocr.Text.UseSelected", "Use selected text")
            : localizer.GetOrDefault("SearchablePdfOcr.CreateSearchablePDF", "Create searchable PDF");
        CommitNote = textOnly ? localizer.GetOrDefault("Ocr.Text.CommitNote", "Extract the selected words without creating a PDF.")
            : localizer.GetOrDefault("Ocr.Review.CommitNote", "Only selected eligible words will be added to the searchable PDF.");
        Pages = review.Ocr.Pages.Select(page => new OcrReviewPageChoice(page.PageNumber,
            localizer.FormatOrDefault("Ocr.Review.Page", "Page {0}", page.PageNumber))).ToArray();
        SelectedPage = Pages.FirstOrDefault();
    }

    public IReadOnlyList<OcrReviewPageChoice> Pages { get; }
    public string CommitLabel { get; }
    public string CommitNote { get; }
    [ObservableProperty] private IReadOnlyList<OcrReviewWord> _words = [];
    public string Summary => _localizer.FormatOrDefault("Ocr.Review.Summary", "Selected words: {0:N0} · Reviewed pages: {1:N0}. Rejected words remain excluded.",
        _review.Ocr.AcceptedWordCount - _excluded.Count, Pages.Count);
    [ObservableProperty] private OcrReviewPageChoice? _selectedPage;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasSelectedWord))]
    private OcrReviewWord? _selectedWord;
    public bool HasSelectedWord => SelectedWord is not null;
    [ObservableProperty] private Bitmap? _preview;
    [ObservableProperty] private bool _isZoomed;
    [ObservableProperty] private double _pageWidth;
    [ObservableProperty] private double _pageHeight;
    [ObservableProperty] private string _pageDetails = string.Empty;
    [ObservableProperty] private string _diagnostics = string.Empty;
    [ObservableProperty] private string? _previewError;
    [ObservableProperty]
    [NotifyCanExecuteChangedFor(nameof(CommitCommand))]
    private bool _isLoadingPreview;
    private bool CanCommit => !_disposed && !IsLoadingPreview && Preview is not null && PreviewError is null && !_completion.Task.IsCompleted;

    partial void OnSelectedPageChanged(OcrReviewPageChoice? value) {
        if (_disposed || value is null) return;
        PreviewTask = LoadPageAsync(value);
    }

    private async Task LoadPageAsync(OcrReviewPageChoice page) {
        _previewCancellation?.Cancel();
        using var operation = new CancellationTokenSource();
        _previewCancellation = operation;
        Preview?.Dispose(); Preview = null;
        PreviewError = null; IsLoadingPreview = true;
        var evidence = _review.Ocr.Pages.Single(item => item.PageNumber == page.Number);
        (PageWidth, PageHeight) = _review.GetPageSize(page.Number);
        Words = evidence.WordEvidence.Select(item => {
            string reason = item.Disposition switch {
                PdfOcrWordDisposition.Accepted => T("Eligible", "Eligible"),
                PdfOcrWordDisposition.LowConfidence => T("LowConfidence", "Below confidence threshold; overlap not evaluated"),
                _ => T("Overlap", "Already covered by native text")
            };
            return new OcrReviewWord(item, item.Disposition == PdfOcrWordDisposition.Accepted && !_excluded.Contains(item.Word), reason, WordChanged);
        }).ToArray();
        SelectedWord = Words.FirstOrDefault();
        PageDetails = _localizer.FormatOrDefault("Ocr.Review.PageDetails", "Language: {0} · Provider: {1} · Eligible: {2:N0} · Low confidence: {3:N0} · Native overlap: {4:N0}",
            evidence.Language ?? T("Unknown", "Not reported"), evidence.Provider ?? T("Unknown", "Not reported"),
            evidence.Words.Count, evidence.RejectedLowConfidenceCount, evidence.RejectedNativeOverlapCount);
        Diagnostics = string.Join(Environment.NewLine, evidence.Diagnostics);
        try {
            var rendered = await Task.Run(() => _review.RenderPage(page.Number, operation.Token), operation.Token).ConfigureAwait(true);
            operation.Token.ThrowIfCancellationRequested();
            using var stream = new MemoryStream(rendered.Bytes ?? throw new IOException(T("NoPreview", "The OCR source page could not be rendered.")));
            var bitmap = new Bitmap(stream);
            if (_disposed || !ReferenceEquals(_previewCancellation, operation)) bitmap.Dispose();
            else {
                Preview = bitmap;
                Diagnostics = string.Join(Environment.NewLine, evidence.Diagnostics.Concat(rendered.Diagnostics));
            }
        } catch (OperationCanceledException) when (operation.IsCancellationRequested) {
        } catch (Exception error) {
            if (!_disposed && ReferenceEquals(_previewCancellation, operation)) PreviewError = error.Message;
        } finally {
            if (ReferenceEquals(_previewCancellation, operation)) {
                _previewCancellation = null; IsLoadingPreview = false;
                CommitCommand.NotifyCanExecuteChanged();
            }
        }
    }

    private void WordChanged(OcrReviewWord word) {
        SelectedWord = word;
        if (word.IsIncluded) _excluded.Remove(word.Evidence.Word);
        else if (word.IsEligible) _excluded.Add(word.Evidence.Word);
        OnPropertyChanged(nameof(Summary));
    }

    [RelayCommand(CanExecute = nameof(CanCommit))]
    private void Commit() {
        if (!CanCommit) return;
        _completion.TrySetResult(_review.Ocr.Pages.SelectMany(page => page.Words).Where(word => !_excluded.Contains(word)).ToArray());
        CommitCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand]
    private void Cancel() {
        if (!_disposed) _cancel();
    }

    [RelayCommand]
    private void IncludeEligiblePage() {
        foreach (var word in Words.Where(word => word.IsEligible)) word.IsIncluded = true;
    }

    [RelayCommand]
    private void ExcludePage() {
        foreach (var word in Words) word.IsIncluded = false;
    }

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _previewCancellation?.Cancel();
        Preview?.Dispose(); Preview = null;
        _completion.TrySetCanceled();
        CommitCommand.NotifyCanExecuteChanged();
    }

    private string T(string suffix, string fallback) => _localizer.GetOrDefault("Ocr.Review." + suffix, fallback);
}
