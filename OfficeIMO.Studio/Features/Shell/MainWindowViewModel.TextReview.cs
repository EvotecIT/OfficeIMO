using System.Collections.ObjectModel;
using System.ComponentModel;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Workspace;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private long _textReviewGeneration;
    private long _textReviewRevision;
    private PdfWorkspace? _textReviewWorkspace;
    private CancellationTokenSource? _textInspectionCancellation;
    private CancellationTokenSource? _textPreviewCancellation;
    private CancellationTokenSource? _textPreparationCancellation;
    private CancellationTokenSource? _textSearchCancellation;
    private PreparedTextEdit? _preparedTextEdit;
    private bool _textReviewIsBatch;
    [ObservableProperty] private PdfTextDraftViewModel? _textEditDraft;
    [ObservableProperty] private Bitmap? _textPreviewBefore;
    [ObservableProperty] private Bitmap? _textPreviewAfter;
    [ObservableProperty] private string? _textPreviewSummary;
    [ObservableProperty] private bool _textPreviewFullPage;
    [ObservableProperty] private Avalonia.Rect? _textPreviewRegion;
    [ObservableProperty] private PdfTextReplacementMatchViewModel? _selectedReplacementMatch;
    public ObservableCollection<PdfTextReplacementMatchViewModel> ReplacementMatches { get; } = new();
    public bool HasTextPreview => _preparedTextEdit is not null && TextPreviewAfter is not null;
    public bool HasReplacementMatches => ReplacementMatches.Count > 0;

    partial void OnSelectedObjectTextChanged(string value) {
        if (!_textReviewIsBatch && TextEditDraft is { } draft) draft.Text = value;
    }
    partial void OnReplaceAllReplacementTextChanged(string value) {
        if (_textReviewIsBatch && TextEditDraft is { } draft) draft.Text = value;
    }
    partial void OnReplaceAllFindTextChanged(string value) { if (_textReviewIsBatch || _textSearchCancellation is not null) ClearTextReview(); }
    partial void OnReplaceAllMatchCaseChanged(bool value) { if (_textReviewIsBatch || _textSearchCancellation is not null) ClearTextReview(); }
    partial void OnReplaceAllWholeWordsChanged(bool value) { if (_textReviewIsBatch || _textSearchCancellation is not null) ClearTextReview(); }

    internal async Task BeginInlineTextEditAsync(PdfEditorSelection selection) {
        ClearTextReview();
        if (_workspace is not { } workspace || selection.Kind != PdfEditorSelectionKind.Text || !CanEditPageContent) return;
        _textReviewWorkspace = workspace;
        _textReviewRevision = workspace.Revision;
        var draft = CreateTextDraft(selection.Text ?? string.Empty);
        TextEditDraft = draft;
        foreach (var page in Pages) page.InlineTextDraft = page.PageNumber == selection.PageNumber ? draft : null;
        using var inspectionCancellation = new CancellationTokenSource();
        _textInspectionCancellation = inspectionCancellation;
        try {
            PdfTextMatch match = await workspace.InspectSelectedTextAsync(selection, inspectionCancellation.Token).ConfigureAwait(true);
            if (!ReferenceEquals(draft, TextEditDraft) || !ReferenceEquals(workspace, _workspace) || workspace.Revision != _textReviewRevision) return;
            draft.SourceStyle = _localizer.FormatOrDefault("TextEdit.SourceStyle", "Source: {0}, {1:0.##} pt. Replacement uses a standard PDF font; review any substitution warnings.", match.SourceFont ?? match.SuggestedFont.ToString(), match.FontSize);
            draft.IsReady = true;
        } catch (OperationCanceledException) when (inspectionCancellation.IsCancellationRequested) {
            // The selection or document changed while this inspection was queued or running.
        } catch (Exception error) {
            if (ReferenceEquals(draft, TextEditDraft)) draft.SourceStyle = error.Message;
        } finally {
            if (ReferenceEquals(_textInspectionCancellation, inspectionCancellation)) _textInspectionCancellation = null;
        }
    }

    private PdfTextDraftViewModel CreateTextDraft(string text) {
        var draft = new PdfTextDraftViewModel(text, PreviewTextEditCommand, CancelTextEditCommand, _localizer);
        draft.PropertyChanged += OnTextDraftChanged;
        return draft;
    }

    private void OnTextDraftChanged(object? sender, PropertyChangedEventArgs args) {
        if (sender is not PdfTextDraftViewModel draft || !ReferenceEquals(draft, TextEditDraft)) return;
        if (args.PropertyName is nameof(PdfTextDraftViewModel.SourceStyle) or nameof(PdfTextDraftViewModel.IsReady)) return;
        if (args.PropertyName == nameof(PdfTextDraftViewModel.Text)) {
            if (_textReviewIsBatch) ReplaceAllReplacementText = draft.Text; else SelectedObjectText = draft.Text;
        }
        InvalidateTextPreview();
    }

    private void InvalidateTextPreview() {
        _textPreparationCancellation?.Cancel();
        _textSearchCancellation?.Cancel();
        _textPreviewCancellation?.Cancel();
        _textPreviewCancellation = null;
        _textReviewGeneration++;
        _preparedTextEdit = null;
        ClearRenderedTextPreview();
        TextPreviewSummary = null;
    }

    private void ClearRenderedTextPreview() {
        TextPreviewBefore?.Dispose();
        TextPreviewAfter?.Dispose();
        TextPreviewBefore = null;
        TextPreviewAfter = null;
        OnPropertyChanged(nameof(HasTextPreview));
    }

    private void ClearTextReview() {
        _textInspectionCancellation?.Cancel();
        _textInspectionCancellation = null;
        if (_textReviewIsBatch) foreach (var page in Pages) page.ActiveSearchHighlight = null;
        InvalidateTextPreview();
        if (TextEditDraft is { } draft) draft.PropertyChanged -= OnTextDraftChanged;
        TextEditDraft = null;
        foreach (var page in Pages) page.InlineTextDraft = null;
        foreach (var match in ReplacementMatches) match.PropertyChanged -= OnReplacementMatchChanged;
        ReplacementMatches.Clear();
        SelectedReplacementMatch = null;
        _textReviewWorkspace = null;
        _textReviewIsBatch = false;
        OnPropertyChanged(nameof(HasReplacementMatches));
    }

    [RelayCommand] private void CancelTextEdit() => ClearObjectSelection();

    [RelayCommand]
    private async Task FindTextReplacementsAsync(CancellationToken token) {
        if (_workspace is not { } workspace || !CanEditPageContent || string.IsNullOrEmpty(ReplaceAllFindText)) return;
        ClearObjectSelection();
        string find = ReplaceAllFindText;
        bool matchCase = ReplaceAllMatchCase, wholeWords = ReplaceAllWholeWords;
        long revision = workspace.Revision, generation = _textReviewGeneration;
        IReadOnlyList<PdfTextMatch>? matches = null;
        using var searchOperation = CancellationTokenSource.CreateLinkedTokenSource(token);
        _textSearchCancellation = searchOperation;
        bool success;
        try {
            success = await RunStandaloneAsync(async cancellation => {
                matches = await workspace.FindReplacementMatchesAsync(find, matchCase, wholeWords, cancellation).ConfigureAwait(true);
            }, searchOperation.Token).ConfigureAwait(true);
        } finally {
            if (ReferenceEquals(_textSearchCancellation, searchOperation)) _textSearchCancellation = null;
        }
        if (!success || matches is null || generation != _textReviewGeneration || !ReferenceEquals(_workspace, workspace) || revision != workspace.Revision ||
            find != ReplaceAllFindText || matchCase != ReplaceAllMatchCase || wholeWords != ReplaceAllWholeWords) return;
        if (matches.Count > 2000) { ErrorMessage = _localizer.GetOrDefault("TextEdit.TooManyMatches", "Narrow the search to at most 2,000 occurrences before reviewing replacements."); return; }
        _textReviewIsBatch = true;
        _textReviewWorkspace = workspace;
        _textReviewRevision = revision;
        TextEditDraft = CreateTextDraft(ReplaceAllReplacementText);
        TextEditDraft.SourceStyle = _localizer.GetOrDefault("TextEdit.BatchStyle", "Each selected occurrence keeps its detected size unless overridden. Review font substitutions and nearby text flow in the preview.");
        TextEditDraft.IsReady = matches.Count > 0;
        for (int i = 0; i < matches.Count; i++) {
            var item = new PdfTextReplacementMatchViewModel(i, matches[i]);
            item.PropertyChanged += OnReplacementMatchChanged;
            ReplacementMatches.Add(item);
        }
        SelectedReplacementMatch = ReplacementMatches.FirstOrDefault();
        OnPropertyChanged(nameof(HasReplacementMatches));
    }

    private void OnReplacementMatchChanged(object? sender, PropertyChangedEventArgs args) {
        if (args.PropertyName == nameof(PdfTextReplacementMatchViewModel.IsIncluded)) InvalidateTextPreview();
    }

    partial void OnSelectedReplacementMatchChanged(PdfTextReplacementMatchViewModel? value) {
        _textPreviewCancellation?.Cancel();
        ClearRenderedTextPreview();
        if (value is null) return;
        var bounds = value.Match.VisualBounds;
        foreach (var page in Pages) page.ActiveSearchHighlight = page.PageNumber == value.PageNumber
            ? new Avalonia.Rect(bounds.Left, bounds.Top, bounds.Width, bounds.Height) : null;
        NavigateToPage(value.PageNumber);
        UpdateTextPreviewRegion(value.PageNumber);
        if (_preparedTextEdit is not null) _ = ShowTextPreviewPageAsync(value.PageNumber);
    }

    partial void OnTextPreviewFullPageChanged(bool value) => UpdateTextPreviewRegion(SelectedReplacementMatch?.PageNumber ?? SelectedObject?.PageNumber ?? 1);

    private void UpdateTextPreviewRegion(int pageNumber) {
        TextPreviewRegion = null;
        if (TextPreviewFullPage || Pages.FirstOrDefault(page => page.PageNumber == pageNumber) is not { } page) return;
        Avalonia.Rect bounds;
        if (_textReviewIsBatch && SelectedReplacementMatch is { } match && match.PageNumber == pageNumber) {
            var area = match.Match.VisualBounds;
            bounds = new(area.Left, area.Top, area.Width, area.Height);
        } else if (SelectedObject is { Kind: PdfEditorSelectionKind.Text } selection && selection.PageNumber == pageNumber) {
            bounds = new(selection.Bounds.Left, selection.Bounds.Top, selection.Bounds.Width, selection.Bounds.Height);
        } else return;
        double width = page.VisualPageWidth, height = page.VisualPageHeight;
        double cropWidth = Math.Min(width, Math.Max(240, bounds.Width + 120));
        double cropHeight = Math.Min(height, Math.Max(100, bounds.Height + 70));
        double left = Math.Clamp(bounds.Left - 30, 0, Math.Max(0, width - cropWidth));
        double top = Math.Clamp(bounds.Top - 30, 0, Math.Max(0, height - cropHeight));
        TextPreviewRegion = new(left / width, top / height, cropWidth / width, cropHeight / height);
    }

    [RelayCommand]
    private async Task PreviewTextEditAsync(CancellationToken token) {
        if (_workspace is not { } workspace || TextEditDraft is not { IsReady: true } draft ||
            !ReferenceEquals(_textReviewWorkspace, workspace) || workspace.Revision != _textReviewRevision) return;
        InvalidateTextPreview();
        long generation = _textReviewGeneration;
        string replacement = draft.Text;
        PdfTextEditOptions options;
        try { options = draft.CaptureOptions(); } catch (Exception error) { ErrorMessage = error.Message; return; }
        PdfEditorSelection? selection = SelectedObject;
        var included = ReplacementMatches.Where(match => match.IsIncluded).ToArray();
        bool batch = _textReviewIsBatch;
        if (batch && included.Length == 0) return;
        string find = ReplaceAllFindText;
        bool matchCase = ReplaceAllMatchCase, wholeWords = ReplaceAllWholeWords;
        PreparedTextEdit? prepared = null;
        using var preparation = CancellationTokenSource.CreateLinkedTokenSource(token);
        _textPreparationCancellation = preparation;
        bool success;
        try {
            success = await RunStandaloneAsync(async cancellation => {
                prepared = batch
                    ? await workspace.PreviewTextReplacementsAsync(find, replacement, matchCase, wholeWords,
                        included.Select(match => match.Index).ToArray(), options, included.Select(match => match.PageNumber).ToArray(), cancellation).ConfigureAwait(true)
                    : selection is { Kind: PdfEditorSelectionKind.Text }
                        ? await workspace.PreviewSelectedTextAsync(selection, replacement, options, cancellation).ConfigureAwait(true) : null;
            }, preparation.Token).ConfigureAwait(true);
        } finally {
            if (ReferenceEquals(_textPreparationCancellation, preparation)) _textPreparationCancellation = null;
        }
        if (!success || prepared is null || generation != _textReviewGeneration || !ReferenceEquals(workspace, _workspace) || workspace.Revision != prepared.Revision) return;
        _preparedTextEdit = prepared;
        TextPreviewSummary = _localizer.FormatOrDefault("TextEdit.PreviewSummary", "Prepared {0:N0} replacement(s). Review the rendered result before applying.", prepared.AffectedCount)
            + (prepared.Warnings.Count > 0 ? " " + string.Join(" ", prepared.Warnings) : string.Empty);
        OnPropertyChanged(nameof(HasTextPreview));
        await ShowTextPreviewPageAsync(SelectedReplacementMatch?.PageNumber ?? prepared.Pages[0], token).ConfigureAwait(true);
    }

    private long _textPreviewPageGeneration;
    internal async Task ShowTextPreviewPageAsync(int page, CancellationToken token = default) {
        if (_workspace is not { } workspace || _preparedTextEdit is not { } prepared) return;
        _textPreviewCancellation?.Cancel();
        using var operation = CancellationTokenSource.CreateLinkedTokenSource(token);
        _textPreviewCancellation = operation;
        ClearRenderedTextPreview();
        UpdateTextPreviewRegion(page);
        long generation = ++_textPreviewPageGeneration;
        try {
            (byte[] before, byte[] after) = await workspace.RenderTextPreviewAsync(prepared, page, operation.Token).ConfigureAwait(true);
            operation.Token.ThrowIfCancellationRequested();
            if (!ReferenceEquals(prepared, _preparedTextEdit) || generation != _textPreviewPageGeneration) return;
            using var beforeStream = new MemoryStream(before, writable: false);
            using var afterStream = new MemoryStream(after, writable: false);
            var beforeImage = new Bitmap(beforeStream);
            Bitmap afterImage;
            try { afterImage = new Bitmap(afterStream); } catch { beforeImage.Dispose(); throw; }
            TextPreviewBefore?.Dispose(); TextPreviewAfter?.Dispose();
            TextPreviewBefore = beforeImage; TextPreviewAfter = afterImage;
            OnPropertyChanged(nameof(HasTextPreview));
        } catch (OperationCanceledException) { }
        catch (Exception error) { if (generation == _textPreviewPageGeneration && ReferenceEquals(prepared, _preparedTextEdit)) { InvalidateTextPreview(); ErrorMessage = error.Message; } }
        finally { if (ReferenceEquals(_textPreviewCancellation, operation)) _textPreviewCancellation = null; }
    }

    [RelayCommand]
    private async Task ApplyReviewedTextEditAsync(CancellationToken token) {
        if (!HasTextPreview || _workspace is not { } workspace || _preparedTextEdit is not { } prepared ||
            !ReferenceEquals(workspace, _textReviewWorkspace) || workspace.Revision != prepared.Revision) return;
        bool success = await RunMutationAsync(cancellation => workspace.ApplyPreparedTextEditAsync(prepared, cancellation, CreateProgress()), token).ConfigureAwait(true);
        if (success) ClearTextReview();
    }
}
