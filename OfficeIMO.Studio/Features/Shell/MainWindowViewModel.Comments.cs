using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Comments;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindowViewModel {
    private IReadOnlyList<CommentThreadViewModel> _allCommentThreads = [];
    private readonly List<(CommentThreadViewModel Thread, string Text)> _commentDrafts = [];
    private bool _refreshingComments;

    [ObservableProperty] private IReadOnlyList<CommentThreadViewModel> _commentThreads = [];
    [ObservableProperty] private CommentThreadViewModel? _selectedCommentThread;
    [ObservableProperty] private string _commentAuthorFilter = string.Empty;
    [ObservableProperty] private CommentStatusChoice? _selectedCommentStatus;
    [ObservableProperty] private string _commentReplyText = string.Empty;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasCommentCatalogError))] private string? _commentCatalogError;
    [ObservableProperty, NotifyPropertyChangedFor(nameof(HasUnassignedCommentDrafts))] private IReadOnlyList<CommentDraftViewModel> _unassignedCommentDrafts = [];
    [ObservableProperty] private CommentDraftViewModel? _selectedUnassignedCommentDraft;

    public IReadOnlyList<CommentStatusChoice> CommentStatuses => [
        new("all", UiText("Comments.All")), new("open", UiText("Comments.Unresolved")), new("resolved", UiText("Comments.Resolved"))];
    public bool HasCommentThread => SelectedCommentThread is not null;
    public bool HasUnassignedCommentDrafts => UnassignedCommentDrafts.Count > 0;
    public bool CanRestoreCommentDraft => HasCommentThread && SelectedUnassignedCommentDraft is not null && string.IsNullOrEmpty(CommentReplyText);
    public bool HasCommentCatalogError => !string.IsNullOrEmpty(CommentCatalogError);
    public bool HasNoCommentThreads => CommentThreads.Count == 0;
    public bool CanReviewComment => !IsOpening && !IsWorkspaceBusy && SelectedCommentThread?.Annotation.ObjectNumber is not null && CanEditAnnotations;
    public bool CanReplyToComment => CanReviewComment && !string.IsNullOrWhiteSpace(CommentReplyText);
    public bool CanResolveComment => CanReviewComment && SelectedCommentThread?.IsResolved == false;
    public bool CanReopenComment => CanReviewComment && SelectedCommentThread?.IsResolved == true;

    partial void OnSelectedCommentThreadChanging(CommentThreadViewModel? value) {
        if (_refreshingComments || SelectedCommentThread is null) return;
        _commentDrafts.RemoveAll(draft => draft.Thread.Matches(SelectedCommentThread));
        if (!string.IsNullOrEmpty(CommentReplyText)) _commentDrafts.Add((SelectedCommentThread, CommentReplyText));
    }

    partial void OnSelectedCommentThreadChanged(CommentThreadViewModel? value) {
        if (!_refreshingComments) {
            var drafts = _commentDrafts.Where(draft => value?.Matches(draft.Thread) == true).ToArray();
            CommentReplyText = drafts.Length == 1 ? drafts[0].Text : string.Empty;
            if (value?.Annotation.PageNumber is int page) NavigateToPage(page);
        }
        NotifyCommentActions();
        UpdateCommentAnchor();
        RefreshUnassignedCommentDrafts();
    }

    partial void OnCommentReplyTextChanged(string value) => NotifyCommentActions();
    partial void OnCommentAuthorFilterChanged(string value) => FilterCommentThreads();
    partial void OnSelectedCommentStatusChanged(CommentStatusChoice? value) => FilterCommentThreads();
    partial void OnSelectedUnassignedCommentDraftChanged(CommentDraftViewModel? value) => NotifyCommentActions();

    private void RefreshCommentThreads(bool documentTransition) {
        if (documentTransition) {
            SelectedCommentThread = null;
            _commentDrafts.Clear();
            CommentReplyText = string.Empty;
            CommentAuthorFilter = string.Empty;
            SelectedCommentStatus = null;
        }
        CommentCatalogError = null;
        try {
            _allCommentThreads = _workspace is null ? [] : PdfAnnotationReviewCatalog.Build(_workspace.DocumentInfo.Annotations)
                .Threads.Where(thread => thread.Root.Annotation.Subtype is not ("Popup" or "Link" or "Widget"))
                .Select(thread => new CommentThreadViewModel(thread,
                    thread.Root.Annotation.ObjectNumber is int number ? _workspace.GetAnnotationIdentity(number) : Guid.NewGuid(), _localizer)).ToArray();
        } catch (InvalidOperationException ex) {
            _allCommentThreads = [];
            CommentCatalogError = ex.Message;
        }
        FilterCommentThreads();
    }

    private void FilterCommentThreads() {
        var old = SelectedCommentThread;
        var filtered = _allCommentThreads.Where(thread =>
            (string.IsNullOrWhiteSpace(CommentAuthorFilter) || thread.Entries.Any(entry => entry.Author.Contains(CommentAuthorFilter.Trim(), StringComparison.CurrentCultureIgnoreCase))) &&
            (SelectedCommentStatus?.Id switch { "open" => !thread.IsResolved, "resolved" => thread.IsResolved, _ => true })).ToArray();
        var matches = old is null ? [] : filtered.Where(thread => thread.Matches(old)).ToArray();
        var selected = matches.Length == 1 ? matches[0] : null;
        // Swap the source and selection together: a bound selector can transiently clear its selection.
        _refreshingComments = true;
        try { CommentThreads = filtered; SelectedCommentThread = selected; }
        finally { _refreshingComments = false; }
        if (selected is null && old is not null) {
            _commentDrafts.RemoveAll(draft => draft.Thread.Matches(old));
            if (!string.IsNullOrEmpty(CommentReplyText)) _commentDrafts.Add((old, CommentReplyText));
            CommentReplyText = string.Empty;
        }
        OnPropertyChanged(nameof(HasNoCommentThreads));
        NotifyCommentActions();
        UpdateCommentAnchor();
        RefreshUnassignedCommentDrafts();
    }

    private void RefreshUnassignedCommentDrafts() {
        var identities = _allCommentThreads.Select(thread => thread.Identity).ToHashSet();
        var selected = SelectedUnassignedCommentDraft;
        UnassignedCommentDrafts = _commentDrafts.Where(draft => !identities.Contains(draft.Thread.Identity))
            .Select(draft => new CommentDraftViewModel(draft.Thread.Identity, draft.Thread.Label, draft.Text)).ToArray();
        SelectedUnassignedCommentDraft = UnassignedCommentDrafts.FirstOrDefault(draft => draft.Identity == selected?.Identity) ?? UnassignedCommentDrafts.FirstOrDefault();
    }

    [RelayCommand(CanExecute = nameof(CanRestoreCommentDraft))]
    private void RestoreCommentDraft() {
        if (!CanRestoreCommentDraft || SelectedUnassignedCommentDraft is not { } draft) return;
        CommentReplyText = draft.Text;
        _commentDrafts.RemoveAll(item => item.Thread.Identity == draft.Identity);
        RefreshUnassignedCommentDrafts();
    }

    private void UpdateCommentAnchor() {
        foreach (var page in Pages) page.CommentAnchorObjectNumber = page.PageNumber == SelectedCommentThread?.Annotation.PageNumber
            ? SelectedCommentThread.Annotation.ObjectNumber : null;
    }

    private void NotifyCommentActions() {
        OnPropertyChanged(nameof(HasCommentThread));
        OnPropertyChanged(nameof(CanReviewComment));
        OnPropertyChanged(nameof(CanReplyToComment));
        OnPropertyChanged(nameof(CanResolveComment));
        OnPropertyChanged(nameof(CanReopenComment));
        OnPropertyChanged(nameof(CanRestoreCommentDraft));
        ReplyToCommentCommand.NotifyCanExecuteChanged();
        ResolveCommentCommand.NotifyCanExecuteChanged();
        ReopenCommentCommand.NotifyCanExecuteChanged();
        RestoreCommentDraftCommand.NotifyCanExecuteChanged();
    }

    [RelayCommand]
    private void NextUnresolvedComment() {
        if (CommentThreads.Count == 0) return;
        int start = SelectedCommentThread is null ? -1 : CommentThreads.ToList().IndexOf(SelectedCommentThread);
        for (int offset = 1; offset <= CommentThreads.Count; offset++) {
            var thread = CommentThreads[(start + offset) % CommentThreads.Count];
            if (!thread.IsResolved) {
                SelectedCommentThread = thread;
                // Commands must navigate even when the selector already holds this instance.
                if (thread.Annotation.PageNumber is int page) {
                    NavigateToPage(page);
                    if (SelectedPage is not null) SelectedPage.CommentAnchorObjectNumber = null;
                    UpdateCommentAnchor();
                }
                return;
            }
        }
    }

    [RelayCommand(CanExecute = nameof(CanReplyToComment))]
    private async Task ReplyToCommentAsync(CancellationToken cancellationToken) {
        if (!CanReplyToComment || _workspace is null || SelectedCommentThread?.Annotation.ObjectNumber is not int number) return;
        var thread = SelectedCommentThread;
        string reply = CommentReplyText;
        if (await RunMutationAsync(token => _workspace.AddAnnotationReplyAsync(number, reply, EditorAuthor,
            ParseColor(EditorColorHex), token, CreateProgress()), cancellationToken).ConfigureAwait(true)) {
            _commentDrafts.RemoveAll(draft => draft.Thread.Matches(thread) && draft.Text == reply);
            if (SelectedCommentThread?.Matches(thread) == true && CommentReplyText == reply) CommentReplyText = string.Empty;
            RefreshUnassignedCommentDrafts();
        }
    }

    [RelayCommand(CanExecute = nameof(CanResolveComment))]
    private Task ResolveCommentAsync(CancellationToken cancellationToken) => SetCommentStateAsync(PdfAnnotationReviewState.Completed, cancellationToken);

    [RelayCommand(CanExecute = nameof(CanReopenComment))]
    private Task ReopenCommentAsync(CancellationToken cancellationToken) => SetCommentStateAsync(PdfAnnotationReviewState.None, cancellationToken);

    private async Task SetCommentStateAsync(PdfAnnotationReviewState state, CancellationToken cancellationToken) {
        if (!CanReviewComment || _workspace is null || SelectedCommentThread?.Annotation.ObjectNumber is not int number) return;
        await RunMutationAsync(token => _workspace.SetAnnotationReviewStateAsync(number, state, token, CreateProgress()), cancellationToken).ConfigureAwait(true);
    }
}
