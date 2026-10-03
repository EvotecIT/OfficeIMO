using System.Collections.ObjectModel;
using System.Xml.Linq;
using Avalonia.Media.Imaging;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Localization;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>Book editing and publishing surface over the shared project and storage owners.</summary>
public sealed partial class BookWorkbenchViewModel : ObservableObject, IDisposable {
    private readonly Func<IStudioFileDialogs> _dialogs;
    private readonly StudioStorageAccess _storage;
    private readonly IOfficeWorkflowPublicationGuard? _guard;
    private readonly Func<Task<UnsavedChangesDecision>> _confirmUnsaved;
    private readonly IStudioLocalizer _localizer;
    private readonly Dictionary<string, string> _bodyDrafts = new(StringComparer.Ordinal);
    private readonly Dictionary<string, string> _titleDrafts = new(StringComparer.Ordinal);
    private BookProject? _project;
    private string? _sourceLocation, _projectLocation, _projectFingerprint;
    private CancellationTokenSource? _operation;
    private bool _refreshing, _styleChanged, _disposed, _hasDraftEdits;

    internal BookWorkbenchViewModel(Func<IStudioFileDialogs> dialogs, StudioStorageAccess storage,
        IOfficeWorkflowPublicationGuard? guard, Func<Task<UnsavedChangesDecision>> confirmUnsaved,
        IStudioLocalizer? localizer = null) {
        _dialogs = dialogs; _storage = storage; _guard = guard; _confirmUnsaved = confirmUnsaved;
        _localizer = localizer ?? StudioLocalization.Current;
        _status = T("Ready", "Import a manuscript, open an EPUB or book project, or start a new book.");
    }

    public ObservableCollection<BookChapterChoice> Chapters { get; } = [];
    public ObservableCollection<string> Diagnostics { get; } = [];
    public string Title => T("Title", "Publish a book");
    public string Description => T("Description", "Create reflowable EPUB books from DOCX, Markdown and HTML. Review the import, organize chapters and save an editable book project.");
    public bool CanEdit => !_disposed && !IsBusy;
    public bool HasBook => _project != null;
    public bool CanEditBook => CanEdit && HasBook;
    public bool CanExport => CanEditBook && _project!.CanExport;
    public bool CanMoveUp => CanEditBook && SelectedChapter?.Index > 0;
    public bool CanMoveDown => CanEditBook && SelectedChapter != null && SelectedChapter.Index < Chapters.Count - 1;
    public bool CanPreview => CanEditBook && SelectedChapter != null;
    public bool CanEditChapter => CanEditBook && SelectedChapter?.IsXhtml == true;
    public bool CanRemoveChapter => CanEditBook && SelectedChapter != null && Chapters.Count > 1;
    public bool CanUndo => CanEditBook && !_hasDraftEdits && _project!.CanUndo;
    public bool CanRedo => CanEditBook && !_hasDraftEdits && _project!.CanRedo;
    public bool NeedsReview => _project != null && !_project.CanExport && !_project.ImportDiagnostics.Any(d => d.LossKind == OfficeConversionLossKind.Failure);
    public bool CanAcceptLoss => CanEdit && NeedsReview;
    internal BookProject? Project => _project;

    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private bool _isDirty;
    [ObservableProperty] private string _bookTitle = string.Empty;
    [ObservableProperty] private string _language = "en";
    [ObservableProperty] private string _creator = string.Empty;
    [ObservableProperty] private string _chapterTitle = string.Empty;
    [ObservableProperty] private string _chapterBody = string.Empty;
    [ObservableProperty] private string _stylesheet = string.Empty;
    [ObservableProperty] private string _status;
    [ObservableProperty] private string _projectName = string.Empty;
    [ObservableProperty] private BookChapterChoice? _selectedChapter;
    [ObservableProperty] private Bitmap? _preview;
    [ObservableProperty] private string _previewStatus = string.Empty;

    partial void OnBookTitleChanged(string value) => MarkDraft();
    partial void OnLanguageChanged(string value) => MarkDraft();
    partial void OnCreatorChanged(string value) => MarkDraft();
    partial void OnChapterTitleChanged(string value) {
        if (!_refreshing && SelectedChapter?.IsXhtml == true) { _titleDrafts[SelectedChapter.Id] = value; MarkDraft(); }
    }
    partial void OnStylesheetChanged(string value) { if (!_refreshing) { _styleChanged = true; MarkDraft(); } }
    partial void OnChapterBodyChanged(string value) {
        if (!_refreshing && SelectedChapter != null) { _bodyDrafts[SelectedChapter.Id] = value; MarkDraft(); }
    }
    partial void OnSelectedChapterChanged(BookChapterChoice? value) {
        ClearPreview();
        bool wasRefreshing = _refreshing;
        _refreshing = true;
        try {
            ChapterTitle = value == null ? string.Empty : _titleDrafts.TryGetValue(value.Id, out string? title) ? title : value.Title;
            ChapterBody = value?.IsXhtml != true ? string.Empty : _bodyDrafts.TryGetValue(value.Id, out string? draft) ? draft :
                _project!.Publication.GetContentXml(value.Id).Root!.Element(XName.Get("body", "http://www.w3.org/1999/xhtml"))!.ToString(SaveOptions.OmitDuplicateNamespaces);
        } finally { _refreshing = wasRefreshing; }
        NotifyCommands();
    }
    partial void OnIsBusyChanged(bool value) => NotifyCommands();
    private void MarkDraft() { if (!_refreshing && _project != null) { IsDirty = true; _hasDraftEdits = true; NotifyCommands(); } }
    private string T(string key, string fallback) => _localizer.GetOrDefault("Book." + key, fallback);
    private void NotifyCommands() {
        foreach (string name in new[] { nameof(CanEdit), nameof(HasBook), nameof(CanEditBook), nameof(CanEditChapter), nameof(CanExport), nameof(CanMoveUp), nameof(CanMoveDown), nameof(CanPreview), nameof(NeedsReview), nameof(CanAcceptLoss) }) OnPropertyChanged(name);
        NewBookCommand.NotifyCanExecuteChanged(); OpenBookCommand.NotifyCanExecuteChanged(); SaveProjectCommand.NotifyCanExecuteChanged();
        ExportBookCommand.NotifyCanExecuteChanged(); ApplyEditsCommand.NotifyCanExecuteChanged(); ChooseCoverCommand.NotifyCanExecuteChanged();
        MoveUpCommand.NotifyCanExecuteChanged(); MoveDownCommand.NotifyCanExecuteChanged(); RenameChapterCommand.NotifyCanExecuteChanged();
        PreviewChapterCommand.NotifyCanExecuteChanged(); AcceptImportLossCommand.NotifyCanExecuteChanged();
        AddChapterCommand.NotifyCanExecuteChanged(); RemoveChapterCommand.NotifyCanExecuteChanged(); OnPropertyChanged(nameof(CanRemoveChapter));
        UndoEditCommand.NotifyCanExecuteChanged(); RedoEditCommand.NotifyCanExecuteChanged(); OnPropertyChanged(nameof(CanUndo)); OnPropertyChanged(nameof(CanRedo));
    }
    private void RefreshBook(int index = 0) {
        bool wasRefreshing = _refreshing;
        _refreshing = true;
        try {
            BookTitle = _project!.Publication.Title; Language = _project.Publication.Language; Creator = _project.Publication.Creator ?? string.Empty;
            Chapters.Clear();
            EpubDocument reading = _project.Publication.Read(new EpubReadOptions { MaxChapters = _project.Publication.Spine.Count });
            for (int position = 0; position < _project.Publication.Spine.Count; position++) {
                EpubSpineItem spine = _project.Publication.Spine[position];
                var item = _project.Publication.Manifest.Single(entry => entry.Id == spine.ManifestId);
                string title = reading.Chapters.FirstOrDefault(chapter => chapter.Path == item.Reference.ContainerPath)?.Title ?? item.Id;
                Chapters.Add(new(position, item.Id, title, string.Equals(item.MediaType, "application/xhtml+xml", StringComparison.OrdinalIgnoreCase)));
            }
            SelectedChapter = Chapters.FirstOrDefault(chapter => chapter.Index == index) ?? Chapters.FirstOrDefault();
            var style = _project.Publication.Manifest.FirstOrDefault(item => item.Id == "book-project-style" && item.MediaType == "text/css");
            if (!_styleChanged) Stylesheet = style != null && HtmlResourcePipeline.TryDecodeStylesheet(_project.Publication.GetResourceBytes(style.Id), "text/css", out string css) ? css : string.Empty;
            Diagnostics.Clear();
            foreach (var diagnostic in _project.ImportDiagnostics) Diagnostics.Add(diagnostic.Code + ": " + diagnostic.Message);
        } finally { _refreshing = wasRefreshing; }
        NotifyCommands();
    }
    private async Task RunAsync(Func<CancellationToken, Task> action) {
        if (!CanEdit) return;
        IsBusy = true;
        using var cancellation = new CancellationTokenSource(); _operation = cancellation;
        try { await action(cancellation.Token).ConfigureAwait(true); }
        catch (OperationCanceledException) when (cancellation.IsCancellationRequested) { Status = T("Cancelled", "Cancelled. Your current book is retained."); }
        catch (Exception error) { Status = error.Message; }
        finally { _operation = null; IsBusy = false; }
    }
    private async Task ApplyDraftsAsync(CancellationToken token) {
        var edits = new BookProjectEdits { Title = BookTitle, Language = Language, Creator = Creator,
            ChapterBodies = new Dictionary<string, string>(_bodyDrafts), ChapterTitles = new Dictionary<string, string>(_titleDrafts), Stylesheet = _styleChanged ? Stylesheet : null };
        await Task.Run(() => _project!.ApplyEdits(edits, token), token).ConfigureAwait(true);
        _bodyDrafts.Clear(); _titleDrafts.Clear(); _styleChanged = false; _hasDraftEdits = false;
        RefreshBook(SelectedChapter?.Index ?? 0);
    }

    [RelayCommand(CanExecute = nameof(CanEdit))]
    private async Task NewBookAsync() {
        if (!await PrepareCloseAsync()) return;
        await RunAsync(_ => { _project = BookProject.Create(T("Untitled", "Untitled book")); ResetLocation(); IsDirty = true; RefreshBook(); return Task.CompletedTask; });
    }
    [RelayCommand(CanExecute = nameof(CanEditBook))]
    private Task ApplyEditsAsync() => RunAsync(async token => { await ApplyDraftsAsync(token); Status = T("Applied", "Book changes validated."); });
    [RelayCommand(CanExecute = nameof(CanEditChapter))]
    private Task RenameChapterAsync() => RunAsync(async token => {
        await ApplyDraftsAsync(token);
        IsDirty = true;
    });
    [RelayCommand(CanExecute = nameof(CanMoveUp))] private Task MoveUpAsync() => MoveAsync(-1);
    [RelayCommand(CanExecute = nameof(CanMoveDown))] private Task MoveDownAsync() => MoveAsync(1);
    [RelayCommand(CanExecute = nameof(CanUndo))]
    private Task UndoEditAsync() => RunAsync(async token => { await Task.Run(() => _project!.Undo(token), token); IsDirty = true; RefreshBook(SelectedChapter?.Index ?? 0); });
    [RelayCommand(CanExecute = nameof(CanRedo))]
    private Task RedoEditAsync() => RunAsync(async token => { await Task.Run(() => _project!.Redo(token), token); IsDirty = true; RefreshBook(SelectedChapter?.Index ?? 0); });
    [RelayCommand(CanExecute = nameof(CanEditBook))]
    private Task AddChapterAsync() => RunAsync(async token => {
        await ApplyDraftsAsync(token);
        await Task.Run(() => _project!.AddChapter(T("NewChapter", "New chapter"), token), token);
        IsDirty = true; RefreshBook(_project!.Publication.Spine.Count - 1);
    });
    [RelayCommand(CanExecute = nameof(CanRemoveChapter))]
    private Task RemoveChapterAsync() => RunAsync(async token => {
        int index = SelectedChapter!.Index; await ApplyDraftsAsync(token);
        await Task.Run(() => _project!.RemoveChapter(index, token), token);
        IsDirty = true; RefreshBook(Math.Min(index, _project!.Publication.Spine.Count - 1));
    });
    private Task MoveAsync(int offset) => RunAsync(async token => {
        int index = SelectedChapter!.Index; await ApplyDraftsAsync(token);
        await Task.Run(() => _project!.MoveChapter(index, index + offset, token), token); IsDirty = true; RefreshBook(index + offset);
    });
    [RelayCommand(CanExecute = nameof(CanAcceptLoss))]
    private void AcceptImportLoss() { _project!.AcknowledgeImportLoss(); IsDirty = true; NotifyCommands(); }
    [RelayCommand] private void Cancel() => _operation?.Cancel();
    private void ClearPreview() { Preview?.Dispose(); Preview = null; PreviewStatus = string.Empty; }
    private void ResetLocation() {
        _sourceLocation = _projectLocation = _projectFingerprint = _sourceIdentity = _projectIdentity = null; ProjectName = string.Empty;
        _bodyDrafts.Clear(); _titleDrafts.Clear(); _styleChanged = false; _hasDraftEdits = false;
        bool wasRefreshing = _refreshing;
        _refreshing = true;
        try { Stylesheet = string.Empty; } finally { _refreshing = wasRefreshing; }
        ClearPreview();
    }
    internal async Task<bool> PrepareCloseAsync() {
        if (IsBusy) return false;
        if (!IsDirty) return true;
        UnsavedChangesDecision decision = await _confirmUnsaved().ConfigureAwait(true);
        if (decision == UnsavedChangesDecision.Discard) return true;
        if (decision != UnsavedChangesDecision.Save) return false;
        await SaveProjectAsync().ConfigureAwait(true); return !IsDirty;
    }
    public void Dispose() { if (_disposed) return; _disposed = true; _operation?.Cancel(); ClearPreview(); }
}

/// <summary>A current XHTML spine position and its primary navigation title.</summary>
public sealed record BookChapterChoice(int Index, string Id, string Title, bool IsXhtml = true);
