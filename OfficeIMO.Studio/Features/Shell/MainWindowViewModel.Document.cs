using System.Collections.ObjectModel;
using System.Globalization;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Document properties, bookmark editing, attachments, headers and footers, exports and form data.</summary>
public sealed partial class MainWindowViewModel {
    private IStudioFileDialogs _fileDialogs = NoStudioFileDialogs.Instance;
    private PdfWorkspace? _propertiesWorkspace;
    private (string Title, string Author, string Subject, string Keywords) _savedProperties = (string.Empty, string.Empty, string.Empty, string.Empty);

    /// <summary>Replaces the file pickers used by document exports, attachments and form data.</summary>
    internal IStudioFileDialogs FileDialogs {
        get => _fileDialogs;
        set => _fileDialogs = value ?? NoStudioFileDialogs.Instance;
    }

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasPropertyChanges))]
    private string _propertyTitle = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasPropertyChanges))]
    private string _propertyAuthor = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasPropertyChanges))]
    private string _propertySubject = string.Empty;

    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(HasPropertyChanges))]
    private string _propertyKeywords = string.Empty;

    [ObservableProperty]
    private string _bookmarkTitleDraft = string.Empty;

    [ObservableProperty]
    private PdfAttachmentItemViewModel? _selectedAttachment;

    [ObservableProperty] private string _headerLeft = string.Empty;
    [ObservableProperty] private string _headerCenter = string.Empty;
    [ObservableProperty] private string _headerRight = string.Empty;
    [ObservableProperty] private string _footerLeft = string.Empty;
    [ObservableProperty] private string _footerCenter = "{page} / {pages}";
    [ObservableProperty] private string _footerRight = string.Empty;
    [ObservableProperty] private double _headerFooterFontSize = 9D;
    [ObservableProperty] private int _headerFooterStartNumber = 1;
    [ObservableProperty] private string _headerFooterPages = string.Empty;

    /// <summary>Open and Show in folder for result screens.</summary>
    public Features.Workflows.StudioOutputActions OutputActions { get; private set; } = null!;

    public ObservableCollection<PdfAttachmentItemViewModel> DocumentAttachments { get; } = [];

    public bool HasDocumentAttachments => DocumentAttachments.Count > 0;

    public bool CanEditMetadata => !IsWorkspaceBusy && _workspace?.CanEditMetadata == true;

    public bool CanEditBookmarks => !IsWorkspaceBusy && _workspace?.CanEditBookmarks == true;

    public bool CanEditAttachments => !IsWorkspaceBusy && _workspace?.CanEditAttachments == true;

    public bool HasPropertyChanges =>
        !string.Equals(PropertyTitle.Trim(), _savedProperties.Title, StringComparison.Ordinal) ||
        !string.Equals(PropertyAuthor.Trim(), _savedProperties.Author, StringComparison.Ordinal) ||
        !string.Equals(PropertySubject.Trim(), _savedProperties.Subject, StringComparison.Ordinal) ||
        !string.Equals(PropertyKeywords.Trim(), _savedProperties.Keywords, StringComparison.Ordinal);

    public bool HasEditableBookmarkSelection => SelectedBookmark?.IsEditable == true;

    public string DocumentVersionText => _workspace?.DocumentInfo?.EffectiveVersion is { Length: > 0 } version ? "PDF " + version : string.Empty;

    public string DocumentCreatedText => FormatDate(_workspace?.Metadata?.CreationDate);

    public string DocumentModifiedText => FormatDate(_workspace?.Metadata?.ModificationDate);

    private static string FormatDate(DateTimeOffset? value) =>
        value is { } date ? date.ToLocalTime().ToString("g", CultureInfo.CurrentCulture) : string.Empty;

    partial void OnSelectedAttachmentChanged(PdfAttachmentItemViewModel? value) => NotifyAttachmentActions();

    // Keeps document facts, property drafts and the attachment list in step with the current revision.
    private void RefreshDocumentStructure() {
        PdfMetadata? metadata = _workspace?.Metadata;
        var current = (metadata?.Title ?? string.Empty, metadata?.Author ?? string.Empty, metadata?.Subject ?? string.Empty, metadata?.Keywords ?? string.Empty);
        // A pending draft survives unrelated edits to the same document, never a document switch.
        bool keepDraft = HasPropertyChanges && _workspace is not null && ReferenceEquals(_propertiesWorkspace, _workspace);
        _propertiesWorkspace = _workspace;
        _savedProperties = current;
        if (!keepDraft) {
            PropertyTitle = current.Item1;
            PropertyAuthor = current.Item2;
            PropertySubject = current.Item3;
            PropertyKeywords = current.Item4;
        }
        OnPropertyChanged(nameof(HasPropertyChanges));
        OnPropertyChanged(nameof(DocumentVersionText));
        OnPropertyChanged(nameof(DocumentCreatedText));
        OnPropertyChanged(nameof(DocumentModifiedText));

        string? selectedAttachment = SelectedAttachment?.FileName;
        DocumentAttachments.Clear();
        foreach (PdfAttachmentInfo attachment in _workspace?.Attachments ?? [])
            DocumentAttachments.Add(new PdfAttachmentItemViewModel(attachment.FileName, attachment.Description,
                FormatByteSize(attachment.DecodedSizeBytes ?? attachment.SizeBytes)));
        SelectedAttachment = DocumentAttachments.FirstOrDefault(item => item.FileName == selectedAttachment);
        OnPropertyChanged(nameof(HasDocumentAttachments));
        InvalidateComplianceResult();
        NotifyDocumentStructureActions();
    }

    private void NotifyDocumentStructureActions() {
        OnPropertyChanged(nameof(CanEditMetadata));
        OnPropertyChanged(nameof(CanEditBookmarks));
        OnPropertyChanged(nameof(CanEditAttachments));
        OnPropertyChanged(nameof(CanFillAndSign));
        ApplyPropertiesCommand.NotifyCanExecuteChanged();
        ResetPropertiesCommand.NotifyCanExecuteChanged();
        AddAttachmentCommand.NotifyCanExecuteChanged();
        CheckComplianceCommand.NotifyCanExecuteChanged();
        ApplyHeaderFooterCommand.NotifyCanExecuteChanged();
        NotifyBookmarkActions();
        NotifyAttachmentActions();
    }

    private void NotifyBookmarkActions() {
        OnPropertyChanged(nameof(HasEditableBookmarkSelection));
        AddBookmarkCommand.NotifyCanExecuteChanged();
        RenameBookmarkCommand.NotifyCanExecuteChanged();
        RemoveBookmarkCommand.NotifyCanExecuteChanged();
        MoveBookmarkUpCommand.NotifyCanExecuteChanged();
        MoveBookmarkDownCommand.NotifyCanExecuteChanged();
        IndentBookmarkCommand.NotifyCanExecuteChanged();
        OutdentBookmarkCommand.NotifyCanExecuteChanged();
        RetargetBookmarkCommand.NotifyCanExecuteChanged();
        GenerateBookmarksCommand.NotifyCanExecuteChanged();
    }

    private void NotifyAttachmentActions() {
        SaveAttachmentCommand.NotifyCanExecuteChanged();
        RemoveAttachmentCommand.NotifyCanExecuteChanged();
    }

    partial void OnPropertyTitleChanged(string value) => NotifyPropertyCommands();
    partial void OnPropertyAuthorChanged(string value) => NotifyPropertyCommands();
    partial void OnPropertySubjectChanged(string value) => NotifyPropertyCommands();
    partial void OnPropertyKeywordsChanged(string value) => NotifyPropertyCommands();

    private void NotifyPropertyCommands() {
        ApplyPropertiesCommand.NotifyCanExecuteChanged();
        ResetPropertiesCommand.NotifyCanExecuteChanged();
    }

    private bool CanApplyProperties() => CanEditMetadata && HasPropertyChanges;

    [RelayCommand(CanExecute = nameof(CanApplyProperties))]
    private async Task ApplyPropertiesAsync(CancellationToken cancellationToken) {
        if (_workspace is not { } workspace) return;
        string title = PropertyTitle, author = PropertyAuthor, subject = PropertySubject, keywords = PropertyKeywords;
        await RunMutationAsync(token => workspace.UpdateMetadataAsync(title, author, subject, keywords, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
        RefreshDocumentStructure();
    }

    private bool CanResetProperties() => HasPropertyChanges;

    [RelayCommand(CanExecute = nameof(CanResetProperties))]
    private void ResetProperties() {
        PropertyTitle = _savedProperties.Title;
        PropertyAuthor = _savedProperties.Author;
        PropertySubject = _savedProperties.Subject;
        PropertyKeywords = _savedProperties.Keywords;
    }

    // Bookmarks -----------------------------------------------------------------------------------

    private bool CanAddBookmark() => CanEditBookmarks && SelectedPage is not null;

    [RelayCommand(CanExecute = nameof(CanAddBookmark))]
    private Task AddBookmarkAsync(CancellationToken cancellationToken) {
        int page = SelectedPage?.PageNumber ?? 1;
        string title = UiFormat("Bookmarks.DefaultTitle", page);
        PdfBookmarkViewModel? anchor = SelectedBookmark is { IsEditable: true } selected ? selected : null;
        // New entries go after the selection at the same level, or at the end of the outline.
        return EditBookmarksAsync([page], session => session.Move(session.Add(title, page).Id, anchor?.ParentId, anchor is null ? -1 : anchor.Index + 1),
            cancellationToken, selectAfter: bookmarks => bookmarks.LastOrDefault(item => item.Title == title && item.PageNumber == page));
    }

    private bool CanEditSelectedBookmark() => CanEditBookmarks && HasEditableBookmarkSelection;

    private bool CanRenameBookmark() => CanEditSelectedBookmark() && !string.IsNullOrWhiteSpace(BookmarkTitleDraft) &&
        !string.Equals(BookmarkTitleDraft.Trim(), SelectedBookmark?.Title, StringComparison.Ordinal);

    partial void OnBookmarkTitleDraftChanged(string value) => RenameBookmarkCommand.NotifyCanExecuteChanged();

    [RelayCommand(CanExecute = nameof(CanRenameBookmark))]
    private Task RenameBookmarkAsync(CancellationToken cancellationToken) {
        PdfBookmarkViewModel bookmark = SelectedBookmark!;
        string title = BookmarkTitleDraft.Trim();
        return EditBookmarksAsync([], session => session.Rename(bookmark.Id!, title), cancellationToken,
            selectAfter: bookmarks => bookmarks.FirstOrDefault(item => item.Id == bookmark.Id));
    }

    [RelayCommand(CanExecute = nameof(CanEditSelectedBookmark))]
    private Task RemoveBookmarkAsync(CancellationToken cancellationToken) {
        PdfBookmarkViewModel bookmark = SelectedBookmark!;
        return EditBookmarksAsync([], session => session.Remove(bookmark.Id!), cancellationToken, selectAfter: _ => null);
    }

    private bool CanMoveBookmarkUp() => CanEditSelectedBookmark() && SelectedBookmark!.Index > 0;

    [RelayCommand(CanExecute = nameof(CanMoveBookmarkUp))]
    private Task MoveBookmarkUpAsync(CancellationToken cancellationToken) {
        PdfBookmarkViewModel bookmark = SelectedBookmark!;
        return EditBookmarksAsync([], session => session.Move(bookmark.Id!, bookmark.ParentId, bookmark.Index - 1), cancellationToken,
            selectAfter: bookmarks => FindMoved(bookmarks, bookmark));
    }

    private bool CanMoveBookmarkDown() => CanEditSelectedBookmark() && Bookmarks.Any(item => item.ParentId == SelectedBookmark!.ParentId && item.Index == SelectedBookmark.Index + 1 && item.IsEditable);

    [RelayCommand(CanExecute = nameof(CanMoveBookmarkDown))]
    private Task MoveBookmarkDownAsync(CancellationToken cancellationToken) {
        PdfBookmarkViewModel bookmark = SelectedBookmark!;
        return EditBookmarksAsync([], session => session.Move(bookmark.Id!, bookmark.ParentId, bookmark.Index + 1), cancellationToken,
            selectAfter: bookmarks => FindMoved(bookmarks, bookmark));
    }

    private PdfBookmarkViewModel? PreviousSibling(PdfBookmarkViewModel bookmark) =>
        Bookmarks.FirstOrDefault(item => item.ParentId == bookmark.ParentId && item.Index == bookmark.Index - 1 && item.IsEditable);

    private bool CanIndentBookmark() => CanEditSelectedBookmark() && PreviousSibling(SelectedBookmark!) is not null;

    [RelayCommand(CanExecute = nameof(CanIndentBookmark))]
    private Task IndentBookmarkAsync(CancellationToken cancellationToken) {
        PdfBookmarkViewModel bookmark = SelectedBookmark!;
        PdfBookmarkViewModel parent = PreviousSibling(bookmark)!;
        return EditBookmarksAsync([], session => session.Move(bookmark.Id!, parent.Id), cancellationToken,
            selectAfter: bookmarks => FindMoved(bookmarks, bookmark));
    }

    private bool CanOutdentBookmark() => CanEditSelectedBookmark() && SelectedBookmark!.ParentId is not null;

    [RelayCommand(CanExecute = nameof(CanOutdentBookmark))]
    private Task OutdentBookmarkAsync(CancellationToken cancellationToken) {
        PdfBookmarkViewModel bookmark = SelectedBookmark!;
        PdfBookmarkViewModel? parent = Bookmarks.FirstOrDefault(item => item.Id == bookmark.ParentId);
        if (parent is null) return Task.CompletedTask;
        return EditBookmarksAsync([], session => session.Move(bookmark.Id!, parent.ParentId, parent.Index + 1), cancellationToken,
            selectAfter: bookmarks => FindMoved(bookmarks, bookmark));
    }

    private bool CanRetargetBookmark() => CanEditSelectedBookmark() && SelectedPage is not null && SelectedPage.PageNumber != SelectedBookmark!.PageNumber;

    [RelayCommand(CanExecute = nameof(CanRetargetBookmark))]
    private Task RetargetBookmarkAsync(CancellationToken cancellationToken) {
        PdfBookmarkViewModel bookmark = SelectedBookmark!;
        int page = SelectedPage!.PageNumber;
        return EditBookmarksAsync([page], session => session.Retarget(bookmark.Id!, page), cancellationToken,
            selectAfter: bookmarks => bookmarks.FirstOrDefault(item => item.Id == bookmark.Id));
    }

    private bool CanGenerateBookmarks() => CanEditBookmarks && _workspace?.ViewInfo.CanExtractContent == true;

    [RelayCommand(CanExecute = nameof(CanGenerateBookmarks))]
    private Task GenerateBookmarksAsync(CancellationToken cancellationToken) =>
        EditBookmarksAsync([], session => session.RebuildFromHeadings(), cancellationToken, selectAfter: _ => null,
            describe: () => Bookmarks.Count == 0 ? UiText("Bookmarks.NoHeadings") : UiFormat("Bookmarks.Generated", Bookmarks.Count));

    // Ids are positional, so the moved entry is found again by title and target page.
    private static PdfBookmarkViewModel? FindMoved(IEnumerable<PdfBookmarkViewModel> bookmarks, PdfBookmarkViewModel moved) =>
        bookmarks.FirstOrDefault(item => item.Title == moved.Title && item.PageNumber == moved.PageNumber);

    private async Task EditBookmarksAsync(IReadOnlyList<int> pages, Action<PdfBookmarkEditSession> edit, CancellationToken cancellationToken,
        Func<IReadOnlyList<PdfBookmarkViewModel>, PdfBookmarkViewModel?> selectAfter, Func<string?>? describe = null) {
        if (_workspace is not { } workspace) return;
        bool succeeded = await RunMutationAsync(token => workspace.EditBookmarksAsync(UiText("Operation.Bookmarks"), pages, edit, token, CreateProgress()),
            cancellationToken).ConfigureAwait(true);
        if (!succeeded) return;
        SelectedBookmark = selectAfter(Bookmarks);
        NotifyBookmarkActions();
        if (describe?.Invoke() is { } status) OperationStatus = status;
    }

    // Attachments ---------------------------------------------------------------------------------

    [RelayCommand(CanExecute = nameof(CanEditAttachments))]
    private async Task AddAttachmentAsync(CancellationToken cancellationToken) {
        if (_workspace is not { } workspace) return;
        string? source = await _fileDialogs.PickOpenFileAsync(UiText("Attachments.Add"), StudioFileType.Any(UiText("Attachments.AllFiles")), cancellationToken).ConfigureAwait(true);
        if (source is null || !ReferenceEquals(workspace, _workspace)) return;
        await RunMutationAsync(async token => {
            var snapshot = await _services.Storage.ReadSnapshotAsync(source, token).ConfigureAwait(true);
            await workspace.AddAttachmentAsync(_services.Storage.Describe(source).Name, snapshot.Bytes, token, CreateProgress()).ConfigureAwait(true);
        }, cancellationToken).ConfigureAwait(true);
        RefreshDocumentStructure();
    }

    private bool HasAttachmentSelection(PdfAttachmentItemViewModel? item) => (item ?? SelectedAttachment) is not null && !IsWorkspaceBusy;

    [RelayCommand(CanExecute = nameof(HasAttachmentSelection))]
    private async Task SaveAttachmentAsync(PdfAttachmentItemViewModel? item, CancellationToken cancellationToken) {
        if (_workspace is not { } workspace || (item ?? SelectedAttachment) is not { } attachment) return;
        string extension = Path.GetExtension(attachment.FileName);
        var type = new StudioFileType(string.IsNullOrEmpty(extension) ? UiText("Attachments.AllFiles") : extension.TrimStart('.').ToUpperInvariant(),
            [string.IsNullOrEmpty(extension) ? "*" : extension.TrimStart('.')]);
        string? destination = await _fileDialogs.PickSaveFileAsync(UiText("Attachments.Save"), attachment.FileName, type, cancellationToken).ConfigureAwait(true);
        if (destination is null || !ReferenceEquals(workspace, _workspace)) return;
        await RunStandaloneAsync(token => workspace.SaveAttachmentAsync(attachment.FileName, destination, token), cancellationToken,
            UiFormat("Attachments.Saved", attachment.FileName)).ConfigureAwait(true);
    }

    private bool CanRemoveAttachment(PdfAttachmentItemViewModel? item) => CanEditAttachments && (item ?? SelectedAttachment) is not null;

    [RelayCommand(CanExecute = nameof(CanRemoveAttachment))]
    private async Task RemoveAttachmentAsync(PdfAttachmentItemViewModel? item, CancellationToken cancellationToken) {
        if (_workspace is not { } workspace || (item ?? SelectedAttachment) is not { } attachment) return;
        await RunMutationAsync(token => workspace.RemoveAttachmentAsync(attachment.FileName, token, CreateProgress()), cancellationToken).ConfigureAwait(true);
        RefreshDocumentStructure();
    }

    // Header and footer ---------------------------------------------------------------------------

    private bool CanApplyHeaderFooter() => CanEditPageContent && !IsWorkspaceBusy;

    [RelayCommand(CanExecute = nameof(CanApplyHeaderFooter))]
    private async Task ApplyHeaderFooterAsync(CancellationToken cancellationToken) {
        if (_workspace is not { } workspace) return;
        var options = new PdfHeaderFooterOptions {
            HeaderLeft = HeaderLeft, HeaderCenter = HeaderCenter, HeaderRight = HeaderRight,
            FooterLeft = FooterLeft, FooterCenter = FooterCenter, FooterRight = FooterRight,
            FontSize = Math.Clamp(HeaderFooterFontSize, 5D, 36D),
            StartNumber = Math.Max(1, HeaderFooterStartNumber),
            PageRange = string.IsNullOrWhiteSpace(HeaderFooterPages) ? null : HeaderFooterPages
        };
        if (!options.HasContent) { ErrorMessage = UiText("HeaderFooter.Empty"); return; }
        try { options.ResolvePages(workspace.Pages.Count); } catch (Exception) { ErrorMessage = UiText("HeaderFooter.InvalidPages"); return; }
        await RunMutationAsync(token => workspace.ApplyHeaderFooterAsync(options, token, CreateProgress()), cancellationToken).ConfigureAwait(true);
    }

    // Page size --------------------------------------------------------------------------------------

    public IReadOnlyList<PageSizeChoice> PageSizeChoices { get; } = [
        new("A4", PageSizes.A4), new("Letter", PageSizes.Letter), new("Legal", PageSizes.Legal), new("A3", PageSizes.A3), new("A5", PageSizes.A5)
    ];

    public IReadOnlyList<ResizeModeChoice> ResizeModeChoices => [
        new(PdfPageResizeMode.Fit, UiText("Resize.Fit")), new(PdfPageResizeMode.Fill, UiText("Resize.Fill")), new(PdfPageResizeMode.Stretch, UiText("Resize.Stretch"))
    ];

    [ObservableProperty] private PageSizeChoice? _selectedPageSize;
    [ObservableProperty] private ResizeModeChoice? _selectedResizeMode;

    [RelayCommand]
    private async Task ResizeSelectedPagesAsync(CancellationToken cancellationToken) {
        int[] pages = GetSelectedPages();
        if (_workspace is not { } workspace || pages.Length == 0 || !CanMutateSelection) return;
        PageSizeChoice size = SelectedPageSize ??= PageSizeChoices[0];
        ResizeModeChoice mode = SelectedResizeMode ??= ResizeModeChoices[0];
        await RunMutationAsync(token => workspace.ResizePagesAsync(pages, size.Size, mode.Mode, token, CreateProgress()), cancellationToken, pages).ConfigureAwait(true);
    }

    // Export and form data ------------------------------------------------------------------------

    private bool CanExportDocument(string? kind) => _workspace is not null && !IsWorkspaceBusy &&
        (kind == nameof(PdfExportKind.FormData) ? FormFields.Count > 0 : _workspace.ViewInfo.CanExtractContent);

    [RelayCommand(CanExecute = nameof(CanExportDocument))]
    private async Task ExportDocumentAsync(string? kind, CancellationToken cancellationToken) {
        if (_workspace is not { } workspace || !Enum.TryParse(kind, out PdfExportKind exportKind)) return;
        string baseName = Path.GetFileNameWithoutExtension(workspace.FileName);
        string label = UiText("Export." + exportKind);
        string? destination = exportKind == PdfExportKind.Images
            ? await _fileDialogs.PickFolderAsync(UiText("Export.ChooseImageFolder"), cancellationToken).ConfigureAwait(true)
            : await _fileDialogs.PickSaveFileAsync(label, baseName + ExportExtension(exportKind),
                new StudioFileType(label, [ExportExtension(exportKind).TrimStart('.')]), cancellationToken).ConfigureAwait(true);
        if (destination is null || !ReferenceEquals(workspace, _workspace)) return;
        IReadOnlyDictionary<string, PdfFormFieldValue>? formDrafts = exportKind == PdfExportKind.FormData
            ? CaptureFormDrafts() : null;
        if (exportKind == PdfExportKind.FormData && formDrafts is null) return;
        int written = 0;
        await RunStandaloneAsync(async token => written = await workspace.ExportDocumentAsync(exportKind, destination, token, formDrafts).ConfigureAwait(true),
            cancellationToken, describeSuccess: () => exportKind == PdfExportKind.Images
                ? UiFormat("Export.ImagesSaved", written, _services.Storage.Describe(destination).Name)
                : UiFormat("Export.Saved", _services.Storage.Describe(destination).Name)).ConfigureAwait(true);
    }

    private static string ExportExtension(PdfExportKind kind) => kind switch {
        PdfExportKind.Markdown => ".md",
        PdfExportKind.Json => ".json",
        PdfExportKind.FormData => ".xfdf",
        _ => ".txt"
    };

    private bool CanConvertDocument(string? target) => _workspace is not null && !IsWorkspaceBusy && !string.IsNullOrWhiteSpace(DocumentPath);

    /// <summary>Queues the saved copy of this PDF for conversion to Word, Excel or PowerPoint.</summary>
    [RelayCommand(CanExecute = nameof(CanConvertDocument))]
    private void ConvertDocument(string? target) {
        if (DocumentPath is not { Length: > 0 } path || string.IsNullOrWhiteSpace(target)) return;
        IsFocusReading = false;
        WorkspaceMode = StudioWorkspaceMode.Convert;
        if (!ConversionWorkbench.QueueForTarget(path, target)) {
            ErrorMessage = UiFormat("Export.NoRoute", target.TrimStart('.').ToUpperInvariant());
            return;
        }
        if (IsDirty) OperationStatus = UiText("Export.SavedCopyOnly");
    }

    private bool CanImportFormData() => CanFillForms && FormFields.Count > 0 && !IsWorkspaceBusy;

    [RelayCommand(CanExecute = nameof(CanImportFormData))]
    private async Task ImportFormDataAsync(CancellationToken cancellationToken) {
        if (_workspace is not { } workspace) return;
        string? source = await _fileDialogs.PickOpenFileAsync(UiText("Forms.ImportData"), new StudioFileType("XFDF", ["xfdf", "xml"]), cancellationToken).ConfigureAwait(true);
        if (source is null || !ReferenceEquals(workspace, _workspace)) return;
        await RunMutationAsync(token => workspace.ImportFormDataAsync(source, token, CreateProgress()), cancellationToken).ConfigureAwait(true);
    }
}

public sealed record PageSizeChoice(string Label, PageSize Size);

public sealed record ResizeModeChoice(PdfPageResizeMode Mode, string Label);

public sealed record PdfAttachmentItemViewModel(string FileName, string? Description, string SizeText) {
    public string Detail => string.IsNullOrWhiteSpace(Description) ? SizeText : SizeText + " · " + Description;
}
