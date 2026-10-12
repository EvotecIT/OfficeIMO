using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Platform.Storage;
using Avalonia.Styling;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Editor;
using OfficeIMO.Studio.Features.Home;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Studio.Infrastructure.Diagnostics;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindow : Window {
    private bool _allowClose;
    private bool _closePromptOpen;
    private bool _changingActiveDocument;
    private readonly StudioApplicationServices _services;
    private bool _commandPaletteOpen;
    private readonly StudioSessionController _session;
    private bool _windowClosed;

    public MainWindow() : this((Application.Current as App)?.Services ?? StudioApplicationServices.CreateDefault()) { }

    internal MainWindow(StudioApplicationServices services) {
        _services = services ?? throw new ArgumentNullException(nameof(services));
        _services.Storage.Attach(() => StorageProvider);
        // One document-less owner keeps Convert, OCR, Jobs, Settings and the other workbenches for the whole window.
        _applicationDocument = CreateDocumentViewModel((path, token) => TabHost!.OpenDocumentAsync(path, token));
        _applicationDocument.FindOpenDocument = FindOpenDocument;
        _applicationDocument.PropertyChanged += OnApplicationDocumentChanged;
        // Recognition setup opens Settings on the document the user is looking at.
        _applicationDocument.OcrReadiness.SetUpRequested += (_, _) => (ViewModel ?? _applicationDocument).ShowTextRecognitionSetup();
        TabHost = new StudioDocumentTabHost(CreateDocumentViewModel, ActivateDocument,
            document => ConfirmActiveCloseAsync([document]), _applicationDocument);
        ViewModel = TabHost.ActiveDocument;
        _session = new StudioSessionController(TabHost, _services, token => PickFileSafelyAsync(PickSavePdfAsync, token));
        ViewModel.Session = _session;
        InitializeComponent();
        AssistantHost.PropertyChanged += (_, change) => {
            if (change.Property == SplitView.IsPaneOpenProperty) ApplyResponsiveLayout(Bounds.Width);
        };
        DocumentTabs.DataContext = TabHost;
        CompactDocumentPicker.DataContext = TabHost;
        OpenDocumentTabButton.DataContext = TabHost;
        AppTitleText.DataContext = TabHost;
        DataContext = ViewModel;
        CommandSearchShortcut.Text = StudioShortcuts.Describe(StudioShortcuts.CommandSearch);
        ToolTip.SetTip(CommandSearchButton, _services.Localizer.Format("Shell.CommandSearchTip", StudioShortcuts.Describe(StudioShortcuts.CommandSearch)));
        InitializeChrome();
        RestoreWindowPlacement();
        InitializeAppMenu();
        InitializeAppleShell();
        InitializeDocumentTabInteractions();
        DocumentWorkspace.Panes = TabHost.Panes;
        AttachOperationToast(ViewModel);

        SizeChanged += OnWindowSizeChanged;
        KeyDown += OnWindowKeyDown;
        // Escape abandons a signature, initials or date waiting to be placed, before the page uses it to clear a selection.
        AddHandler(KeyDownEvent, (_, e) => {
            if (e.Key != Key.Escape || e.KeyModifiers != KeyModifiers.None || !ViewModel.IsPlacingFillSignItem) return;
            ViewModel.CancelFillSignPlacementCommand.Execute(null);
            e.Handled = true;
        }, Avalonia.Interactivity.RoutingStrategies.Tunnel);
        AddHandler(DragDrop.DragOverEvent, OnDragOver);
        AddHandler(DragDrop.DragLeaveEvent, OnDragLeave);
        AddHandler(DragDrop.DropEvent, OnDrop);
        PagesList.SizeChanged += (_, _) =>
            ViewModel.SetViewportSize(PagesList.Bounds.Width, PagesList.Bounds.Height);
        PagesList.SelectionChanged += (_, _) => {
            if (_changingActiveDocument || PagesList.SelectedItem is not PdfPageViewModel page) return;
            if (!ReferenceEquals(ViewModel.SelectedPage, page)) ViewModel.SelectedPage = page;
            PagesList.ScrollIntoView(page);
        };
        GridPagesList.SelectionChanged += (_, _) => {
            if (!_changingActiveDocument && GridPagesList.SelectedItem is ReaderGridRowViewModel row) {
                GridPagesList.ScrollIntoView(row);
            }
        };
        OrganizerList.SelectionChanged += (_, eventArgs) => {
            if (_changingActiveDocument || DocumentWorkspace.IsOrganizerPanelChanging) return;
            ViewModel.UpdateOrganizerSelection(
                eventArgs.AddedItems.OfType<PdfOrganizerPageViewModel>(),
                eventArgs.RemovedItems.OfType<PdfOrganizerPageViewModel>());
        };
        Opened += OnOpened;
        Closing += OnClosing;
        _services.Preferences.PersistenceFailed += OnPreferencesPersistenceFailed;
        _services.Preferences.Changed += OnPreferencesChanged;
        Closed += (_, _) => {
            _windowClosed = true;
            _services.Preferences.PersistenceFailed -= OnPreferencesPersistenceFailed;
            _services.Preferences.Changed -= OnPreferencesChanged;
            SaveWindowPlacement();
            _startupCompleted.TrySetResult(); _session.Dispose(); TabHost.Dispose(); _services.Storage.Dispose();
        };
    }

    /// <summary>Tells the user that an unexpected interface failure was contained and how to protect their work.</summary>
    internal void ReportUnexpectedError() {
        if (!_windowClosed) ViewModel.ErrorMessage = _services.Localizer.Get("Application.UnexpectedError");
    }

    private void OnPreferencesChanged(object? sender, EventArgs e) => SyncThemeToggle();

    private void OnPreferencesPersistenceFailed(object? sender, Exception error) {
        _services.Diagnostics.Write(StudioDiagnosticLevel.Warning, "Preferences", "SaveFailed", error);
        if (!_windowClosed) ViewModel.ErrorMessage = _services.Localizer.Get("Preferences.SaveFailed");
    }

    public StudioDocumentTabHost TabHost { get; }

    internal MainWindowViewModel ViewModel { get; private set; }

    private readonly MainWindowViewModel _applicationDocument;
    private bool _keepScreenOnActivation;

    private MainWindowViewModel? FindOpenDocument(string path) => TabHost.Tabs.Select(tab => tab.Document).FirstOrDefault(document =>
        document.DocumentPath is { Length: > 0 } source && OfficeIMO.Internal.OfficeStorageIdentity.AreEquivalent(source, path));

    // Workbench results and failures are reported by the window owner; show them on the visible document.
    private void OnApplicationDocumentChanged(object? sender, System.ComponentModel.PropertyChangedEventArgs e) {
        if (e.PropertyName != nameof(MainWindowViewModel.ErrorMessage) || _applicationDocument.ErrorMessage is not { Length: > 0 } message) return;
        _applicationDocument.ErrorMessage = null;
        if (!_windowClosed) ViewModel.ErrorMessage = message;
    }

    private MainWindowViewModel CreateDocumentViewModel(Func<string, CancellationToken, Task> openDocumentInTab) {
        MainWindowViewModel? document = null;
        document = new(
            pickPdf: token => PickFileSafelyAsync(PickPdfAsync, token),
            pickSavePdf: token => PickFileSafelyAsync(PickSavePdfAsync, token),
            pickPrintOutput: token => PickFileSafelyAsync(PickPrintOutputAsync, token),
            pickImportPdfs: token => PickFilesSafelyAsync(PickPdfsAsync, token),
            pickOutputFolder: PickOutputFolderAsync,
            openUri: OpenUriAsync,
            confirmUnsavedChanges: ConfirmUnsavedChangesAsync,
            confirmProviderWrite: ConfirmProviderWriteAsync,
            confirmWorkflowProviderWrite: location => new ProviderSaveDialog(_services.Storage.Describe(location).Name,
                _services.Localizer, workflowOutput: true, folderOutput: _services.Storage.IsFolder(location)).ShowDialog<bool>(this),
            pickImage: PickImageAsync,
            // Deleting pages stays undoable until save; the result toast offers Undo instead of a blocking dialog.
            confirmPageDeletion: static _ => Task.FromResult(true),
            reviewPageMove: preview => new PageMoveDialog(preview).ShowDialog<bool>(this),
            reviewPageSplit: preview => new PageSplitDialog(preview).ShowDialog<bool>(this),
            showPageSplitResult: result => new PageSplitDialog(result).ShowDialog(this),
            reviewProtection: preview => new PdfProtectionDialog(preview).ShowDialog<bool>(this),
            showProtectionResult: result => new PdfProtectionDialog(result).ShowDialog(this),
            reviewSigning: preview => new PdfSigningDialog(preview).ShowDialog<bool>(this),
            reviewWatermark: preview => new WatermarkDialog(preview).ShowDialog<bool>(this),
            showSigningResult: result => new PdfSigningDialog(result).ShowDialog(this),
            reviewPageExtraction: preview => new PageExtractionDialog(preview).ShowDialog<bool>(this),
            showPageExtractionResult: result => new PageExtractionDialog(result).ShowDialog(this),
            reviewPageImport: preview => new PageImportDialog(preview).ShowDialog<bool>(this),
            pickProvenanceFile: token => FileDialogs.PickOpenFileAsync(_services.Localizer.Get("Provenance.ChooseFile"),
                new StudioFileType("Provenance assets", OfficeProvenanceWorkflowCatalog.All.SelectMany(item => item.Extensions).Select(extension => extension.TrimStart('.')).Distinct().ToArray()), token),
            pickWorkflowFiles: token => PickFilesSafelyAsync(PickWorkflowFilesAsync, token),
            pickOcrFiles: token => PickFilesSafelyAsync(PickOcrFilesAsync, token),
            pickSaveComparisonReport: token => PickFileSafelyAsync(PickSaveComparisonReportAsync, token),
            pickSaveRedactionReport: token => PickFileSafelyAsync(PickSaveRedactionReportAsync, token),
            recentDocumentStore: _services.DocumentHistory.RecentDocuments,
            promptPdfPassword: PromptPdfPasswordAsync,
            canSaveAsPath: path => document is not null && TabHost.CanDocumentOwnPath(document, path),
            openDocumentInTab: openDocumentInTab,
            pickAssemblyFolder: PickAssemblyFolderAsync,
            services: _services,
            canPublishPath: path => TabHost.CanPublishPath(path),
            publicationGuard: new StudioWorkflowPublicationGuard((path, isDirectory) =>
                isDirectory ? TabHost.CanPublishDirectory(path) : TabHost.CanPublishPath(path)),
            confirmBookChanges: () => new UnsavedChangesDialog(document?.BookWorkbench.BookTitle ?? "Book", _services.Localizer).ShowDialog<UnsavedChangesDecision>(this),
            bookPublicationGuard: new StudioWorkflowPublicationGuard((path, _) => TabHost.CanPublishBookPath(document, path)),
            sharedWorkbenches: _applicationDocument);
        document.Session = _session;
        document.FileDialogs = FileDialogs;
        document.CreateSignatureDialog = kind => new Features.Sign.SignatureDialog(kind, _services.Localizer).ShowDialog<StudioSignatureDraft?>(this);
        ConfigureAssistantHost(document);
        return document;
    }

    private void ActivateDocument(MainWindowViewModel document) {
        if (ReferenceEquals(ViewModel, document)) return;
        CommandPalette.Dismiss();
        ViewModel.SaveDocumentViewState();
        StudioWorkspaceMode previousScreen = ViewModel.WorkspaceMode;
        _changingActiveDocument = true;
        try {
            DetachOperationToast(ViewModel);
            document.WorkspaceMode = ChooseScreenForActivatedDocument(previousScreen, document);
            ViewModel = document;
            DataContext = document;
            RefreshNativeMenus();
            AttachOperationToast(document);
            document.SetViewportSize(PagesList.Bounds.Width, PagesList.Bounds.Height);
        } finally {
            _changingActiveDocument = false;
        }
    }

    /// <summary>
    /// A document tab always shows its document. Closing a tab while an application screen (Convert, Jobs,
    /// Settings…) is visible keeps that screen, and with no document left the window returns to Home.
    /// </summary>
    private StudioWorkspaceMode ChooseScreenForActivatedDocument(StudioWorkspaceMode previous, MainWindowViewModel document) {
        bool applicationScreen = previous != StudioWorkspaceMode.PdfWorkspace;
        if (applicationScreen && (_keepScreenOnActivation || !document.HasDocument)) return previous;
        return document.HasDocument ? StudioWorkspaceMode.PdfWorkspace : StudioWorkspaceMode.Home;
    }

    internal bool IsCompactLayout { get; private set; }

    internal bool AreFitShortcutsVisible => FitWidthButton.IsVisible && FitPageButton.IsVisible;

    internal bool IsConversionCompact => ConversionView.IsCompactLayout;

    internal bool IsDocumentHealthCompact => DocumentHealthView.IsCompactLayout;

    internal ListBox ReaderPagesListControl => PagesList;

    internal ListBox ReaderGridPagesListControl => GridPagesList;

    private ListBox PagesList => DocumentWorkspace.PagesListControl;

    private ListBox OrganizerList => DocumentWorkspace.OrganizerListControl;

    private ListBox GridPagesList => DocumentWorkspace.GridPagesListControl;

    private Button FitWidthButton => DocumentWorkspace.FitWidthButtonControl;

    private Button FitPageButton => DocumentWorkspace.FitPageButtonControl;

    private void OnWindowSizeChanged(object? sender, SizeChangedEventArgs e) => ApplyResponsiveLayout(e.NewSize.Width);

    private void OnFindClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) => FocusDocumentSearch();

    private void FocusDocumentSearch() {
        if (!ViewModel.HasDocument) return;
        ViewModel.WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;
        DocumentWorkspace.FocusSearch();
    }

    private async void OnCommandsClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) => await ShowCommandPaletteAsync();

    internal StudioCommandPalette CommandPaletteControl => CommandPalette;

    internal async Task ShowCommandPaletteAsync() {
        if (_commandPaletteOpen) return;
        _commandPaletteOpen = true;
        var document = ViewModel;
        IInputElement? previousFocus = FocusManager?.GetFocusedElement();
        try {
            StudioCommandItem? command = await CommandPalette.ShowAsync(document.Commands);
            (previousFocus as Control)?.Focus();
            if (command is not null && ReferenceEquals(ViewModel, document)) await command.ExecuteAsync();
        } finally {
            _commandPaletteOpen = false;
        }
    }

    private void OnToggleThemeClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) {
        // High contrast is an accessibility choice made in Settings; the quick toggle never replaces it.
        if (_services.Preferences.Current.Theme == StudioThemePreference.HighContrast) return;
        StudioThemePreference preference = ActualThemeVariant == ThemeVariant.Dark
            ? StudioThemePreference.Light
            : StudioThemePreference.Dark;
        _services.Preferences.UpdatePresentation(current => current with { Theme = preference });
        SyncThemeToggle();
    }

    private void InitializeChrome() {
        ActualThemeVariantChanged += (_, _) => { SyncThemeToggle(); ApplyBackdrop(); };
        PropertyChanged += (_, change) => {
            if (change.Property == IsExtendedIntoWindowDecorationsProperty ||
                change.Property == WindowStateProperty ||
                change.Property == ActualTransparencyLevelProperty) {
                UpdateCaptionSpacing();
                ApplyBackdrop();
            }
        };
        SyncThemeToggle();
        UpdateCaptionSpacing();
        if (OperatingSystem.IsWindows()) TransparencyLevelHint = [WindowTransparencyLevel.Mica, WindowTransparencyLevel.None];
        ApplyBackdrop();
    }

    private void SyncThemeToggle() {
        bool highContrast = _services.Preferences.Current.Theme == StudioThemePreference.HighContrast;
        ThemeToggle.Classes.Set("dark", ActualThemeVariant == ThemeVariant.Dark);
        ThemeToggle.IsEnabled = !highContrast;
        string tip = _services.Localizer.Get(highContrast ? "MainWindow.HighContrastAppearanceInSettings" : "MainWindow.SwitchLightAndDarkAppearance");
        ToolTip.SetTip(ThemeToggle, tip);
        Avalonia.Automation.AutomationProperties.SetHelpText(ThemeToggle, tip);
    }

    // The operating system draws caption buttons over the extended title bar; keep tabs and actions clear of them.
    private void UpdateCaptionSpacing() {
        bool extended = IsExtendedIntoWindowDecorations;
        if (OperatingSystem.IsMacOS()) {
            CaptionButtonsSpacer.Width = 0;
        } else {
            CaptionButtonsSpacer.Width = extended ? 138 : 0;
        }
    }

    // Windows 11 Mica shows through the title bar and rail; every other platform keeps the solid chrome brush.
    private void ApplyBackdrop() {
        bool mica = ActualTransparencyLevel == WindowTransparencyLevel.Mica &&
                    ActualThemeVariant != StudioThemeVariants.HighContrast;
        if (mica) Background = Avalonia.Media.Brushes.Transparent;
        else if (this.TryFindResource("StudioChromeBrush", ActualThemeVariant, out object? brush) && brush is Avalonia.Media.IBrush chrome)
            Background = chrome;
    }

    private async void OnDocumentTabPointerReleased(object? sender, PointerReleasedEventArgs e) {
        if (e.InitialPressMouseButton != MouseButton.Middle ||
            sender is not Control { DataContext: StudioDocumentTabViewModel tab }) return;
        e.Handled = true;
        await CloseDocumentTabAsync(tab);
    }

    private async void OnWindowKeyDown(object? sender, KeyEventArgs e) {
        if (CommandPalette.IsOpen) return;
        Key key = e.Key;
        KeyModifiers modifiers = e.KeyModifiers;
        bool Is(KeyGesture gesture) => StudioShortcuts.Matches(gesture, key, modifiers);
        bool mac = OperatingSystem.IsMacOS();
        if (TryOpenAppMenuFromKeyboard(e)) {
            e.Handled = true;
            return;
        }
        if (!mac && Is(StudioShortcuts.FullScreen)) {
            ToggleFullScreen();
            e.Handled = true;
            return;
        }
        // Windows closes on Alt+F4 itself; Linux desktops expect the application to handle Ctrl+Q.
        if (!mac && !OperatingSystem.IsWindows() && Is(StudioShortcuts.Quit)) {
            Close();
            e.Handled = true;
            return;
        }
        if (key == Key.F6 && TabHost.Panes.IsSplit && ViewModel.IsPdfWorkspaceMode && !ViewModel.IsComparisonOpen) {
            TabHost.Panes.SwitchPane();
            DocumentWorkspace.FocusActiveIndependentPane();
            e.Handled = true;
            return;
        }
        if (ViewModel.Commands["FocusReading"].Gestures.Any(Is)) {
            ViewModel.IsAssistantVisible = false;
            await ViewModel.Commands["FocusReading"].ExecuteAsync();
            e.Handled = true;
            return;
        }
        if (key == Key.Escape && ViewModel.IsFocusReading) {
            ViewModel.IsFocusReading = false;
            e.Handled = true;
            return;
        }
        if (key == Key.Escape && ViewModel.IsAssistantVisible) {
            ViewModel.IsAssistantVisible = false;
            e.Handled = true;
            return;
        }
        if (Is(StudioShortcuts.CommandSearch) || Is(StudioShortcuts.CommandSearchAlternate)) {
            await ShowCommandPaletteAsync();
            e.Handled = true;
            return;
        }
        if (Is(StudioShortcuts.Find)) {
            FocusDocumentSearch();
            e.Handled = true;
            return;
        }
        if (Is(StudioShortcuts.NextTab) || Is(StudioShortcuts.PreviousTab)) {
            TabHost.SelectRelativeTab(Is(StudioShortcuts.PreviousTab));
            e.Handled = true;
            return;
        }
        if (mac && modifiers == (KeyModifiers.Meta | KeyModifiers.Shift) && key is Key.OemOpenBrackets or Key.OemCloseBrackets) {
            TabHost.SelectRelativeTab(key == Key.OemOpenBrackets);
            e.Handled = true;
            return;
        }
        if (Is(StudioShortcuts.CloseTab) || !mac && modifiers == KeyModifiers.Control && key == Key.F4) {
            await CloseSelectedDocumentTabAsync();
            e.Handled = true;
            return;
        }
        if (Is(StudioShortcuts.ReopenTab)) {
            await TabHost.ReopenClosedTabAsync();
            e.Handled = true;
            return;
        }

        // Document commands come from the catalog, so their keys match what menus and command search show.
        // A focused text field keeps its own undo, redo, zoom-like and single-letter keys.
        bool textEntry = IsTextEntryFocused();
        if (!textEntry && modifiers == KeyModifiers.None && TrySelectToolShortcut(key)) {
            e.Handled = true;
            return;
        }
        if (ViewModel.Commands.FindByGesture(key, modifiers, textEntry) is { } command) {
            await command.ExecuteAsync();
            e.Handled = true;
            return;
        }
        if (textEntry) return;

        // Page keys belong to the open document; Home, Tools and workflow screens keep their own navigation.
        if (!ViewModel.IsPdfWorkspaceMode) return;
        switch (e.Key) {
            case Key.PageUp:
            case Key.Left:
                ViewModel.PreviousPageCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.PageDown:
            case Key.Right:
                ViewModel.NextPageCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.Home:
                ViewModel.FirstPageCommand.Execute(null);
                e.Handled = true;
                break;
            case Key.End:
                ViewModel.LastPageCommand.Execute(null);
                e.Handled = true;
                break;
        }
    }

    private bool IsTextEntryFocused() {
        Control? focused = FocusManager?.GetFocusedElement() as Control;
        return focused is TextBox or ComboBox or NumericUpDown ||
               focused?.FindAncestorOfType<TextBox>() is not null ||
               focused?.FindAncestorOfType<ComboBox>() is not null ||
               focused?.FindAncestorOfType<NumericUpDown>() is not null;
    }

    private async Task<string?> PickFileSafelyAsync(Func<CancellationToken, Task<string?>> picker, CancellationToken token) {
        try { return await picker(token); }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { return null; }
        catch (Exception error) when (error is not OutOfMemoryException) {
            if (!_windowClosed) ViewModel.ErrorMessage = Infrastructure.StudioMessages.Describe(error);
            return null;
        }
    }
    private async Task<IReadOnlyList<string>> PickFilesSafelyAsync(
        Func<CancellationToken, Task<IReadOnlyList<string>>> picker, CancellationToken token) {
        try { return await picker(token); }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { return Array.Empty<string>(); }
        catch (Exception error) when (error is not OutOfMemoryException) {
            if (!_windowClosed) ViewModel.ErrorMessage = Infrastructure.StudioMessages.Describe(error);
            return Array.Empty<string>();
        }
    }

    private async Task<string?> PickPdfAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!PickerAvailable(StorageProvider.CanOpen)) return null;

        IReadOnlyList<IStorageFile> files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = _services.Localizer.Get("Picker.OpenPdf"),
            AllowMultiple = false,
            FileTypeFilter = [
                new FilePickerFileType(_services.Localizer.Get("Picker.PdfDocuments")) {
                    Patterns = StudioPickerPatterns.CaseInsensitive("*.pdf"),
                    MimeTypes = ["application/pdf"],
                    AppleUniformTypeIdentifiers = ["com.adobe.pdf"]
                }
            ]
        });
        return await _services.Storage.RegisterSingleAsync(files, cancellationToken).ConfigureAwait(true);
    }

    private async Task<IReadOnlyList<string>> PickPdfsAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!PickerAvailable(StorageProvider.CanOpen)) return Array.Empty<string>();

        IReadOnlyList<IStorageFile> files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = _services.Localizer.Get("Picker.AddPdfDocuments"),
            AllowMultiple = true,
            FileTypeFilter = [
                new FilePickerFileType(_services.Localizer.Get("Picker.PdfDocuments")) {
                    Patterns = StudioPickerPatterns.CaseInsensitive("*.pdf"),
                    MimeTypes = ["application/pdf"],
                    AppleUniformTypeIdentifiers = ["com.adobe.pdf"]
                }
            ]
        });
        return await _services.Storage.RegisterManyAsync(files, cancellationToken).ConfigureAwait(true);
    }

    private async Task<IReadOnlyList<string>> PickWorkflowFilesAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!PickerAvailable(StorageProvider.CanOpen)) return Array.Empty<string>();

        IReadOnlyList<IStorageFile> files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = _services.Localizer.Get("Picker.AddDocumentsOrImages"),
            AllowMultiple = true,
            FileTypeFilter = [
                new FilePickerFileType(_services.Localizer.Get("Picker.SupportedFiles")) {
                    Patterns = StudioPickerPatterns.CaseInsensitive("*.docx", "*.xlsx", "*.pptx", "*.pdf", "*.html", "*.htm",
                        "*.pages", "*.numbers", "*.key",
                        "*.png", "*.jpg", "*.jpeg", "*.gif", "*.bmp", "*.tif", "*.tiff",
                        "*.webp", "*.ico", "*.pcx", "*.zip")
                }
            ]
        });
        return await _services.Storage.RegisterManyAsync(files, cancellationToken).ConfigureAwait(true);
    }

    private async Task<IReadOnlyList<string>> PickOcrFilesAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!PickerAvailable(StorageProvider.CanOpen)) return [];
        var files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = _services.Localizer.Get("OcrSession.Add"), AllowMultiple = true,
            FileTypeFilter = [new FilePickerFileType(_services.Localizer.Get("OcrSession.Files")) {
                Patterns = StudioPickerPatterns.CaseInsensitive("*.pdf", "*.png", "*.jpg", "*.jpeg", "*.bmp", "*.tif", "*.tiff", "*.gif", "*.webp")
            }]
        });
        return await _services.Storage.RegisterManyAsync(files, cancellationToken).ConfigureAwait(true);
    }

    private async Task<string?> PickAssemblyFolderAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!PickerAvailable(StorageProvider.CanPickFolder)) return null;
        IReadOnlyList<IStorageFolder> folders = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions {
            Title = _services.Localizer.Get("Picker.AddSourceFolder"),
            AllowMultiple = false
        });
        return await _services.Storage.RegisterFolderAsync(folders, cancellationToken).ConfigureAwait(true);
    }

    private async void OnClosing(object? sender, WindowClosingEventArgs e) {
        if (_allowClose) return;
        if (!TabHost.HasBusyDocuments && !_session.IsBusy && !TabHost.HasDirtyDocuments && !_closePromptOpen) { _session.CaptureForShutdown(); return; }
        e.Cancel = true;
        if (_closePromptOpen) return;
        _closePromptOpen = true;
        try {
            if ((TabHost.HasBusyDocuments || _session.IsBusy) && !await ConfirmActiveCloseAsync(TabHost.OperationDocuments, wholeWindow: true)) return;
            if (TabHost.HasBusyDocuments || _session.IsBusy) return;
            if (!await TabHost.RequestCloseAllAsync()) return;
            _allowClose = true;
            Close();
        } finally {
            _closePromptOpen = false;
        }
    }

    private Task<bool> ConfirmActiveCloseAsync(IEnumerable<MainWindowViewModel> documents, bool wholeWindow = false) =>
        new ActiveOperationsDialog(documents, _services.Localizer, wholeWindow ? TabHost : null, wholeWindow ? _session : null,
            wholeWindow ? _services.Jobs : null).ShowDialog<bool>(this);

    private async Task<UnsavedChangesDecision> ConfirmUnsavedChangesAsync() {
        var dialog = new UnsavedChangesDialog(ViewModel.DocumentName.TrimEnd(' ', '*'), _services.Localizer);
        return await dialog.ShowDialog<UnsavedChangesDecision>(this);
    }

    private async Task<string?> PromptPdfPasswordAsync(
        string documentName,
        bool invalidPassword,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        var dialog = new PdfPasswordDialog(documentName, invalidPassword, _services.Localizer);
        string? password = await dialog.ShowDialog<string?>(this);
        cancellationToken.ThrowIfCancellationRequested();
        return password;
    }

    private async Task<string?> PickSavePdfAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!PickerAvailable(StorageProvider.CanSave)) return null;
        IStorageFile? file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions {
            Title = _services.Localizer.Get("Picker.SavePdf"),
            SuggestedFileName = System.IO.Path.GetFileNameWithoutExtension(ViewModel.DocumentName.TrimEnd(' ', '*')),
            DefaultExtension = "pdf",
            FileTypeChoices = [
                new FilePickerFileType(_services.Localizer.Get("Picker.PdfDocuments")) {
                    Patterns = StudioPickerPatterns.CaseInsensitive("*.pdf"),
                    MimeTypes = ["application/pdf"],
                    AppleUniformTypeIdentifiers = ["com.adobe.pdf"]
                }
            ]
        });
        string? location = await _services.Storage.RegisterSingleAsync(file is null ? [] : [file], cancellationToken).ConfigureAwait(true);
        return location is not null && await ConfirmProviderWriteAsync(location) ? location : null;
    }

    private async Task<string?> PickSaveRedactionReportAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!PickerAvailable(StorageProvider.CanSave)) return null;
        IStorageFile? file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions {
            Title = _services.Localizer.Get("Redaction.ExportReport"),
            SuggestedFileName = "redaction-evidence.json", DefaultExtension = "json",
            FileTypeChoices = [new FilePickerFileType("JSON") { Patterns = StudioPickerPatterns.CaseInsensitive("*.json"), MimeTypes = ["application/json"] }]
        });
        string? location = await _services.Storage.RegisterSingleAsync(file is null ? [] : [file], cancellationToken).ConfigureAwait(true);
        return location is not null && await ConfirmProviderWriteAsync(location) ? location : null;
    }

    private async Task<string?> PickSaveComparisonReportAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!PickerAvailable(StorageProvider.CanSave)) return null;
        IStorageFile? file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions {
            Title = _services.Localizer.GetOrDefault("Comparison.ExportReport", "Export comparison report"),
            SuggestedFileName = "comparison.html", DefaultExtension = "html",
            FileTypeChoices = [new FilePickerFileType("HTML") { Patterns = StudioPickerPatterns.CaseInsensitive("*.html"), MimeTypes = ["text/html"] }]
        });
        string? location = await _services.Storage.RegisterSingleAsync(file is null ? [] : [file], cancellationToken).ConfigureAwait(true);
        return location is not null && await ConfirmProviderWriteAsync(location) ? location : null;
    }

    private Task<bool> ConfirmProviderWriteAsync(string location) =>
        !_services.Storage.UsesProviderPublication(location) ? Task.FromResult(true)
            : new ProviderSaveDialog(_services.Storage.Describe(location).Name, _services.Localizer).ShowDialog<bool>(this);
    private async Task<string?> PickOutputFolderAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!PickerAvailable(StorageProvider.CanPickFolder)) return null;
        IReadOnlyList<IStorageFolder> folders = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions {
            Title = _services.Localizer.Get("Picker.ChooseOutputFolder"),
            AllowMultiple = false
        });
        return await _services.Storage.RegisterFolderAsync(folders, cancellationToken).ConfigureAwait(true);
    }

    private async Task<byte[]?> PickImageAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!PickerAvailable(StorageProvider.CanOpen)) return null;
        IReadOnlyList<IStorageFile> files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = _services.Localizer.Get("Picker.ChooseImage"),
            AllowMultiple = false,
            FileTypeFilter = [
                new FilePickerFileType(_services.Localizer.Get("Picker.PngOrJpegImages")) {
                    Patterns = StudioPickerPatterns.CaseInsensitive("*.png", "*.jpg", "*.jpeg"),
                    MimeTypes = ["image/png", "image/jpeg"],
                    AppleUniformTypeIdentifiers = ["public.png", "public.jpeg"]
                }
            ]
        });
        return await StudioStorageInput.ReadImageAsync(files, cancellationToken).ConfigureAwait(true);
    }

    private async Task OpenUriAsync(Uri uri) {
        bool opened = await Launcher.LaunchUriAsync(uri);
        if (!opened) throw new InvalidOperationException(_services.Localizer.Get("Error.CouldNotOpenLink"));
    }


}
