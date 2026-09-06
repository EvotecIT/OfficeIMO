using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Platform.Storage;
using Avalonia.Styling;
using Avalonia.VisualTree;
using OfficeIMO.Studio.Features.Organizer;
using OfficeIMO.Studio.Features.Home;
using OfficeIMO.Studio.Features.Reader;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;
using OfficeIMO.Studio.Infrastructure.Diagnostics;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindow : Window {
    private static readonly DataFormat<string> OrganizerPageFormat =
        DataFormat.CreateInProcessFormat<string>("officeimo-studio-organizer-page");
    private string? _initialDocumentPath;
    private bool _initialDocumentOpened;
    private bool _allowClose;
    private bool _closePromptOpen;
    private PointerPressedEventArgs? _organizerDragPress;
    private PdfOrganizerPageViewModel? _organizerDragPage;
    private Point _organizerDragStart;
    private bool _organizerDragStarted;
    private bool _changingActiveDocument;
    private readonly StudioApplicationServices _services;
    private bool _commandPaletteOpen;
    private readonly StudioSessionController _session;
    private bool _windowClosed;

    public MainWindow() : this((Application.Current as App)?.Services ?? StudioApplicationServices.CreateDefault()) { }

    internal MainWindow(StudioApplicationServices services) {
        _services = services ?? throw new ArgumentNullException(nameof(services));
        _services.Storage.Attach(() => StorageProvider);
        TabHost = new StudioDocumentTabHost(CreateDocumentViewModel, ActivateDocument,
            document => ConfirmActiveCloseAsync([document]));
        ViewModel = TabHost.ActiveDocument;
        _session = new StudioSessionController(TabHost, _services, token => PickFileSafelyAsync(PickSavePdfAsync, token));
        ViewModel.Session = _session;
        InitializeComponent();
        DocumentTabs.DataContext = TabHost;
        OpenDocumentTabButton.DataContext = TabHost;
        DataContext = ViewModel;

        SizeChanged += OnWindowSizeChanged;
        KeyDown += OnWindowKeyDown;
        AddHandler(DragDrop.DragOverEvent, OnDragOver);
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
            if (_changingActiveDocument) return;
            ViewModel.UpdateOrganizerSelection(
                eventArgs.AddedItems.OfType<PdfOrganizerPageViewModel>(),
                eventArgs.RemovedItems.OfType<PdfOrganizerPageViewModel>());
        };
        OrganizerList.KeyDown += OnOrganizerKeyDown;
        OrganizerList.AddHandler(PointerPressedEvent, OnOrganizerPointerPressed, handledEventsToo: true);
        OrganizerList.AddHandler(PointerMovedEvent, OnOrganizerPointerMoved, handledEventsToo: true);
        OrganizerList.AddHandler(PointerReleasedEvent, OnOrganizerPointerReleased, handledEventsToo: true);
        OrganizerList.AddHandler(DragDrop.DragOverEvent, OnOrganizerDragOver);
        OrganizerList.AddHandler(DragDrop.DropEvent, OnOrganizerDrop);
        Opened += OnOpened;
        Closing += OnClosing;
        Closed += (_, _) => { _windowClosed = true; _session.Dispose(); TabHost.Dispose(); _services.Storage.Dispose(); };
    }

    public StudioDocumentTabHost TabHost { get; }

    internal MainWindowViewModel ViewModel { get; private set; }

    private MainWindowViewModel CreateDocumentViewModel(Func<string, CancellationToken, Task> openDocumentInTab) {
        MainWindowViewModel? document = null;
        document = new(
            pickPdf: token => PickFileSafelyAsync(PickPdfAsync, token),
            pickSavePdf: token => PickFileSafelyAsync(PickSavePdfAsync, token),
            pickImportPdfs: token => PickFilesSafelyAsync(PickPdfsAsync, token),
            pickOutputFolder: PickOutputFolderAsync,
            openUri: OpenUriAsync,
            confirmUnsavedChanges: ConfirmUnsavedChangesAsync,
            confirmProviderWrite: ConfirmProviderWriteAsync,
            pickImage: PickImageAsync,
            confirmPageDeletion: ConfirmPageDeletionAsync,
            pickWorkflowFiles: token => PickFilesSafelyAsync(PickWorkflowFilesAsync, token),
            recentDocumentStore: _services.DocumentHistory.RecentDocuments,
            promptPdfPassword: PromptPdfPasswordAsync,
            canSaveAsPath: path => document is not null && TabHost.CanDocumentOwnPath(document, path),
            openDocumentInTab: openDocumentInTab,
            pickAssemblyFolder: PickAssemblyFolderAsync,
            services: _services,
            canPublishPath: path => TabHost.CanPublishPath(path),
            publicationGuard: new StudioWorkflowPublicationGuard((path, isDirectory) =>
                isDirectory ? TabHost.CanPublishDirectory(path) : TabHost.CanPublishPath(path)));
        document.Session = _session;
        return document;
    }

    private void ActivateDocument(MainWindowViewModel document) {
        if (ReferenceEquals(ViewModel, document)) return;
        ViewModel.SaveDocumentViewState();
        _changingActiveDocument = true;
        try {
            ViewModel = document;
            DataContext = document;
            ClearOrganizerDrag();
            document.SetViewportSize(PagesList.Bounds.Width, PagesList.Bounds.Height);
        } finally {
            _changingActiveDocument = false;
        }
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

    internal void ApplyResponsiveLayout(double width) {
        IsCompactLayout = width < 1180D;
        double workspaceWidth = Math.Max(0D, width - 116D);
        DocumentWorkspace.ApplyResponsiveLayout(workspaceWidth);
        ConversionView.ApplyResponsiveLayout(workspaceWidth);
        DocumentHealthView.ApplyResponsiveLayout(workspaceWidth);
    }

    private void OnFindClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) => FocusDocumentSearch();

    private void FocusDocumentSearch() {
        if (!ViewModel.HasDocument) return;
        ViewModel.WorkspaceMode = StudioWorkspaceMode.PdfWorkspace;
        DocumentWorkspace.FocusSearch();
    }

    private async void OnCommandsClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) => await ShowCommandPaletteAsync();

    internal async Task ShowCommandPaletteAsync() {
        if (_commandPaletteOpen) return;
        _commandPaletteOpen = true;
        try {
            var palette = new StudioCommandPalette(ViewModel.Commands);
            StudioCommandItem? command = await palette.ShowDialog<StudioCommandItem?>(this);
            if (command is not null) await command.ExecuteAsync();
        } finally {
            _commandPaletteOpen = false;
        }
    }

    private void OnToggleThemeClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) {
        if (Application.Current is not { } application) return;
        application.RequestedThemeVariant = application.ActualThemeVariant == ThemeVariant.Dark
            ? ThemeVariant.Light
            : ThemeVariant.Dark;
        StudioThemePreference preference = application.RequestedThemeVariant == ThemeVariant.Dark
            ? StudioThemePreference.Dark
            : StudioThemePreference.Light;
        _services.Preferences.Update(current => current with { Theme = preference });
    }

    private async void OnWindowKeyDown(object? sender, KeyEventArgs e) {
        if (e.Key == Key.F9) {
            await ViewModel.Commands["FocusReading"].ExecuteAsync();
            e.Handled = true;
            return;
        }
        if (e.Key == Key.Escape && ViewModel.IsFocusReading) {
            ViewModel.IsFocusReading = false;
            e.Handled = true;
            return;
        }
        bool primaryModifier = e.KeyModifiers.HasFlag(OperatingSystem.IsMacOS() ? KeyModifiers.Meta : KeyModifiers.Control);
        if (primaryModifier && e.KeyModifiers.HasFlag(KeyModifiers.Shift) && e.Key == Key.P) {
            await ShowCommandPaletteAsync();
            e.Handled = true;
            return;
        }
        if (primaryModifier && e.Key == Key.F) {
            FocusDocumentSearch();
            e.Handled = true;
            return;
        }

        if (primaryModifier && e.Key == Key.Tab) {
            TabHost.SelectRelativeTab(e.KeyModifiers.HasFlag(KeyModifiers.Shift));
            e.Handled = true;
            return;
        }
        if (primaryModifier && e.Key == Key.W) {
            await TabHost.CloseSelectedTabAsync();
            e.Handled = true;
            return;
        }
        if (primaryModifier && e.Key == Key.O) {
            await ViewModel.Commands["Open"].ExecuteAsync();
            e.Handled = true;
            return;
        }

        if (primaryModifier && e.Key == Key.S) {
            await ViewModel.Commands[e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? "SaveAs" : "Save"].ExecuteAsync();
            e.Handled = true;
            return;
        }

        if (primaryModifier && e.Key == Key.P) {
            await ViewModel.Commands["Print"].ExecuteAsync();
            e.Handled = true;
            return;
        }

        if (IsTextEntryFocused()) return;

        if (primaryModifier && e.Key == Key.Z) {
            await ViewModel.Commands[e.KeyModifiers.HasFlag(KeyModifiers.Shift) ? "Redo" : "Undo"].ExecuteAsync();
            e.Handled = true;
            return;
        }

        if (primaryModifier) {
            switch (e.Key) {
                case Key.D0:
                case Key.NumPad0:
                    ViewModel.Commands["FitPage"].Execute(null);
                    e.Handled = true;
                    return;
                case Key.D1:
                case Key.NumPad1:
                    ViewModel.Commands["ActualSize"].Execute(null);
                    e.Handled = true;
                    return;
                case Key.D2:
                case Key.NumPad2:
                    ViewModel.Commands["FitWidth"].Execute(null);
                    e.Handled = true;
                    return;
                case Key.OemPlus:
                case Key.Add:
                    ViewModel.Commands["ZoomIn"].Execute(null);
                    e.Handled = true;
                    return;
                case Key.OemMinus:
                case Key.Subtract:
                    ViewModel.Commands["ZoomOut"].Execute(null);
                    e.Handled = true;
                    return;
            }
        }

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

    internal void OpenInitialDocument(string[]? args) {
        string? candidate = args?.FirstOrDefault(static argument => !string.IsNullOrWhiteSpace(argument));
        if (candidate is null) return;
        try {
            _initialDocumentPath = System.IO.Path.GetFullPath(candidate);
        } catch (Exception) when (candidate.Length > 0) {
            _initialDocumentPath = candidate;
        }
    }

    private async void OnOpened(object? sender, EventArgs e) {
        ViewModel.SetViewportSize(PagesList.Bounds.Width, PagesList.Bounds.Height);
        try {
            var cleanup = await _services.Recovery.CleanupExpiredAsync();
            if (cleanup.FailedFiles > 0) _services.Diagnostics.Write(StudioDiagnosticLevel.Warning, "Recovery", "CleanupIncomplete");
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
            _services.Diagnostics.Write(StudioDiagnosticLevel.Warning, "Recovery", "CleanupFailed", error);
        }
        if (_windowClosed) return;
        await _session.InspectAsync();
        if (_windowClosed) return;
        if (_initialDocumentOpened || string.IsNullOrWhiteSpace(_initialDocumentPath)) return;
        _initialDocumentOpened = true;
        await TabHost.OpenDocumentAsync(_initialDocumentPath);
    }

    private async Task<string?> PickFileSafelyAsync(Func<CancellationToken, Task<string?>> picker, CancellationToken token) {
        try { return await picker(token); }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { return null; }
        catch (Exception error) when (error is not OutOfMemoryException) {
            if (!_windowClosed) ViewModel.ErrorMessage = error.Message;
            return null;
        }
    }
    private async Task<IReadOnlyList<string>> PickFilesSafelyAsync(
        Func<CancellationToken, Task<IReadOnlyList<string>>> picker, CancellationToken token) {
        try { return await picker(token); }
        catch (OperationCanceledException) when (token.IsCancellationRequested) { return Array.Empty<string>(); }
        catch (Exception error) when (error is not OutOfMemoryException) {
            if (!_windowClosed) ViewModel.ErrorMessage = error.Message;
            return Array.Empty<string>();
        }
    }

    private async Task<string?> PickPdfAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!StorageProvider.CanOpen) return null;

        IReadOnlyList<IStorageFile> files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = _services.Localizer.Get("Picker.OpenPdf"),
            AllowMultiple = false,
            FileTypeFilter = [
                new FilePickerFileType(_services.Localizer.Get("Picker.PdfDocuments")) {
                    Patterns = ["*.pdf"],
                    MimeTypes = ["application/pdf"],
                    AppleUniformTypeIdentifiers = ["com.adobe.pdf"]
                }
            ]
        });
        return await _services.Storage.RegisterSingleAsync(files, cancellationToken).ConfigureAwait(true);
    }

    private async Task<IReadOnlyList<string>> PickPdfsAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!StorageProvider.CanOpen) return Array.Empty<string>();

        IReadOnlyList<IStorageFile> files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = _services.Localizer.Get("Picker.AddPdfDocuments"),
            AllowMultiple = true,
            FileTypeFilter = [
                new FilePickerFileType(_services.Localizer.Get("Picker.PdfDocuments")) {
                    Patterns = ["*.pdf"],
                    MimeTypes = ["application/pdf"],
                    AppleUniformTypeIdentifiers = ["com.adobe.pdf"]
                }
            ]
        });
        return await _services.Storage.RegisterManyAsync(files, cancellationToken).ConfigureAwait(true);
    }

    private async Task<IReadOnlyList<string>> PickWorkflowFilesAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!StorageProvider.CanOpen) return Array.Empty<string>();

        IReadOnlyList<IStorageFile> files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = _services.Localizer.Get("Picker.AddDocumentsOrImages"),
            AllowMultiple = true,
            FileTypeFilter = [
                new FilePickerFileType(_services.Localizer.Get("Picker.SupportedFiles")) {
                    Patterns = [
                        "*.docx", "*.xlsx", "*.pptx", "*.pdf", "*.html", "*.htm",
                        "*.png", "*.jpg", "*.jpeg", "*.gif", "*.bmp", "*.tif", "*.tiff",
                        "*.webp", "*.ico", "*.pcx", "*.zip"
                    ]
                }
            ]
        });
        return await _services.Storage.RegisterManyAsync(files, cancellationToken).ConfigureAwait(true);
    }

    private async Task<string?> PickAssemblyFolderAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!StorageProvider.CanPickFolder) return null;
        IReadOnlyList<IStorageFolder> folders = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions {
            Title = _services.Localizer.Get("Picker.AddSourceFolder"),
            AllowMultiple = false
        });
        cancellationToken.ThrowIfCancellationRequested();
        return folders.FirstOrDefault()?.Path.LocalPath;
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
        new ActiveOperationsDialog(documents, _services.Localizer, wholeWindow ? TabHost : null, wholeWindow ? _session : null).ShowDialog<bool>(this);

    private async Task<UnsavedChangesDecision> ConfirmUnsavedChangesAsync() {
        var dialog = new UnsavedChangesDialog(ViewModel.DocumentName.TrimEnd(' ', '*'), _services.Localizer);
        return await dialog.ShowDialog<UnsavedChangesDecision>(this);
    }

    private async Task<bool> ConfirmPageDeletionAsync(int pageCount) {
        var dialog = new PageDeletionDialog(pageCount, _services.Localizer);
        return await dialog.ShowDialog<bool>(this);
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
        if (!StorageProvider.CanSave) return null;
        IStorageFile? file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions {
            Title = _services.Localizer.Get("Picker.SavePdf"),
            SuggestedFileName = System.IO.Path.GetFileNameWithoutExtension(ViewModel.DocumentName.TrimEnd(' ', '*')),
            DefaultExtension = "pdf",
            FileTypeChoices = [
                new FilePickerFileType(_services.Localizer.Get("Picker.PdfDocuments")) {
                    Patterns = ["*.pdf"],
                    MimeTypes = ["application/pdf"],
                    AppleUniformTypeIdentifiers = ["com.adobe.pdf"]
                }
            ]
        });
        string? location = await _services.Storage.RegisterSingleAsync(file is null ? [] : [file], cancellationToken).ConfigureAwait(true);
        return location is not null && await ConfirmProviderWriteAsync(location) ? location : null;
    }

    private Task<bool> ConfirmProviderWriteAsync(string location) =>
        !_services.Storage.UsesProviderPublication(location) ? Task.FromResult(true)
            : new ProviderSaveDialog(_services.Storage.Describe(location).Name, _services.Localizer).ShowDialog<bool>(this);
    private async Task<string?> PickOutputFolderAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!StorageProvider.CanOpen) return null;
        IReadOnlyList<IStorageFolder> folders = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions {
            Title = _services.Localizer.Get("Picker.ChooseOutputFolder"),
            AllowMultiple = false
        });
        cancellationToken.ThrowIfCancellationRequested();
        return folders.FirstOrDefault()?.Path.LocalPath;
    }

    private async Task<byte[]?> PickImageAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!StorageProvider.CanOpen) return null;
        IReadOnlyList<IStorageFile> files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = _services.Localizer.Get("Picker.ChooseImage"),
            AllowMultiple = false,
            FileTypeFilter = [
                new FilePickerFileType(_services.Localizer.Get("Picker.PngOrJpegImages")) {
                    Patterns = ["*.png", "*.jpg", "*.jpeg"],
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

    private void OnOrganizerPointerPressed(object? sender, PointerPressedEventArgs e) {
        if (!e.GetCurrentPoint(OrganizerList).Properties.IsLeftButtonPressed) return;
        _organizerDragPage = FindOrganizerPage(e.Source);
        if (_organizerDragPage is null) return;
        ViewModel.NavigateToOrganizerPage(_organizerDragPage.PageNumber);
        _organizerDragPress = e;
        _organizerDragStart = e.GetPosition(OrganizerList);
        _organizerDragStarted = false;
    }

    private async void OnOrganizerPointerMoved(object? sender, PointerEventArgs e) {
        if (_organizerDragStarted || _organizerDragPress is null || _organizerDragPage is null ||
            !e.GetCurrentPoint(OrganizerList).Properties.IsLeftButtonPressed) return;
        Point current = e.GetPosition(OrganizerList);
        if (Math.Abs(current.X - _organizerDragStart.X) < 6D && Math.Abs(current.Y - _organizerDragStart.Y) < 6D) return;

        _organizerDragStarted = true;
        var transfer = new DataTransfer();
        transfer.Add(DataTransferItem.Create(
            OrganizerPageFormat,
            _organizerDragPage.PageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture)));
        PointerPressedEventArgs press = _organizerDragPress;
        ClearOrganizerDrag();
        await DragDrop.DoDragDropAsync(press, transfer, DragDropEffects.Move);
    }

    private void OnOrganizerPointerReleased(object? sender, PointerReleasedEventArgs e) => ClearOrganizerDrag();

    private void OnOrganizerKeyDown(object? sender, KeyEventArgs e) {
        if (e.Key is not (Key.Enter or Key.Space)) return;
        PdfOrganizerPageViewModel? page = FindOrganizerPage(e.Source)
            ?? OrganizerList.SelectedItem as PdfOrganizerPageViewModel;
        if (page is null) return;
        ViewModel.NavigateToOrganizerPage(page.PageNumber);
    }

    private void OnOrganizerDragOver(object? sender, DragEventArgs e) {
        e.DragEffects = TryGetOrganizerPage(e, out _) && FindOrganizerPage(e.Source) is not null
            ? DragDropEffects.Move
            : DragDropEffects.None;
        e.Handled = true;
    }

    private async void OnOrganizerDrop(object? sender, DragEventArgs e) {
        e.Handled = true;
        PdfOrganizerPageViewModel? target = FindOrganizerPage(e.Source);
        if (target is not null && TryGetOrganizerPage(e, out int draggedPage)) {
            await ViewModel.ReorderByDropAsync(draggedPage, target.PageNumber);
        }
    }

    private static bool TryGetOrganizerPage(DragEventArgs e, out int pageNumber) {
        foreach (IDataTransferItem item in e.DataTransfer.Items) {
            string? value = item.TryGetValue(OrganizerPageFormat);
            if (int.TryParse(value, System.Globalization.NumberStyles.None, System.Globalization.CultureInfo.InvariantCulture, out pageNumber)) {
                return true;
            }
        }
        pageNumber = 0;
        return false;
    }

    private static PdfOrganizerPageViewModel? FindOrganizerPage(object? source) {
        Control? control = source as Control;
        while (control is not null) {
            if (control.DataContext is PdfOrganizerPageViewModel page) return page;
            control = control.Parent as Control;
        }
        return null;
    }

    private void ClearOrganizerDrag() {
        _organizerDragPress = null;
        _organizerDragPage = null;
        _organizerDragStarted = false;
    }

    private void OnDragOver(object? sender, DragEventArgs e) {
        e.DragEffects = ViewModel.CanStartDocumentTransition && GetDroppedPdf(e) is not null
            ? DragDropEffects.Copy
            : DragDropEffects.None;
        e.Handled = true;
    }

    private async void OnDrop(object? sender, DragEventArgs e) {
        e.Handled = true;
        if (GetDroppedPdf(e) is not { } file) return;
        try {
            string location = await _services.Storage.RegisterAsync(file, CancellationToken.None);
            if (!_windowClosed) await TabHost.OpenDocumentAsync(location);
        } catch (Exception error) when (error is not OutOfMemoryException) {
            if (!_windowClosed) ViewModel.ErrorMessage = error.Message;
        }
    }

    private static IStorageFile? GetDroppedPdf(DragEventArgs e) => e.DataTransfer.TryGetFiles()?
        .OfType<IStorageFile>().FirstOrDefault(file => string.Equals(System.IO.Path.GetExtension(file.Name), ".pdf", StringComparison.OrdinalIgnoreCase));

    internal static bool TryGetPdfPath(IEnumerable<string?>? candidates, out string? path) {
        path = candidates?.FirstOrDefault(static candidate =>
            !string.IsNullOrWhiteSpace(candidate) &&
            string.Equals(System.IO.Path.GetExtension(candidate), ".pdf", StringComparison.OrdinalIgnoreCase));
        return path is not null;
    }
}
