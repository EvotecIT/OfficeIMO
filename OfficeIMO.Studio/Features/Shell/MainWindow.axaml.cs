using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Platform.Storage;
using Avalonia.Styling;
using OfficeIMO.Studio.Features.Organizer;

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

    public MainWindow() {
        InitializeComponent();
        ViewModel = new MainWindowViewModel(
            pickPdf: PickPdfAsync,
            pickSavePdf: PickSavePdfAsync,
            pickImportPdfs: PickPdfsAsync,
            pickOutputFolder: PickOutputFolderAsync,
            openUri: OpenUriAsync,
            confirmUnsavedChanges: ConfirmUnsavedChangesAsync,
            pickImage: PickImageAsync,
            confirmPageDeletion: ConfirmPageDeletionAsync,
            pickWorkflowFiles: PickWorkflowFilesAsync);
        DataContext = ViewModel;

        SizeChanged += OnWindowSizeChanged;
        AddHandler(DragDrop.DragOverEvent, OnDragOver);
        AddHandler(DragDrop.DropEvent, OnDrop);
        PagesList.SizeChanged += (_, _) =>
            ViewModel.SetViewportSize(PagesList.Bounds.Width, PagesList.Bounds.Height);
        PagesList.SelectionChanged += (_, _) => {
            if (ViewModel.SelectedPage is not null) {
                PagesList.ScrollIntoView(ViewModel.SelectedPage);
            }
        };
        OrganizerList.SelectionChanged += (_, eventArgs) => {
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
        Closed += (_, _) => ViewModel.Dispose();
    }

    internal MainWindowViewModel ViewModel { get; }

    internal bool IsCompactLayout { get; private set; }

    internal bool AreFitShortcutsVisible => FitWidthButton.IsVisible && FitPageButton.IsVisible;

    internal bool IsConversionCompact => ConversionView.IsCompactLayout;

    internal bool IsDocumentHealthCompact => DocumentHealthView.IsCompactLayout;

    private void OnWindowSizeChanged(object? sender, SizeChangedEventArgs e) => ApplyResponsiveLayout(e.NewSize.Width);

    internal void ApplyResponsiveLayout(double width) {
        IsCompactLayout = width < 1180D;
        DocumentNameText.MaxWidth = IsCompactLayout ? 150D : 230D;
        SelectedPagePositionText.MinWidth = IsCompactLayout ? 78D : 110D;
        FitWidthButton.IsVisible = !IsCompactLayout;
        FitPageButton.IsVisible = !IsCompactLayout;
        double workspaceWidth = Math.Max(0D, width - 204D);
        ConversionView.ApplyResponsiveLayout(workspaceWidth);
        DocumentHealthView.ApplyResponsiveLayout(workspaceWidth);
    }

    private void OnToggleThemeClick(object? sender, Avalonia.Interactivity.RoutedEventArgs e) {
        if (Application.Current is not { } application) return;
        application.RequestedThemeVariant = application.ActualThemeVariant == ThemeVariant.Dark
            ? ThemeVariant.Light
            : ThemeVariant.Dark;
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
        if (_initialDocumentOpened || string.IsNullOrWhiteSpace(_initialDocumentPath)) return;
        _initialDocumentOpened = true;
        await ViewModel.OpenDocumentAsync(_initialDocumentPath);
    }

    private async Task<string?> PickPdfAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!StorageProvider.CanOpen) return null;

        IReadOnlyList<IStorageFile> files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = "Open PDF",
            AllowMultiple = false,
            FileTypeFilter = [
                new FilePickerFileType("PDF documents") {
                    Patterns = ["*.pdf"],
                    MimeTypes = ["application/pdf"],
                    AppleUniformTypeIdentifiers = ["com.adobe.pdf"]
                }
            ]
        });
        cancellationToken.ThrowIfCancellationRequested();
        return files.FirstOrDefault()?.Path.LocalPath;
    }

    private async Task<IReadOnlyList<string>> PickPdfsAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!StorageProvider.CanOpen) return Array.Empty<string>();

        IReadOnlyList<IStorageFile> files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = "Add PDF documents",
            AllowMultiple = true,
            FileTypeFilter = [
                new FilePickerFileType("PDF documents") {
                    Patterns = ["*.pdf"],
                    MimeTypes = ["application/pdf"],
                    AppleUniformTypeIdentifiers = ["com.adobe.pdf"]
                }
            ]
        });
        cancellationToken.ThrowIfCancellationRequested();
        return files.Select(static file => file.Path.LocalPath).ToArray();
    }

    private async Task<IReadOnlyList<string>> PickWorkflowFilesAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!StorageProvider.CanOpen) return Array.Empty<string>();

        IReadOnlyList<IStorageFile> files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = "Add documents to conversion queue",
            AllowMultiple = true,
            FileTypeFilter = [
                new FilePickerFileType("Office, PDF, and HTML documents") {
                    Patterns = ["*.docx", "*.xlsx", "*.pptx", "*.pdf", "*.html", "*.htm"]
                }
            ]
        });
        cancellationToken.ThrowIfCancellationRequested();
        return files.Select(static file => file.Path.LocalPath).ToArray();
    }

    private async void OnClosing(object? sender, WindowClosingEventArgs e) {
        if (_allowClose) return;
        if (ViewModel.CanCancelOperation) {
            e.Cancel = true;
            ViewModel.CancelCurrentOperation();
            return;
        }
        if (!ViewModel.IsDirty) return;
        e.Cancel = true;
        if (_closePromptOpen) return;
        _closePromptOpen = true;
        try {
            if (!await ViewModel.RequestCloseDocumentAsync()) return;
            _allowClose = true;
            Close();
        } finally {
            _closePromptOpen = false;
        }
    }

    private async Task<UnsavedChangesDecision> ConfirmUnsavedChangesAsync() {
        var dialog = new UnsavedChangesDialog(ViewModel.DocumentName.TrimEnd(' ', '*'));
        return await dialog.ShowDialog<UnsavedChangesDecision>(this);
    }

    private async Task<bool> ConfirmPageDeletionAsync(int pageCount) {
        var dialog = new PageDeletionDialog(pageCount);
        return await dialog.ShowDialog<bool>(this);
    }

    private async Task<string?> PickSavePdfAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!StorageProvider.CanSave) return null;
        IStorageFile? file = await StorageProvider.SaveFilePickerAsync(new FilePickerSaveOptions {
            Title = "Save PDF",
            SuggestedFileName = System.IO.Path.GetFileNameWithoutExtension(ViewModel.DocumentName.TrimEnd(' ', '*')),
            DefaultExtension = "pdf",
            FileTypeChoices = [
                new FilePickerFileType("PDF documents") {
                    Patterns = ["*.pdf"],
                    MimeTypes = ["application/pdf"],
                    AppleUniformTypeIdentifiers = ["com.adobe.pdf"]
                }
            ]
        });
        cancellationToken.ThrowIfCancellationRequested();
        return file?.Path.LocalPath;
    }

    private async Task<string?> PickOutputFolderAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!StorageProvider.CanOpen) return null;
        IReadOnlyList<IStorageFolder> folders = await StorageProvider.OpenFolderPickerAsync(new FolderPickerOpenOptions {
            Title = "Choose output folder",
            AllowMultiple = false
        });
        cancellationToken.ThrowIfCancellationRequested();
        return folders.FirstOrDefault()?.Path.LocalPath;
    }

    private async Task<string?> PickImageAsync(CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (!StorageProvider.CanOpen) return null;
        IReadOnlyList<IStorageFile> files = await StorageProvider.OpenFilePickerAsync(new FilePickerOpenOptions {
            Title = "Choose image",
            AllowMultiple = false,
            FileTypeFilter = [
                new FilePickerFileType("PNG or JPEG images") {
                    Patterns = ["*.png", "*.jpg", "*.jpeg"],
                    MimeTypes = ["image/png", "image/jpeg"],
                    AppleUniformTypeIdentifiers = ["public.png", "public.jpeg"]
                }
            ]
        });
        cancellationToken.ThrowIfCancellationRequested();
        return files.FirstOrDefault()?.Path.LocalPath;
    }

    private async Task OpenUriAsync(Uri uri) {
        bool opened = await Launcher.LaunchUriAsync(uri);
        if (!opened) throw new InvalidOperationException("The operating system could not open this link.");
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
        e.DragEffects = ViewModel.CanStartDocumentTransition && TryGetPdfPath(e, out _)
            ? DragDropEffects.Copy
            : DragDropEffects.None;
        e.Handled = true;
    }

    private async void OnDrop(object? sender, DragEventArgs e) {
        e.Handled = true;
        if (TryGetPdfPath(e, out string? path) && path is not null) {
            await ViewModel.OpenDocumentAsync(path);
        }
    }

    private static bool TryGetPdfPath(DragEventArgs e, out string? path) {
        IEnumerable<string?>? candidates = e.DataTransfer
            .TryGetFiles()?
            .Select(static item => item.Path.LocalPath);
        return TryGetPdfPath(candidates, out path);
    }

    internal static bool TryGetPdfPath(IEnumerable<string?>? candidates, out string? path) {
        path = candidates?.FirstOrDefault(static candidate =>
            !string.IsNullOrWhiteSpace(candidate) &&
            string.Equals(System.IO.Path.GetExtension(candidate), ".pdf", StringComparison.OrdinalIgnoreCase));
        return path is not null;
    }
}
