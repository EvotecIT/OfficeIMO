using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;

namespace OfficeIMO.Studio.Features.Shell;

public sealed partial class MainWindow : Window {
    private string? _initialDocumentPath;
    private bool _initialDocumentOpened;

    public MainWindow() {
        InitializeComponent();
        ViewModel = new MainWindowViewModel(PickPdfAsync);
        DataContext = ViewModel;

        AddHandler(DragDrop.DragOverEvent, OnDragOver);
        AddHandler(DragDrop.DropEvent, OnDrop);
        PageViewport.SizeChanged += (_, _) =>
            ViewModel.SetViewportSize(PageViewport.Bounds.Width, PageViewport.Bounds.Height);
        PagesList.SelectionChanged += (_, _) => {
            if (ViewModel.SelectedPage is not null) {
                PagesList.ScrollIntoView(ViewModel.SelectedPage);
            }
        };
        Opened += OnOpened;
        Closed += (_, _) => ViewModel.Dispose();
    }

    internal MainWindowViewModel ViewModel { get; }

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
        ViewModel.SetViewportSize(PageViewport.Bounds.Width, PageViewport.Bounds.Height);
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

    private void OnDragOver(object? sender, DragEventArgs e) {
        e.DragEffects = TryGetPdfPath(e, out _)
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
