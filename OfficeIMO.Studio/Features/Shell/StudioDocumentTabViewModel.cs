using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;

namespace OfficeIMO.Studio.Features.Shell;

/// <summary>Represents one live document workspace in the desktop tab strip.</summary>
public sealed partial class StudioDocumentTabViewModel : ObservableObject, IDisposable {
    private readonly Func<StudioDocumentTabViewModel, Task> _close;
    private bool _disposed;

    internal StudioDocumentTabViewModel(
        MainWindowViewModel document,
        Func<StudioDocumentTabViewModel, Task> close) {
        Document = document ?? throw new ArgumentNullException(nameof(document));
        _close = close ?? throw new ArgumentNullException(nameof(close));
        _title = document.DocumentName;
        Document.PropertyChanged += OnDocumentPropertyChanged;
    }

    internal MainWindowViewModel Document { get; }

    [ObservableProperty]
    private string _title;

    [RelayCommand]
    private Task CloseAsync() => _close(this);

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        Document.PropertyChanged -= OnDocumentPropertyChanged;
        Document.Dispose();
    }

    private void OnDocumentPropertyChanged(object? sender, PropertyChangedEventArgs e) {
        if (e.PropertyName == nameof(MainWindowViewModel.DocumentName)) Title = Document.DocumentName;
    }
}
