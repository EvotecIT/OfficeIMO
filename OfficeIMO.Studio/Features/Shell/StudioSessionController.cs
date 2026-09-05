using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Security.Cryptography;
using Avalonia.Threading;
using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Diagnostics;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Features.Shell;

internal sealed partial class StudioSessionItem(StudioSessionDocument document) : ObservableObject {
    internal StudioSessionDocument Document { get; } = document;
    public string Name => Path.GetFileName(Document.Path);
    public string SourcePath => Document.Path;
    [ObservableProperty] private string _status = string.Empty;
    [ObservableProperty] private bool _sourceUnchanged;
    [ObservableProperty] private bool _sourceExists;
    [ObservableProperty] private bool _hasRecovery;
    public override string ToString() => Name;
}

/// <summary>Coordinates explicit restart choices and debounced host-state persistence.</summary>
internal sealed partial class StudioSessionController : ObservableObject, IDisposable {
    private readonly StudioDocumentTabHost _host;
    private readonly StudioApplicationServices _services;
    private readonly StudioSessionStore _store;
    private readonly PdfWorkspaceRecoveryStore _recovery;
    private readonly Func<CancellationToken, Task<string?>> _pickCopy;
    private readonly HashSet<MainWindowViewModel> _observed = [];
    private readonly DispatcherTimer _saveTimer;
    private readonly string? _previousActivePath;
    private bool _frozen;
    private bool _disposed;

    internal StudioSessionController(StudioDocumentTabHost host, StudioApplicationServices services,
        Func<CancellationToken, Task<string?>> pickCopy) {
        _host = host;
        _services = services;
        _pickCopy = pickCopy;
        _store = new(services.Paths.SessionPath);
        _recovery = new(services.Paths.RecoveryRoot);
        StudioSessionSnapshot previous = services.Preferences.Current.RememberSession ? _store.Load() : new(1, DateTimeOffset.UtcNow, null, []);
        _previousActivePath = previous.ActivePath;
        foreach (var document in previous.Documents) Pending.Add(new(document));
        _saveTimer = new() { Interval = TimeSpan.FromMilliseconds(500) };
        _saveTimer.Tick += (_, _) => { _saveTimer.Stop(); Flush(); };
        _host.Tabs.CollectionChanged += OnTabsChanged;
        _host.PropertyChanged += OnHostChanged;
        _host.CloseAllPrepared += OnCloseAllPrepared;
        _services.Preferences.Changed += OnPreferencesChanged;
        ObserveTabs();
        if (!services.Preferences.Current.RememberSession) Flush();
    }

    public ObservableCollection<StudioSessionItem> Pending { get; } = [];
    public bool HasPending => Pending.Count > 0;
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string? _error;
    public bool HasError => !string.IsNullOrWhiteSpace(Error);
    partial void OnErrorChanged(string? value) => OnPropertyChanged(nameof(HasError));

    internal async Task InspectAsync(CancellationToken token = default) {
        foreach (var item in Pending.ToArray()) {
            if (_disposed) return;
            // A local edit snapshot remains usable when the original cannot be opened.
            item.HasRecovery = _recovery.Find(item.SourcePath, item.Document.Fingerprint) is not null;
            try {
                item.SourceUnchanged = false;
                item.SourceExists = File.Exists(item.SourcePath);
                if (item.SourceExists) {
                    await using var stream = File.OpenRead(item.SourcePath);
                    string fingerprint = Convert.ToHexString(await SHA256.HashDataAsync(stream, token));
                    item.SourceUnchanged = string.Equals(fingerprint, item.Document.Fingerprint, StringComparison.OrdinalIgnoreCase);
                }
                item.Status = Text(item.SourceUnchanged ? "Ready" : item.SourceExists ? "Changed" : "Missing");
            } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
                item.SourceUnchanged = false;
                item.Status = Text("Unavailable");
            }
        }
    }

    [RelayCommand]
    private async Task RestoreAsync() {
        if (IsBusy) return;
        IsBusy = true;
        Error = null;
        try {
            await InspectAsync();
            if (_disposed) return;
            foreach (var item in Pending.Where(item => item.SourceUnchanged).ToArray()) await OpenItemAsync(item);
            var active = _host.Tabs.FirstOrDefault(tab => PathsEqual(tab.Document.DocumentPath, _previousActivePath));
            if (active is not null) _host.SelectedTab = active;
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
            Error = Text("RestoreFailed");
        } finally { IsBusy = false; Flush(); }
    }

    [RelayCommand]
    private async Task OpenCurrentAsync(StudioSessionItem? item) {
        if (item is null || IsBusy) return;
        IsBusy = true;
        Error = null;
        try { await OpenItemAsync(item); }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException) { Error = Text("RestoreFailed"); }
        finally { IsBusy = false; Flush(); }
    }

    private async Task OpenItemAsync(StudioSessionItem item) {
        if (_disposed) return;
        _services.DocumentViews.Put(item.SourcePath, item.Document.View);
        await _host.OpenDocumentAsync(item.SourcePath);
        if (_host.Tabs.Any(tab => PathsEqual(tab.Document.DocumentPath, item.SourcePath))) RemovePending(item);
        else Error = Text("RestoreFailed");
    }

    [RelayCommand]
    private async Task RecoverCopyAsync(StudioSessionItem? item) {
        if (item is null || IsBusy) return;
        IsBusy = true;
        Error = null;
        string? staging = null;
        try {
            string? selected = await _pickCopy(CancellationToken.None);
            if (_disposed || string.IsNullOrWhiteSpace(selected)) return;
            string destination = Path.GetFullPath(selected);
            if (File.Exists(destination) || PathsEqual(destination, item.SourcePath) || !_host.CanPublishPath(destination)) {
                Error = Text("NewCopyRequired");
                return;
            }
            staging = Path.Combine(Path.GetDirectoryName(destination)!, ".recovered-" + Guid.NewGuid().ToString("N") + ".pdf");
            byte[] recovered = _recovery.ReadVerifiedSnapshot(item.SourcePath, item.Document.Fingerprint) ?? throw new IOException();
            PdfDocument.Load(recovered).Save(staging);
            File.Move(staging, destination, overwrite: false);
            _services.DocumentViews.Put(destination, item.Document.View);
            await _host.OpenDocumentAsync(destination);
            if (_host.Tabs.Any(tab => PathsEqual(tab.Document.DocumentPath, destination))) RemovePending(item);
        } catch (Exception error) when (error is not OutOfMemoryException) {
            Error = Text("RecoverFailed");
        } finally {
            try { if (staging is not null && File.Exists(staging)) File.Delete(staging); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
                _services.Diagnostics.Write(StudioDiagnosticLevel.Warning, "Session", "RecoveryStagingCleanupFailed", error);
            }
            IsBusy = false;
            Flush();
        }
    }

    [RelayCommand]
    private void Forget(StudioSessionItem? item) {
        if (item is null || IsBusy) return;
        RemovePending(item);
        Flush();
    }

    private void RemovePending(StudioSessionItem item) { Pending.Remove(item); OnPropertyChanged(nameof(HasPending)); }

    internal void Flush() {
        if (_frozen || _disposed) return;
        try {
            if (!_services.Preferences.Current.RememberSession) { _store.Clear(); return; }
            var live = _host.Tabs.Select(tab => tab.Document.CaptureSessionDocument()).OfType<StudioSessionDocument>();
            _store.Save(new(1, DateTimeOffset.UtcNow, _host.ActiveDocument.DocumentPath ?? _previousActivePath,
                live.Concat(Pending.Select(item => item.Document)).ToArray()));
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
            _services.Diagnostics.Write(StudioDiagnosticLevel.Warning, "Session", "SessionSaveFailed", error);
        }
    }

    internal void CaptureForShutdown() { Flush(); _frozen = true; _saveTimer.Stop(); }
    private void OnCloseAllPrepared(object? sender, EventArgs args) => CaptureForShutdown();
    private void OnPreferencesChanged(object? sender, EventArgs args) {
        if (!_services.Preferences.Current.RememberSession) {
            Pending.Clear();
            OnPropertyChanged(nameof(HasPending));
        }
        Flush();
    }
    private void OnHostChanged(object? sender, PropertyChangedEventArgs args) => ScheduleSave();
    private void OnTabsChanged(object? sender, NotifyCollectionChangedEventArgs args) { ObserveTabs(); ScheduleSave(); }
    private void OnDocumentChanged(object? sender, PropertyChangedEventArgs args) {
        if (args.PropertyName == nameof(MainWindowViewModel.HasDocument) && sender is MainWindowViewModel { HasDocument: true } opened) {
            foreach (var item in Pending.Where(item => PathsEqual(item.SourcePath, opened.DocumentPath)).ToArray()) RemovePending(item);
        }
        if (args.PropertyName is nameof(MainWindowViewModel.SelectedPage) or nameof(MainWindowViewModel.Zoom) or
            nameof(MainWindowViewModel.DocumentPath) or nameof(MainWindowViewModel.HasDocument) or
            nameof(MainWindowViewModel.IsDirty) or nameof(MainWindowViewModel.IsFocusReading)) ScheduleSave();
    }
    private void ScheduleSave() { if (!_frozen && !_disposed) { _saveTimer.Stop(); _saveTimer.Start(); } }
    private void ObserveTabs() {
        var current = _host.Tabs.Select(tab => tab.Document).ToHashSet();
        foreach (var document in _observed.Except(current).ToArray()) { document.PropertyChanged -= OnDocumentChanged; _observed.Remove(document); }
        foreach (var document in current.Except(_observed)) { document.PropertyChanged += OnDocumentChanged; _observed.Add(document); }
    }
    private string Text(string key) => _services.Localizer.Get("Session." + key);
    private static bool PathsEqual(string? first, string? second) => first is not null && second is not null &&
        string.Equals(first, second, MainWindowViewModel.RecentDocumentPathComparison);

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _saveTimer.Stop();
        _host.Tabs.CollectionChanged -= OnTabsChanged;
        _host.PropertyChanged -= OnHostChanged;
        _host.CloseAllPrepared -= OnCloseAllPrepared;
        _services.Preferences.Changed -= OnPreferencesChanged;
        foreach (var document in _observed) document.PropertyChanged -= OnDocumentChanged;
        _observed.Clear();
    }
}
