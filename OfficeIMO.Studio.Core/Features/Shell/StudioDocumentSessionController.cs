using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Security.Cryptography;
using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Pdf;
using OfficeIMO.Internal;
using OfficeIMO.Core.Internal;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Features.Shell;

internal sealed partial class StudioSessionItem(StudioSessionDocument document) : ObservableObject {
    internal StudioSessionDocument Document { get; } = document;
    public string Name => Document.Storage?.Name ?? OfficeStorageIdentity.GetFileName(Document.Path);
    public string SourcePath => Document.Path;
    [ObservableProperty]
    [NotifyPropertyChangedFor(nameof(IsChecking))]
    private string _status = string.Empty;
    /// <summary>True until the source has been inspected; the banner shows a checking state instead of dead actions.</summary>
    public bool IsChecking => string.IsNullOrEmpty(Status);
    public string? Folder => OfficeStorageIdentity.GetLocalPath(Document.Path) is { } local ? System.IO.Path.GetDirectoryName(local) : null;
    [ObservableProperty] private bool _sourceUnchanged;
    [ObservableProperty] private bool _sourceExists;
    [ObservableProperty] private bool _hasRecovery;
    public override string ToString() => Name;
}

/// <summary>Coordinates explicit restart choices and debounced host-state persistence.</summary>
internal partial class StudioDocumentSessionController<TDocument, TTab> : ObservableObject, IDisposable
    where TDocument : class, IStudioDocument
    where TTab : class, IStudioDocumentTab<TDocument> {
    private readonly StudioDocumentTabs<TDocument, TTab> _host;
    private readonly IStudioSessionEnvironment _services;
    private readonly StudioSessionStore _store;
    private readonly PdfWorkspaceRecoveryStore _recovery;
    private readonly Func<CancellationToken, Task<string?>> _pickCopy;
    private readonly HashSet<TDocument> _observed = [];
    private IDisposable? _saveTimer;
    // Background completions belong to this application, even after its dispatcher shuts down.
    private readonly IStudioScheduler _scheduler;
    private string? _previousActivePath;
    private bool _frozen;
    private bool _disposed;
    private CancellationTokenSource? _operationCancellation;

    internal void CancelActiveOperation() => _operationCancellation?.Cancel();

    internal StudioDocumentSessionController(StudioDocumentTabs<TDocument, TTab> host, IStudioSessionEnvironment services,
        Func<CancellationToken, Task<string?>> pickCopy, IStudioScheduler scheduler) {
        _scheduler = scheduler;
        _host = host;
        _services = services;
        _pickCopy = pickCopy;
        _store = services.SessionStore;
        services.SessionCleared += OnSessionCleared;
        _recovery = services.Recovery;
        _recovery.MaintenanceCompleted += OnRecoveryMaintenanceCompleted;
        StudioSessionSnapshot previous = services.RememberSession ? _store.Load() : new(1, DateTimeOffset.UtcNow, null, []);
        _previousActivePath = previous.ActivePath;
        Pending.CollectionChanged += (_, _) => OnPropertyChanged(nameof(PendingSummary));
        foreach (var document in previous.Documents) {
            if (document.Storage is { } reference) services.Storage.Remember(reference);
            Pending.Add(new(document));
        }
        _host.Tabs.CollectionChanged += OnTabsChanged;
        _host.PropertyChanged += OnHostChanged;
        _host.CloseAllPrepared += OnCloseAllPrepared;
        _services.PreferencesChanged += OnPreferencesChanged;
        ObserveTabs();
        if (!services.RememberSession) Flush();
    }

    public ObservableCollection<StudioSessionItem> Pending { get; } = [];
    public bool HasPending => Pending.Count > 0;
    public string PendingSummary => Pending.Count == 1
        ? _services.Text("SummaryOne")
        : _services.Format("Summary", Pending.Count);
    private void OnSessionCleared(object? sender, EventArgs args) {
        _previousActivePath = null;
        Pending.Clear();
        OnPropertyChanged(nameof(HasPending));
    }
    private void OnRecoveryMaintenanceCompleted(object? sender, EventArgs args) => _scheduler.Post(() => {
        if (_disposed) return;
        foreach (var item in Pending) item.HasRecovery = item.HasRecovery && _recovery.HasSnapshotFiles(item.SourcePath);
    });
    [ObservableProperty] private bool _isBusy;
    [ObservableProperty] private string? _error;
    public bool HasError => !string.IsNullOrWhiteSpace(Error);
    partial void OnErrorChanged(string? value) => OnPropertyChanged(nameof(HasError));
    [ObservableProperty] private string? _storageError;
    public bool HasStorageError => !string.IsNullOrWhiteSpace(StorageError);
    partial void OnStorageErrorChanged(string? value) => OnPropertyChanged(nameof(HasStorageError));

    internal async Task InspectAsync(CancellationToken token = default) {
        foreach (var item in Pending.ToArray()) {
            token.ThrowIfCancellationRequested();
            if (_disposed) return;
            // A local edit snapshot remains usable when the original cannot be opened.
            item.HasRecovery = _recovery.Find(item.SourcePath, item.Document.Fingerprint) is not null;
            try {
                item.SourceUnchanged = false;
                item.SourceExists = false;
                string fingerprint = await _services.Storage.FingerprintAsync(item.SourcePath, token);
                item.SourceExists = true;
                item.SourceUnchanged = string.Equals(fingerprint, item.Document.Fingerprint, StringComparison.OrdinalIgnoreCase);
                item.Status = Text(item.SourceUnchanged ? "Ready" : item.SourceExists ? "Changed" : "Missing");
            } catch (Exception error) when (error is FileNotFoundException or DirectoryNotFoundException) {
                item.SourceUnchanged = false;
                item.SourceExists = false;
                item.Status = Text("Missing");
            } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
                item.SourceUnchanged = false;
                item.Status = Text("Unavailable");
            }
        }
    }

    internal async Task RestoreDocumentsAsync() {
        if (IsBusy) return;
        using var cancellation = new CancellationTokenSource();
        _operationCancellation = cancellation;
        IsBusy = true;
        Error = null;
        try {
            await InspectAsync(cancellation.Token);
            if (_disposed) return;
            foreach (var item in Pending.Where(item => item.SourceUnchanged).ToArray()) await OpenItemAsync(item, cancellation.Token);
            var active = _host.Tabs.FirstOrDefault(tab => PathsEqual(tab.Document.DocumentPath, _previousActivePath));
            if (active is not null) _host.SelectedTab = active;
        } catch (OperationCanceledException) when (cancellation.IsCancellationRequested) {
            Error = null;
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
            Error = Text("RestoreFailed");
        } finally { _operationCancellation = null; IsBusy = false; Flush(); }
    }

    internal async Task OpenCurrentDocumentAsync(StudioSessionItem? item) {
        if (item is null || IsBusy) return;
        using var cancellation = new CancellationTokenSource();
        _operationCancellation = cancellation;
        IsBusy = true;
        Error = null;
        try { await OpenItemAsync(item, cancellation.Token); }
        catch (OperationCanceledException) when (cancellation.IsCancellationRequested) { Error = null; }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException) { Error = Text("RestoreFailed"); }
        finally { _operationCancellation = null; IsBusy = false; Flush(); }
    }

    private async Task OpenItemAsync(StudioSessionItem item, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        if (_disposed) return;
        await _host.OpenDocumentAsync(item.SourcePath, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        var opened = _host.Tabs.FirstOrDefault(tab => PathsEqual(tab.Document.DocumentPath, item.SourcePath));
        if (opened is not null) {
            opened.Document.RestoreSessionViewState(item.Document.View);
            RemovePending(item);
        }
        else Error = Text("RestoreFailed");
    }

    internal async Task RecoverDocumentCopyAsync(StudioSessionItem? item) {
        if (item is null || IsBusy) return;
        using var cancellation = new CancellationTokenSource();
        _operationCancellation = cancellation;
        IsBusy = true;
        Error = null;
        string? staging = null;
        try {
            string? selected = await _pickCopy(cancellation.Token);
            cancellation.Token.ThrowIfCancellationRequested();
            if (_disposed || string.IsNullOrWhiteSpace(selected)) return;
            string destination = OfficeStorageIdentity.Normalize(selected);
            if (File.Exists(destination) || PathsEqual(destination, item.SourcePath) || !_host.CanPublishPath(destination)) {
                Error = Text("NewCopyRequired");
                return;
            }
            byte[] recovered = _recovery.ReadVerifiedSnapshot(item.SourcePath, item.Document.Fingerprint) ?? throw new IOException();
            if (_services.Storage.UsesProviderPublication(destination)) {
                using var serialized = new OfficeBoundedMemoryStream(StudioDocumentStorage.MaximumDocumentBytes);
                await PdfDocument.Load(recovered).SaveAsync(serialized, cancellation.Token);
                await _services.Storage.PublishAsync(destination, serialized.ToArray(), null, token => {
                    token.ThrowIfCancellationRequested();
                    if (_disposed || !_host.CanPublishPath(destination) || PathsEqual(destination, item.SourcePath)) {
                        throw new IOException("Choose a different recovery destination.");
                    }
                    return Task.CompletedTask;
                }, cancellation.Token);
            } else {
                staging = Path.Combine(Path.GetDirectoryName(destination)!, ".recovered-" + Guid.NewGuid().ToString("N") + ".pdf");
                PdfDocument.Load(recovered).Save(staging);
                cancellation.Token.ThrowIfCancellationRequested();
                File.Move(staging, destination, overwrite: false);
            }
            await _host.OpenDocumentAsync(destination, cancellation.Token);
            cancellation.Token.ThrowIfCancellationRequested();
            var opened = _host.Tabs.FirstOrDefault(tab => PathsEqual(tab.Document.DocumentPath, destination));
            if (opened is not null) {
                opened.Document.RestoreSessionViewState(item.Document.View);
                RemovePending(item);
            }
        } catch (OperationCanceledException) when (cancellation.IsCancellationRequested) {
            Error = null;
        } catch (Exception error) when (error is not OutOfMemoryException) {
            Error = Text("RecoverFailed");
        } finally {
            try { if (staging is not null && File.Exists(staging)) File.Delete(staging); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
                _services.ReportFailure("RecoveryStagingCleanupFailed", error);
            }
            _operationCancellation = null;
            IsBusy = false;
            Flush();
        }
    }

    internal void ForgetDocument(StudioSessionItem? item) {
        if (item is null || IsBusy) return;
        RemovePending(item);
        Flush();
    }

    private void RemovePending(StudioSessionItem item) { Pending.Remove(item); OnPropertyChanged(nameof(HasPending)); }

    internal void Flush() {
        if (_frozen || _disposed) return;
        try {
            if (!_services.RememberSession) { _store.Clear(); StorageError = null; return; }
            var live = _host.Tabs.Select(tab => tab.Document.CaptureSessionDocument()).OfType<StudioSessionDocument>();
            _store.Save(new(1, DateTimeOffset.UtcNow, _host.ActiveDocument.DocumentPath ?? _previousActivePath,
                live.Concat(Pending.Select(item => item.Document)).ToArray()));
            StorageError = null;
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
            StorageError = Text(_services.RememberSession ? "StorageSaveFailed" : "StorageClearFailed");
            _services.ReportFailure("SessionSaveFailed", error);
        }
    }

    internal void CaptureForShutdown() { Flush(); _frozen = true; _saveTimer?.Dispose(); _saveTimer = null; }
    private void OnCloseAllPrepared(object? sender, EventArgs args) => CaptureForShutdown();
    private void OnPreferencesChanged(object? sender, EventArgs args) {
        if (!_services.RememberSession) {
            _previousActivePath = null;
            Pending.Clear();
            OnPropertyChanged(nameof(HasPending));
        }
        Flush();
    }
    private void OnHostChanged(object? sender, PropertyChangedEventArgs args) => ScheduleSave();
    private void OnTabsChanged(object? sender, NotifyCollectionChangedEventArgs args) { ObserveTabs(); ScheduleSave(); }
    private void OnDocumentChanged(object? sender, StudioDocumentChangedEventArgs args) {
        if (args.Kind == StudioDocumentChangeKind.Document && sender is TDocument { HasDocument: true } opened) {
            foreach (var item in Pending.Where(item => PathsEqual(item.SourcePath, opened.DocumentPath)).ToArray()) RemovePending(item);
        }
        if (args.Kind is StudioDocumentChangeKind.Document or StudioDocumentChangeKind.Content or StudioDocumentChangeKind.View) ScheduleSave();
    }
    private void ScheduleSave() { if (!_frozen && !_disposed) { _saveTimer?.Dispose(); _saveTimer = null; _saveTimer = _scheduler.Schedule(TimeSpan.FromMilliseconds(500), Flush); } }
    private void ObserveTabs() {
        var current = _host.Tabs.Select(tab => tab.Document).ToHashSet();
        foreach (var document in _observed.Except(current).ToArray()) { document.Changed -= OnDocumentChanged; _observed.Remove(document); }
        foreach (var document in current.Except(_observed)) { document.Changed += OnDocumentChanged; _observed.Add(document); }
    }
    private string Text(string key) => _services.Text(key);
    private static bool PathsEqual(string? first, string? second) => first is not null && second is not null &&
        OfficeStorageIdentity.GetPersistenceKey(first) == OfficeStorageIdentity.GetPersistenceKey(second);

    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        CancelActiveOperation();
        _recovery.MaintenanceCompleted -= OnRecoveryMaintenanceCompleted;
        _services.SessionCleared -= OnSessionCleared;
        _saveTimer?.Dispose(); _saveTimer = null;
        _host.Tabs.CollectionChanged -= OnTabsChanged;
        _host.PropertyChanged -= OnHostChanged;
        _host.CloseAllPrepared -= OnCloseAllPrepared;
        _services.PreferencesChanged -= OnPreferencesChanged;
        foreach (var document in _observed) document.Changed -= OnDocumentChanged;
        _observed.Clear();
        _services.Dispose();
    }
}
