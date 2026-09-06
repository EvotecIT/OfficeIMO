using System.Collections.ObjectModel;
using System.ComponentModel;
using CommunityToolkit.Mvvm.ComponentModel;
using OfficeIMO.Studio.Infrastructure.Localization;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>UI-thread-owned, bounded session history and application workflow resource budget.</summary>
public sealed class StudioJobHistory : ObservableObject {
    public const int MaximumEntries = 500;
    public const int MaximumConcurrentRuns = 2;
    private readonly ObservableCollection<StudioJobRecord> _entries = new();
    private readonly SemaphoreSlim _execution = new(MaximumConcurrentRuns, MaximumConcurrentRuns);
    private readonly IStudioLocalizer _localizer;
    internal StudioJobHistory(IStudioLocalizer localizer) {
        _localizer = localizer;
        Entries = new ReadOnlyObservableCollection<StudioJobRecord>(_entries);
    }
    public ReadOnlyObservableCollection<StudioJobRecord> Entries { get; }
    public int ActiveCount => _entries.Count(entry => entry.IsActive);
    public bool HasEntries => _entries.Count > 0;
    public bool CanClear => _entries.Any(entry => !entry.IsActive);
    public string Summary => _localizer.FormatOrDefault("Jobs.Summary", "{0:N0} active · {1:N0} finished · this session", ActiveCount, _entries.Count - ActiveCount);

    internal StudioJobRecord Start(string title, string input, string? destination, Action cancel, bool batch = false) {
        while (_entries.Count >= MaximumEntries) {
            StudioJobRecord? oldest = _entries.LastOrDefault(entry => !entry.IsActive);
            if (oldest is null) throw new InvalidOperationException("The active job limit has been reached. Wait for a job to finish before starting more work.");
            Remove(oldest);
        }
        var entry = new StudioJobRecord(title, input, destination, cancel, batch, _localizer);
        entry.PropertyChanged += OnEntryChanged;
        _entries.Insert(0, entry);
        NotifyState();
        return entry;
    }

    internal async Task<IDisposable> EnterAsync(CancellationToken cancellationToken) {
        await _execution.WaitAsync(cancellationToken).ConfigureAwait(false);
        return new Permit(_execution);
    }

    internal void ClearFinished() {
        foreach (StudioJobRecord entry in _entries.Where(entry => !entry.IsActive).ToArray()) Remove(entry);
        NotifyState();
    }
    private void Remove(StudioJobRecord entry) {
        entry.PropertyChanged -= OnEntryChanged;
        _entries.Remove(entry);
    }
    private void OnEntryChanged(object? sender, PropertyChangedEventArgs args) {
        if (args.PropertyName == nameof(StudioJobRecord.IsActive)) NotifyState();
    }
    private void NotifyState() {
        OnPropertyChanged(nameof(ActiveCount));
        OnPropertyChanged(nameof(HasEntries));
        OnPropertyChanged(nameof(CanClear));
        OnPropertyChanged(nameof(Summary));
    }
    private sealed class Permit(SemaphoreSlim semaphore) : IDisposable {
        private SemaphoreSlim? _semaphore = semaphore;
        public void Dispose() => Interlocked.Exchange(ref _semaphore, null)?.Release();
    }
}
