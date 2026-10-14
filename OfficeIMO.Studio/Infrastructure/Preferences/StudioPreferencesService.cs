namespace OfficeIMO.Studio.Infrastructure.Preferences;

/// <summary>Owns the current normalized Studio preferences and persists intentional changes.</summary>
internal sealed class StudioPreferencesService {
    private readonly IStudioPreferencesStore _store;

    internal StudioPreferencesService(IStudioPreferencesStore store) {
        _store = store ?? throw new ArgumentNullException(nameof(store));
        Current = _store.Load().Normalize();
    }

    internal event EventHandler? Changed;

    /// <summary>Raised when a presentation change applies to this session but could not be written to disk.</summary>
    internal event EventHandler<Exception>? PersistenceFailed;

    internal StudioPreferences Current { get; private set; }

    /// <summary>Persists and applies a change; a failed write leaves the current preferences unchanged and throws.</summary>
    internal void Update(Func<StudioPreferences, StudioPreferences> update) {
        ArgumentNullException.ThrowIfNull(update);
        StudioPreferences next = _store.Update(update);
        if (next == Current) return;
        Current = next;
        Changed?.Invoke(this, EventArgs.Empty);
    }

    /// <summary>
    /// Applies a presentation-only change (appearance, density, window placement) even when it cannot be saved,
    /// so a locked or read-only profile never undoes the visible choice or crashes the window. Privacy and
    /// data-retention choices must use <see cref="Update"/> so a failed write is reported and rolled back.
    /// </summary>
    internal void UpdatePresentation(Func<StudioPreferences, StudioPreferences> update) {
        ArgumentNullException.ThrowIfNull(update);
        StudioPreferences next;
        Exception? failure = null;
        try {
            next = _store.Update(update);
        } catch (Exception error) when (error is IOException or UnauthorizedAccessException) {
            next = update(Current).Normalize();
            failure = error;
        }
        if (next == Current && failure is null) return;
        Current = next;
        Changed?.Invoke(this, EventArgs.Empty);
        if (failure is not null) PersistenceFailed?.Invoke(this, failure);
    }
}
