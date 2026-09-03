namespace OfficeIMO.Studio.Infrastructure.Preferences;

/// <summary>Owns the current normalized Studio preferences and persists intentional changes.</summary>
internal sealed class StudioPreferencesService {
    private readonly IStudioPreferencesStore _store;

    internal StudioPreferencesService(IStudioPreferencesStore store) {
        _store = store ?? throw new ArgumentNullException(nameof(store));
        Current = _store.Load().Normalize();
    }

    internal event EventHandler? Changed;

    internal StudioPreferences Current { get; private set; }

    internal void Update(Func<StudioPreferences, StudioPreferences> update) {
        ArgumentNullException.ThrowIfNull(update);
        StudioPreferences next = update(Current).Normalize();
        if (next == Current) return;
        _store.Save(next);
        Current = next;
        Changed?.Invoke(this, EventArgs.Empty);
    }
}
