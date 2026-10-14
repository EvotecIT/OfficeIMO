namespace OfficeIMO.Studio.Infrastructure.Preferences;

internal interface IStudioPreferencesStore {
    StudioPreferences Load();

    void Save(StudioPreferences preferences);

    /// <summary>Applies an intentional change to the latest stored values.</summary>
    StudioPreferences Update(Func<StudioPreferences, StudioPreferences> update) {
        StudioPreferences current = Load();
        StudioPreferences next = update(current).Normalize();
        if (next != current) Save(next);
        return next;
    }
}
