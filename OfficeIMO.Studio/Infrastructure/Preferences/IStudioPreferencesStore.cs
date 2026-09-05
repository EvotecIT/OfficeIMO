namespace OfficeIMO.Studio.Infrastructure.Preferences;

internal interface IStudioPreferencesStore {
    StudioPreferences Load();

    void Save(StudioPreferences preferences);
}
