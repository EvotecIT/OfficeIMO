using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeIMO.Studio.Infrastructure.Preferences;

/// <summary>Persists Studio preferences atomically in the current user's application-data directory.</summary>
internal sealed class JsonStudioPreferencesStore : IStudioPreferencesStore {
    private static readonly JsonSerializerOptions SerializerOptions = new() {
        WriteIndented = true,
        Converters = { new JsonStringEnumConverter() }
    };
    private readonly string _path;

    internal JsonStudioPreferencesStore(string path) {
        _path = Path.GetFullPath(path ?? throw new ArgumentNullException(nameof(path)));
    }

    public StudioPreferences Load() {
        try {
            if (!File.Exists(_path)) return new StudioPreferences();
            StudioPreferences? preferences = JsonSerializer.Deserialize<StudioPreferences>(File.ReadAllText(_path), SerializerOptions);
            return (preferences ?? new StudioPreferences()).Normalize();
        } catch (JsonException) {
            return new StudioPreferences();
        } catch (IOException) {
            return new StudioPreferences();
        } catch (UnauthorizedAccessException) {
            return new StudioPreferences();
        }
    }

    public void Save(StudioPreferences preferences) {
        ArgumentNullException.ThrowIfNull(preferences);
        string? directory = Path.GetDirectoryName(_path);
        if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
        string temporaryPath = _path + ".tmp";
        File.WriteAllText(temporaryPath, JsonSerializer.Serialize(preferences.Normalize(), SerializerOptions));
        File.Move(temporaryPath, _path, overwrite: true);
    }
}
