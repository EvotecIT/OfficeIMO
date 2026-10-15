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
        Update(_ => preferences);
    }

    public StudioPreferences Update(Func<StudioPreferences, StudioPreferences> update) {
        ArgumentNullException.ThrowIfNull(update);
        string identity = OfficeIMO.Internal.OfficeStorageIdentity.GetPersistenceKey(_path);
        string key = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(identity)));
        using var mutex = new Mutex(false, "OfficeIMO.Studio.Preferences." + key);
        bool entered = false;
        try {
            try { entered = mutex.WaitOne(TimeSpan.FromSeconds(5)); }
            catch (AbandonedMutexException) { entered = true; }
            if (!entered) throw new IOException("Studio preferences are being changed by another window. Try again.");
            StudioPreferences current = Load();
            StudioPreferences next = update(current).Normalize();
            if (next != current) SaveCore(next);
            return next;
        } finally {
            if (entered) mutex.ReleaseMutex();
        }
    }

    private void SaveCore(StudioPreferences preferences) {
        string? directory = Path.GetDirectoryName(_path);
        if (!string.IsNullOrWhiteSpace(directory)) Directory.CreateDirectory(directory);
        // A unique staging name keeps concurrent writers (another process or window) from sharing one temporary file.
        string temporaryPath = _path + "." + Guid.NewGuid().ToString("N") + ".tmp";
        try {
            var options = new FileStreamOptions { Mode = FileMode.CreateNew, Access = FileAccess.Write, Share = FileShare.None };
            if (!OperatingSystem.IsWindows()) options.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;
            using (var stream = new FileStream(temporaryPath, options)) {
                JsonSerializer.Serialize(stream, preferences.Normalize(), SerializerOptions);
            }
            File.Move(temporaryPath, _path, overwrite: true);
        } finally {
            try { if (File.Exists(temporaryPath)) File.Delete(temporaryPath); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException) { }
        }
    }
}
