namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Provides the platform-local, user-scoped storage locations owned by OfficeIMO Studio.</summary>
internal sealed class StudioDataPaths {
    internal StudioDataPaths(string root) {
        Root = Path.GetFullPath(root ?? throw new ArgumentNullException(nameof(root)));
    }

    internal string Root { get; }

    internal string PreferencesPath => Path.Combine(Root, "preferences.json");

    internal string RecentDocumentsPath => Path.Combine(Root, "recent-documents.json");

    internal string RecoveryRoot => Path.Combine(Root, "Recovery");

    internal string DiagnosticsRoot => Path.Combine(Root, "Diagnostics");

    internal static StudioDataPaths CreateDefault() {
        string localApplicationData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        return new StudioDataPaths(Path.Combine(localApplicationData, "OfficeIMO", "Studio"));
    }
}
