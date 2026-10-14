using OfficeIMO.Studio.Features.Home;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioPersistenceRegressionTests : IDisposable {
    private readonly string _root = Path.Combine(Path.GetTempPath(), "officeimo-studio-persistence-" + Guid.NewGuid().ToString("N"));
    public StudioPersistenceRegressionTests() => Directory.CreateDirectory(_root);
    public void Dispose() => Directory.Delete(_root, true);

    [Fact]
    public void A_stale_windows_placement_update_preserves_another_windows_privacy_settings() {
        string path = Path.Combine(_root, "preferences.json");
        var first = new StudioPreferencesService(new JsonStudioPreferencesStore(path));
        var second = new StudioPreferencesService(new JsonStudioPreferencesStore(path));
        first.Update(current => current with { RememberDocumentHistory = false, CreateRecoverySnapshots = false });
        var placement = new StudioWindowPlacement { X = 120, Y = 80, Width = 900, Height = 620 };
        second.UpdatePresentation(current => current with { MainWindow = placement });
        StudioPreferences persisted = new JsonStudioPreferencesStore(path).Load();
        Assert.False(persisted.RememberDocumentHistory);
        Assert.False(persisted.CreateRecoverySnapshots);
        Assert.Equal(placement, persisted.MainWindow);
        Assert.False(second.Current.RememberDocumentHistory);
    }

    [Fact]
    public void Explicit_removal_updates_retained_history_without_enabling_new_recording() {
        bool enabled = true;
        string first = Path.Combine(_root, "first.pdf"), second = Path.Combine(_root, "second.pdf");
        File.WriteAllText(first, "fixture"); File.WriteAllText(second, "fixture");
        var store = new JsonRecentDocumentStore(Path.Combine(_root, "recent.json"), () => enabled);
        store.Save([new(first, DateTimeOffset.UtcNow), new(second, DateTimeOffset.UtcNow)]);
        enabled = false;
        store.Remove(first);
        Assert.Empty(store.Load());
        enabled = true;
        Assert.Equal(second, Assert.Single(store.Load()).Path);
    }
}
