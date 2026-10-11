using System.IO.Pipes;
using OfficeIMO.Pdf;
using OfficeIMO.Studio.Features.Home;
using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioShellStabilityTests {
    [Fact]
    public async Task Later_launch_forwards_rooted_paths_to_the_running_instance() {
        string name = UniqueInstanceName();
        using StudioSingleInstance primary = StudioSingleInstance.Acquire(name);
        Assert.True(primary.IsPrimary);
        var received = new TaskCompletionSource<IReadOnlyList<string>>(TaskCreationOptions.RunContinuationsAsynchronously);
        primary.StartListening(paths => { received.TrySetResult(paths); return Task.CompletedTask; });

        using StudioSingleInstance later = AcquireFromAnotherLaunch(name);
        Assert.False(later.IsPrimary);
        string document = Path.Combine(Path.GetTempPath(), "forwarded.pdf");
        Assert.True(later.TryForward([document, "relative.pdf"], TimeSpan.FromSeconds(5)));

        IReadOnlyList<string> paths = await received.Task.WaitAsync(TimeSpan.FromSeconds(5));
        Assert.Equal(new[] { document }, paths);
    }

    [Fact]
    public async Task Launch_without_files_is_forwarded_so_the_running_window_comes_forward() {
        string name = UniqueInstanceName();
        using StudioSingleInstance primary = StudioSingleInstance.Acquire(name);
        var received = new TaskCompletionSource<IReadOnlyList<string>>(TaskCreationOptions.RunContinuationsAsynchronously);
        primary.StartListening(paths => { received.TrySetResult(paths); return Task.CompletedTask; });

        using StudioSingleInstance later = AcquireFromAnotherLaunch(name);
        Assert.True(later.TryForward([], TimeSpan.FromSeconds(5)));
        Assert.Empty(await received.Task.WaitAsync(TimeSpan.FromSeconds(5)));
    }

    [Fact]
    public async Task Malformed_client_does_not_stop_the_listener() {
        string name = UniqueInstanceName();
        using StudioSingleInstance primary = StudioSingleInstance.Acquire(name);
        var received = new TaskCompletionSource<IReadOnlyList<string>>(TaskCreationOptions.RunContinuationsAsynchronously);
        primary.StartListening(paths => { received.TrySetResult(paths); return Task.CompletedTask; });

        using (var hostile = new NamedPipeClientStream(".", name, PipeDirection.InOut, PipeOptions.CurrentUserOnly)) {
            hostile.Connect(5000);
            hostile.Write(BitConverter.GetBytes(int.MaxValue));
            hostile.Flush();
            try { Assert.Equal(-1, hostile.ReadByte()); }
            catch (IOException) { }
        }

        using StudioSingleInstance later = AcquireFromAnotherLaunch(name);
        string document = Path.Combine(Path.GetTempPath(), "after-malformed.pdf");
        Assert.True(later.TryForward([document], TimeSpan.FromSeconds(5)));
        Assert.Equal(new[] { document }, await received.Task.WaitAsync(TimeSpan.FromSeconds(5)));
    }

    [Fact]
    public void Unreachable_instance_lets_the_later_launch_run_independently() {
        string name = UniqueInstanceName();
        using StudioSingleInstance primary = StudioSingleInstance.Acquire(name);
        // The owner holds the lock but has not started listening yet.
        using StudioSingleInstance later = AcquireFromAnotherLaunch(name);
        Assert.False(later.TryForward(["/tmp/a.pdf"], TimeSpan.FromMilliseconds(200)));
    }

    [Fact]
    public void Instance_name_depends_on_the_data_root() {
        string first = StudioSingleInstance.CreateName(Path.Combine(Path.GetTempPath(), "profile-a"));
        Assert.Equal(first, StudioSingleInstance.CreateName(Path.Combine(Path.GetTempPath(), "profile-a")));
        Assert.NotEqual(first, StudioSingleInstance.CreateName(Path.Combine(Path.GetTempPath(), "profile-b")));
        Assert.True(first.Length < 40);
    }

    [Fact]
    public void Presentation_change_survives_an_unwritable_store_and_reports_the_failure() {
        var store = new FailingPreferencesStore();
        var preferences = new StudioPreferencesService(store);
        Exception? reported = null;
        int changes = 0;
        preferences.Changed += (_, _) => changes++;
        preferences.PersistenceFailed += (_, error) => reported = error;

        preferences.UpdatePresentation(current => current with { Theme = StudioThemePreference.Dark });

        Assert.Equal(StudioThemePreference.Dark, preferences.Current.Theme);
        Assert.Equal(1, changes);
        Assert.IsType<IOException>(reported);

        // Privacy and retention choices stay transactional: a failed write is not applied.
        Assert.Throws<IOException>(() => preferences.Update(current => current with { RememberDocumentHistory = false }));
        Assert.True(preferences.Current.RememberDocumentHistory);
    }

    [Fact]
    public void Preference_saves_leave_no_staging_files() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-preferences-staging-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            var first = new JsonStudioPreferencesStore(Path.Combine(root, "preferences.json"));
            var second = new JsonStudioPreferencesStore(Path.Combine(root, "preferences.json"));
            first.Save(new StudioPreferences { Theme = StudioThemePreference.Light });
            second.Save(new StudioPreferences { Theme = StudioThemePreference.Dark });

            Assert.Equal(StudioThemePreference.Dark, first.Load().Theme);
            Assert.Equal(new[] { "preferences.json" }, Directory.GetFiles(root).Select(Path.GetFileName));
            if (!OperatingSystem.IsWindows())
                Assert.Equal(UnixFileMode.UserRead | UnixFileMode.UserWrite, File.GetUnixFileMode(Path.Combine(root, "preferences.json")));
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public void Repeated_interface_failures_stop_being_recovered() {
        var app = new App();
        DateTime now = new(2026, 10, 11, 12, 0, 0, DateTimeKind.Utc);
        for (int index = 0; index < 5; index++)
            Assert.True(app.CanRecoverFromUserInterfaceFailure(new InvalidOperationException(), now.AddSeconds(index)));
        Assert.False(app.CanRecoverFromUserInterfaceFailure(new InvalidOperationException(), now.AddSeconds(5)));
        Assert.True(app.CanRecoverFromUserInterfaceFailure(new InvalidOperationException(), now.AddSeconds(30)));
        Assert.False(app.CanRecoverFromUserInterfaceFailure(new OutOfMemoryException(), now.AddSeconds(60)));
    }

    [Fact]
    public async Task Recent_documents_stay_consistent_across_tabs() {
        string root = Path.Combine(Path.GetTempPath(), "officeimo-studio-recent-tabs-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            var services = StudioApplicationServices.Create(new StudioDataPaths(Path.Combine(root, "profile")));
            string first = CreateDocument(Path.Combine(root, "first.pdf"));
            string second = CreateDocument(Path.Combine(root, "second.pdf"));
            using var firstTab = CreateTab(services);
            using var secondTab = CreateTab(services);

            await firstTab.OpenDocumentAsync(first);
            await secondTab.OpenDocumentAsync(second);
            Assert.Equal(new[] { second, first }, firstTab.RecentDocuments.Select(document => document.Path));
            Assert.Equal(new[] { second, first }, services.DocumentHistory.RecentDocuments.Load().Select(document => document.Path));

            secondTab.RemoveRecentDocumentCommand.Execute(secondTab.RecentDocuments.Single(document => document.Path == second));
            Assert.Equal(new[] { first }, firstTab.RecentDocuments.Select(document => document.Path));

            firstTab.ClearRecentDocumentsCommand.Execute(null);
            Assert.Empty(secondTab.RecentDocuments);
            Assert.False(secondTab.HasRecentDocuments);

            // A tab that had the old list must not write cleared entries back.
            await secondTab.OpenDocumentAsync(second);
            Assert.Equal(new[] { second }, services.DocumentHistory.RecentDocuments.Load().Select(document => document.Path));
            Assert.Equal(new[] { second }, firstTab.RecentDocuments.Select(document => document.Path));
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public void Runtime_file_errors_get_plain_wording_while_specific_messages_are_kept() {
        var text = OfficeIMO.Studio.Infrastructure.Localization.StudioLocalization.Current;
        Assert.Equal(text.Get("Errors.AccessDenied"), StudioMessages.Describe(new UnauthorizedAccessException(@"Access to the path 'C:\Users\a\secret.pdf' is denied.")));
        Assert.Equal(text.Get("Errors.NotFound"), StudioMessages.Describe(new FileNotFoundException("Could not find file '/home/a/report.pdf'.")));
        Assert.Equal("Provider permission expired.", StudioMessages.Describe(new UnauthorizedAccessException("Provider permission expired.")));
        Assert.Equal("Page range is invalid", StudioMessages.Describe(new ArgumentException("Page range is invalid", "range")));
        var diskFull = new IOException("disk full", OperatingSystem.IsWindows() ? unchecked((int)0x80070070) : 28);
        Assert.Equal(text.Get("Errors.DiskFull"), StudioMessages.Describe(diskFull));
        if (OperatingSystem.IsWindows())
            Assert.Equal(text.Get("Errors.FileInUse"), StudioMessages.Describe(new IOException("in use", unchecked((int)0x80070020))));
    }

    [Fact]
    public async Task Recent_previews_can_be_turned_off_for_private_work() {
        using var session = TestAppBuilder.StartSession();
        await session.Dispatch(() => {
            var services = TestAppBuilder.CreateTestServices();
            using var tab = CreateTab(services);
            Assert.True(tab.ShowRecentPreviews);

            tab.Settings.ShowRecentPreviews = false;

            Assert.False(services.Preferences.Current.ShowRecentPreviews);
            Assert.False(tab.ShowRecentPreviews);
            Assert.False(new JsonStudioPreferencesStore(services.Paths.PreferencesPath).Load().ShowRecentPreviews);
            var recent = new RecentDocumentViewModel(Path.Combine(services.Paths.Root, "private.pdf"), DateTimeOffset.UtcNow);
            recent.ClearThumbnail();
            Assert.False(recent.HasThumbnail);
            return true;
        }, CancellationToken.None);
    }

    private static MainWindowViewModel CreateTab(StudioApplicationServices services) => new(
        _ => Task.FromResult<string?>(null),
        recentDocumentStore: services.DocumentHistory.RecentDocuments,
        services: services);

    private static string CreateDocument(string path) {
        PdfDocument.Create(compose => compose.Page(page => page.Size(420D, 620D))).Save(path);
        return Path.GetFullPath(path);
    }

    // Named mutexes are re-entrant on one thread; a separate thread stands in for the later process.
    private static StudioSingleInstance AcquireFromAnotherLaunch(string name) =>
        Task.Factory.StartNew(() => StudioSingleInstance.Acquire(name), TaskCreationOptions.LongRunning).GetAwaiter().GetResult();

    private static string UniqueInstanceName() =>
        StudioSingleInstance.CreateName(Path.Combine(Path.GetTempPath(), "officeimo-studio-instance-" + Guid.NewGuid().ToString("N")));

    private sealed class FailingPreferencesStore : IStudioPreferencesStore {
        public StudioPreferences Load() => new();

        public void Save(StudioPreferences preferences) => throw new IOException("Profile is read-only.");
    }
}
