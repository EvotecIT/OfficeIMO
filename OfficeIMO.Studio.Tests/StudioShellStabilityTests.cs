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
        primary.StartListening(paths => received.TrySetResult(paths));

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
        primary.StartListening(paths => received.TrySetResult(paths));

        using StudioSingleInstance later = AcquireFromAnotherLaunch(name);
        Assert.True(later.TryForward([], TimeSpan.FromSeconds(5)));
        Assert.Empty(await received.Task.WaitAsync(TimeSpan.FromSeconds(5)));
    }

    [Fact]
    public async Task Malformed_client_does_not_stop_the_listener() {
        string name = UniqueInstanceName();
        using StudioSingleInstance primary = StudioSingleInstance.Acquire(name);
        var received = new TaskCompletionSource<IReadOnlyList<string>>(TaskCreationOptions.RunContinuationsAsynchronously);
        primary.StartListening(paths => received.TrySetResult(paths));

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
    public void Window_that_cannot_take_files_lets_the_later_launch_open_them() {
        string name = UniqueInstanceName();
        using StudioSingleInstance primary = StudioSingleInstance.Acquire(name);
        primary.StartListening(_ => false);
        using StudioSingleInstance later = AcquireFromAnotherLaunch(name);
        Assert.False(later.TryForward([Path.Combine(Path.GetTempPath(), "rejected.pdf")], TimeSpan.FromSeconds(5)));
    }

    [Fact]
    public async Task Client_that_disconnects_mid_message_does_not_stop_the_listener() {
        string name = UniqueInstanceName();
        using StudioSingleInstance primary = StudioSingleInstance.Acquire(name);
        var received = new TaskCompletionSource<IReadOnlyList<string>>(TaskCreationOptions.RunContinuationsAsynchronously);
        primary.StartListening(paths => received.TrySetResult(paths));
        using (var interrupted = new NamedPipeClientStream(".", name, PipeDirection.InOut, PipeOptions.CurrentUserOnly)) {
            interrupted.Connect(5000);
            interrupted.Write(BitConverter.GetBytes(128));
            interrupted.WriteByte(1);
            interrupted.Flush();
        }
        using StudioSingleInstance later = AcquireFromAnotherLaunch(name);
        string document = Path.Combine(Path.GetTempPath(), "after-interrupted.pdf");
        Assert.True(later.TryForward([document], TimeSpan.FromSeconds(5)));
        Assert.Equal(new[] { document }, await received.Task.WaitAsync(TimeSpan.FromSeconds(5)));
    }

    [Fact]
    public void Listener_takes_the_next_launch_while_an_earlier_batch_is_still_opening() {
        string name = UniqueInstanceName();
        using StudioSingleInstance primary = StudioSingleInstance.Acquire(name);
        var batches = new System.Collections.Concurrent.ConcurrentQueue<IReadOnlyList<string>>();
        // Accepting only queues the batch; the slow open happens elsewhere and must not hold the pipe.
        primary.StartListening(paths => { batches.Enqueue(paths); return true; });
        using StudioSingleInstance second = AcquireFromAnotherLaunch(name);
        using StudioSingleInstance third = AcquireFromAnotherLaunch(name);
        Assert.True(second.TryForward([Path.Combine(Path.GetTempPath(), "slow.pdf")], TimeSpan.FromSeconds(5)));
        Assert.True(third.TryForward([Path.Combine(Path.GetTempPath(), "next.pdf")], TimeSpan.FromSeconds(5)));
        Assert.Equal(2, batches.Count);
    }

    [Fact]
    public async Task Listener_keeps_a_launch_queued_before_the_previous_reply() {
        string name = UniqueInstanceName();
        using StudioSingleInstance primary = StudioSingleInstance.Acquire(name);
        var firstReceived = new TaskCompletionSource(TaskCreationOptions.RunContinuationsAsynchronously);
        using var releaseReply = new ManualResetEventSlim();
        var batches = new System.Collections.Concurrent.ConcurrentQueue<IReadOnlyList<string>>();
        string firstPath = Path.Combine(Path.GetTempPath(), "first.pdf");
        string nextPath = Path.Combine(Path.GetTempPath(), "queued.pdf");
        primary.StartListening(paths => {
            batches.Enqueue(paths);
            if (paths.Contains(firstPath)) {
                firstReceived.TrySetResult();
                Assert.True(releaseReply.Wait(TimeSpan.FromSeconds(5)));
            }
            return true;
        });
        using StudioSingleInstance second = AcquireFromAnotherLaunch(name);
        Task<bool> first = Task.Run(() => second.TryForward([firstPath], TimeSpan.FromSeconds(5)));
        using var deadline = new CancellationTokenSource(TimeSpan.FromSeconds(5));
        await using var queued = new NamedPipeClientStream(".", name, PipeDirection.InOut,
            PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
        try {
            await firstReceived.Task.WaitAsync(deadline.Token);
            Task connecting = queued.ConnectAsync(deadline.Token);
            // Unix can connect to the socket backlog while the current request is active.
            // Windows completes its connection after the current instance disconnects.
            if (!OperatingSystem.IsWindows()) await connecting;
            releaseReply.Set();
            Assert.True(await first);
            await connecting;
            byte[] payload = System.Text.Json.JsonSerializer.SerializeToUtf8Bytes(new { Version = 1, Paths = new[] { nextPath } });
            byte[] length = new byte[4];
            System.Buffers.Binary.BinaryPrimitives.WriteInt32LittleEndian(length, payload.Length);
            await queued.WriteAsync(length, deadline.Token);
            await queued.WriteAsync(payload, deadline.Token);
            await queued.FlushAsync(deadline.Token);
            byte[] reply = new byte[1];
            Assert.Equal(1, await queued.ReadAsync(reply, deadline.Token));
            Assert.Equal(1, reply[0]);
            Assert.Equal(new[] { firstPath, nextPath }, batches.SelectMany(paths => paths));
        } finally {
            releaseReply.Set();
        }
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
    public async Task Instance_that_stops_answering_after_connecting_lets_the_later_launch_run_independently() {
        string name = UniqueInstanceName();
        using StudioSingleInstance primary = StudioSingleInstance.Acquire(name);
        // A stalled owner: it takes the connection and the message but never replies.
        await using var stalled = new NamedPipeServerStream(name, PipeDirection.InOut, 1, PipeTransmissionMode.Byte,
            PipeOptions.Asynchronous | PipeOptions.CurrentUserOnly);
        Task connected = stalled.WaitForConnectionAsync();
        using StudioSingleInstance later = AcquireFromAnotherLaunch(name);
        Assert.False(await Task.Run(() => later.TryForward([Path.Combine(Path.GetTempPath(), "stalled.pdf")], TimeSpan.FromMilliseconds(500)))
            .WaitAsync(TimeSpan.FromSeconds(5)));
        await connected.WaitAsync(TimeSpan.FromSeconds(5));
    }

    [Fact]
    public void Linux_picker_globs_match_any_letter_case() {
        string[] globs = StudioPickerPatterns.CaseInsensitive(["*.pdf", "*.tar.gz"], platformIgnoresCase: false);
        Assert.Equal(new[] { "*.[pP][dD][fF]", "*.[tT][aA][rR].[gG][zZ]" }, globs);
        Assert.All(new[] { "report.pdf", "SCAN.PDF", "report.Pdf", "a.pDf" },
            name => Assert.Matches("^" + System.Text.RegularExpressions.Regex.Escape(globs[0]).Replace(@"\*", ".*").Replace(@"\[", "[") + "$", name));
        Assert.Equal(new[] { "*.pdf" }, StudioPickerPatterns.CaseInsensitive(["*.pdf"], platformIgnoresCase: true));
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
        await session.Dispatch(async () => {
            var services = TestAppBuilder.CreateTestServices();
            using var tab = CreateTab(services);
            Assert.True(tab.ShowRecentPreviews);
            Directory.CreateDirectory(services.Paths.Root);
            string source = Path.Combine(services.Paths.Root, "private.pdf");
            PdfDocument.Create(document => document.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(200), global::OfficeIMO.OfficeLength.Points(300)).Content(content =>
                content.Item(item => item.Paragraph(text => text.Text("Private")))))).Save(source);
            var recent = new RecentDocumentViewModel(source, DateTimeOffset.UtcNow);
            tab.RecentDocuments.Add(recent);
            recent.EnsureThumbnail();
            for (int attempt = 0; attempt < 200 && !recent.HasThumbnail; attempt++) {
                Avalonia.Threading.Dispatcher.UIThread.RunJobs();
                await Task.Delay(20);
            }
            Assert.True(recent.HasThumbnail);

            tab.Settings.ShowRecentPreviews = false;

            // The rendered page disappears immediately, not when Home next refreshes its previews.
            Assert.False(recent.HasThumbnail);
            Assert.False(services.Preferences.Current.ShowRecentPreviews);
            Assert.False(tab.ShowRecentPreviews);
            Assert.False(new JsonStudioPreferencesStore(services.Paths.PreferencesPath).Load().ShowRecentPreviews);
            return true;
        }, CancellationToken.None);
    }

    private static MainWindowViewModel CreateTab(StudioApplicationServices services) => new(
        _ => Task.FromResult<string?>(null),
        recentDocumentStore: services.DocumentHistory.RecentDocuments,
        services: services);

    private static string CreateDocument(string path) {
        PdfDocument.Create(compose => compose.Page(page => page.Size(global::OfficeIMO.OfficeLength.Points(420D), global::OfficeIMO.OfficeLength.Points(620D)))).Save(path);
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
