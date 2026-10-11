using OfficeIMO.Studio.Features.Shell;
using OfficeIMO.Studio.Features.Workspace;
using OfficeIMO.Studio.Infrastructure;
using OfficeIMO.Studio.Infrastructure.Preferences;

namespace OfficeIMO.Studio.Tests;

public sealed class StudioCoreSessionTests {
    [Fact]
    public void OlderSessionRecordsKeepViewAndTimestampDefaults() {
        string root = Path.Combine(Path.GetTempPath(), "studio-core-legacy-session-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string source = Path.Combine(root, "first.pdf");
            string path = Path.Combine(root, "session.json");
            // Earlier session records omit LastUsedAt; partial view records use model defaults.
            File.WriteAllText(path, System.Text.Json.JsonSerializer.Serialize(new {
                SchemaVersion = 1, UpdatedAt = DateTimeOffset.UtcNow, ActivePath = source,
                Documents = new[] { new { Path = source, Fingerprint = new string('A', 64), View = new { PageNumber = 2 } } }
            }));
            var document = Assert.Single(new StudioSessionStore(path).Load().Documents);
            Assert.Equal(2, document.View.PageNumber);
            Assert.Equal(1, document.View.Zoom);
            Assert.Equal(238, document.View.NavigationWidth);
            Assert.Equal(300, document.View.InspectorWidth);
            Assert.Equal(OfficeIMO.Studio.Features.Reader.ViewerZoomMode.FitWidth, document.View.ZoomMode);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task DebouncedSaveKeepsLatestViewStateAndShutdownCancelsPendingSave() {
        string root = Path.Combine(Path.GetTempPath(), "studio-core-session-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            using var tabs = CreateTabs();
            var environment = new TestEnvironment(root);
            var scheduler = new ManualScheduler();
            using var session = new StudioDocumentSessionController<TestDocument, TestTab>(
                tabs, environment, _ => Task.FromResult<string?>(null), scheduler);
            await tabs.OpenDocumentAsync(Path.Combine(root, "first.pdf"));
            await tabs.OpenDocumentAsync(Path.Combine(root, "second.pdf"));
            tabs.Tabs[0].Document.SetView(new() { PageNumber = 3, Zoom = 1.5 });
            tabs.ActiveDocument.SetView(new() { PageNumber = 7, Zoom = 2 });
            Assert.Single(scheduler.Active);
            scheduler.RunPending();
            var saved = environment.SessionStore.Load();
            Assert.Equal(2, saved.Documents.Count);
            Assert.Equal(3, saved.Documents[0].View.PageNumber);
            Assert.Equal(7, saved.Documents[1].View.PageNumber);
            Assert.Equal(tabs.ActiveDocument.DocumentPath, saved.ActivePath);

            tabs.ActiveDocument.SetView(new() { PageNumber = 8 });
            session.CaptureForShutdown();
            Assert.Empty(scheduler.Active);
            Assert.Equal(8, environment.SessionStore.Load().Documents[1].View.PageNumber);
            tabs.ActiveDocument.SetView(new() { PageNumber = 9 });
            scheduler.RunPending();
            Assert.Equal(8, environment.SessionStore.Load().Documents[1].View.PageNumber);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task DisablingSessionPersistenceClearsSavedAndPendingState() {
        string root = Path.Combine(Path.GetTempPath(), "studio-core-privacy-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            using var tabs = CreateTabs();
            var environment = new TestEnvironment(root);
            var scheduler = new ManualScheduler();
            using var session = new StudioDocumentSessionController<TestDocument, TestTab>(
                tabs, environment, _ => Task.FromResult<string?>(null), scheduler);
            await tabs.OpenDocumentAsync(Path.Combine(root, "first.pdf"));
            scheduler.RunPending();
            Assert.Single(environment.SessionStore.Load().Documents);
            environment.DisablePersistence();
            Assert.Empty(environment.SessionStore.Load().Documents);
            session.Dispose();
            Assert.Empty(scheduler.Active);
        } finally { Directory.Delete(root, recursive: true); }
    }

    [Fact]
    public async Task VisibleDocumentsKeepPresentationActiveWhileOnlyTheSelectedTabOwnsCommands() {
        using var tabs = CreateTabs();
        await tabs.OpenDocumentAsync(Path.Combine(Path.GetTempPath(), "pane-first.pdf"));
        var first = tabs.SelectedTab!;
        await tabs.OpenDocumentAsync(Path.Combine(Path.GetTempPath(), "pane-second.pdf"));
        var second = tabs.SelectedTab!;
        tabs.SetPresentedDocuments([first.Document, second.Document]);
        tabs.SelectedTab = first;
        Assert.True(first.Document.PresentationActive);
        Assert.True(second.Document.PresentationActive);
        Assert.Same(first.Document, tabs.ActiveDocument);
        tabs.SelectedTab = second;
        Assert.True(first.Document.PresentationActive);
        Assert.True(second.Document.PresentationActive);
        Assert.Same(second.Document, tabs.ActiveDocument);
        tabs.SetPresentedDocuments([]);
        Assert.False(first.Document.PresentationActive);
        Assert.True(second.Document.PresentationActive);
    }

    [Fact]
    public async Task ApplicationDocumentTakesPartInWindowLifecycleWithoutBecomingATab() {
        var application = new TestDocument();
        var tabs = new StudioDocumentTabs<TestDocument, TestTab>(_ => new(), _ => { }, (document, _) => new(document), applicationDocument: application);
        await tabs.OpenDocumentAsync(Path.Combine(Path.GetTempPath(), "window-owned-first.pdf"));
        Assert.DoesNotContain(tabs.Tabs, tab => ReferenceEquals(tab.Document, application));
        Assert.Contains(application, tabs.OperationDocuments);

        Assert.False(tabs.HasBusyDocuments);
        application.CanCancelOperation = true;
        Assert.True(tabs.HasBusyDocuments);
        application.HasUnsavedAuxiliaryWork = true;
        Assert.True(tabs.HasDirtyDocuments);

        // Output reserved by window-level work cannot be claimed by a document save.
        string reserved = Path.Combine(Path.GetTempPath(), "window-owned-book.epub");
        application.OwnedPath = reserved;
        Assert.False(tabs.CanPublishPath(reserved));

        tabs.CancelAllOperations();
        Assert.Equal(1, application.Cancellations);

        application.AllowClose = false;
        Assert.False(await tabs.RequestCloseAllAsync());
        Assert.Single(tabs.Tabs);
        application.AllowClose = true;
        Assert.True(await tabs.RequestCloseAllAsync());
        Assert.Equal(2, application.CloseRequests);

        tabs.Dispose();
        Assert.True(application.Disposed);
    }

    private static StudioDocumentTabs<TestDocument, TestTab> CreateTabs() =>
        new(_ => new(), _ => { }, (document, _) => new(document));

    private sealed class TestTab(TestDocument document) : IStudioDocumentTab<TestDocument> {
        public TestDocument Document => document;
        public string Title { get; set; } = string.Empty;
        public void Dispose() => document.Dispose();
    }

    private sealed class TestDocument : IStudioDocument {
        private StudioDocumentViewState _view = new();
        public event EventHandler<StudioDocumentChangedEventArgs>? Changed;
        public string DocumentName => Path.GetFileName(DocumentPath) ?? string.Empty;
        public string? DocumentPath { get; private set; }
        public string? LastClosedDocumentPath => null;
        public bool HasDocument => DocumentPath is not null;
        public bool IsDirty => false;
        public bool HasUnsavedAuxiliaryWork { get; set; }
        public bool CanCancelOperation { get; set; }
        internal int Cancellations { get; private set; }
        internal int CloseRequests { get; private set; }
        internal bool AllowClose { get; set; } = true;
        internal bool Disposed { get; private set; }
        internal string? OwnedPath { get; set; }
        public string? ErrorMessage { get; set; }
        public IEnumerable<string> OwnedOutputLocations => OwnedPath is null ? [] : [OwnedPath];
        public bool OwnsOutputPath(string path) => OwnedPath is not null && string.Equals(Path.GetFullPath(OwnedPath), Path.GetFullPath(path), StringComparison.OrdinalIgnoreCase);
        public void Deactivate() { }
        internal bool PresentationActive { get; private set; }
        public void SetPresentationActive(bool active) => PresentationActive = active;
        public Task OpenDocumentAsync(string path, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            DocumentPath = path;
            Changed?.Invoke(this, new(StudioDocumentChangeKind.Document));
            return Task.CompletedTask;
        }
        internal void SetView(StudioDocumentViewState state) {
            _view = state;
            Changed?.Invoke(this, new(StudioDocumentChangeKind.View));
        }
        public StudioSessionDocument CaptureSessionDocument() => new(DocumentPath!, new string('A', 64), _view);
        public void RestoreSessionViewState(StudioDocumentViewState state) => SetView(state);
        public Task<bool> RequestCloseDocumentAsync() => Task.FromResult(true);
        public Task<bool> PrepareCloseDocumentAsync() { CloseRequests++; return Task.FromResult(AllowClose); }
        public Task<bool> CommitPreparedDiscardAsync() => Task.FromResult(true);
        public void CancelPreparedClose() { }
        public void CompletePreparedClose() { }
        public void CancelCurrentOperation() => Cancellations++;
        public void Dispose() => Disposed = true;
    }

    private sealed class TestEnvironment(string root) : IStudioSessionEnvironment {
        public StudioSessionStore SessionStore { get; } = new(Path.Combine(root, "session.json"));
        public PdfWorkspaceRecoveryStore Recovery { get; } = new(Path.Combine(root, "recovery"));
        public StudioDocumentStorage Storage { get; } = new();
        public bool RememberSession { get; private set; } = true;
        public event EventHandler? PreferencesChanged;
        public event EventHandler? SessionCleared { add { } remove { } }
        public string Text(string key) => key;
        public string Format(string key, int count) => key;
        public void ReportFailure(string code, Exception error) => Assert.Fail(code + ": " + error);
        internal void DisablePersistence() { RememberSession = false; PreferencesChanged?.Invoke(this, EventArgs.Empty); }
        public void Dispose() => Storage.Dispose();
    }

    private sealed class ManualScheduler : IStudioScheduler {
        private readonly List<Callback> _callbacks = [];
        internal IEnumerable<Callback> Active => _callbacks.Where(callback => !callback.Canceled);
        public void Post(Action callback) => callback();
        public IDisposable Schedule(TimeSpan delay, Action callback) {
            var scheduled = new Callback(callback);
            _callbacks.Add(scheduled);
            return scheduled;
        }
        internal void RunPending() {
            foreach (var callback in Active.ToArray()) { callback.Dispose(); callback.Action(); }
        }
        internal sealed class Callback(Action action) : IDisposable {
            internal Action Action => action;
            internal bool Canceled { get; private set; }
            public void Dispose() => Canceled = true;
        }
    }
}
