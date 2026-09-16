using OfficeIMO.GoogleWorkspace.Drive;
using System.IO;

namespace OfficeIMO.GoogleWorkspace.Sync {
    /// <summary>Identifies the Drive change feed to which a change or outcome belongs.</summary>
    public enum GoogleWorkspaceChangeSourceKind {
        /// <summary>The account's user change feed.</summary>
        User = 0,
        /// <summary>A single shared-drive change feed.</summary>
        SharedDrive = 1,
    }
    /// <summary>Outcome of reading one Drive change feed.</summary>
    public enum GoogleWorkspaceChangeReadStatus {
        /// <summary>All pages were read and a new token was obtained.</summary>
        Completed = 0,
        /// <summary>No previous token existed; tracking began at the current cursor without historical changes.</summary>
        Initialized = 1,
        /// <summary>The feed failed and its previous token was retained.</summary>
        Failed = 2,
    }

    /// <summary>The user feed or one shared-drive feed tracked by a read call.</summary>
    public sealed class GoogleWorkspaceChangeSource {
        internal GoogleWorkspaceChangeSource(GoogleWorkspaceChangeSourceKind kind, string? driveId) { Kind = kind; DriveId = driveId; }
        /// <summary>Gets which kind of Drive change feed this source represents.</summary>
        public GoogleWorkspaceChangeSourceKind Kind { get; }
        /// <summary>Gets the shared-drive identifier, or null for the user feed.</summary>
        public string? DriveId { get; }
        /// <summary>Gets <c>user</c> or a <c>drive:</c>-prefixed shared-drive identifier.</summary>
        public string Key => Kind == GoogleWorkspaceChangeSourceKind.User ? "user" : "drive:" + DriveId;
    }

    /// <summary>A Drive change paired with the feed from which it was read.</summary>
    public sealed class GoogleWorkspaceTrackedChange {
        internal GoogleWorkspaceTrackedChange(GoogleWorkspaceChangeSource source, GoogleDriveChange change) { Source = source; Change = change; }
        /// <summary>Gets the feed that supplied this change.</summary>
        public GoogleWorkspaceChangeSource Source { get; }
        /// <summary>Gets the Drive change returned by the provider.</summary>
        public GoogleDriveChange Change { get; }
    }

    /// <summary>Limits and filters used while reading user and shared-drive change feeds.</summary>
    public sealed class GoogleWorkspaceChangeReadOptions {
        /// <summary>Gets the mutable list of additional shared drives to track, alongside those already in the checkpoint.</summary>
        public IList<string> SharedDriveIds { get; } = new List<string>();
        /// <summary>Gets or sets the requested Drive page size; defaults to 100.</summary>
        public int PageSize { get; set; } = 100;
        /// <summary>Gets or sets the maximum number of pages accepted from each source; defaults to 10,000.</summary>
        public int MaxPagesPerSource { get; set; } = 10000;
        /// <summary>Gets or sets the maximum number of returned changes accepted from each source; defaults to 10,000.</summary>
        public int MaxChangesPerSource { get; set; } = 10_000;
        /// <summary>Gets or sets the maximum number of returned changes accepted across all sources; defaults to 50,000.</summary>
        public int MaxTotalChanges { get; set; } = 50_000;
        /// <summary>Gets or sets whether Drive includes removed entries; defaults to true.</summary>
        public bool IncludeRemoved { get; set; } = true;
        /// <summary>Gets or sets whether Drive includes corpus-removal changes; defaults to true.</summary>
        public bool IncludeCorpusRemovals { get; set; } = true;
        /// <summary>Gets or sets whether a failed, already-initialized source allows later sources to continue; defaults to true.</summary>
        /// <remarks>Cancellation and failure to initialize a missing token still throw.</remarks>
        public bool ContinueOnSourceFailure { get; set; } = true;
    }

    /// <summary>Cursor and status after reading or initializing one change feed.</summary>
    public sealed class GoogleWorkspaceChangeSourceResult {
        internal GoogleWorkspaceChangeSourceResult(GoogleWorkspaceChangeSource source, GoogleWorkspaceChangeReadStatus status, int changeCount, string? nextToken, Exception? exception) {
            Source = source; Status = status; ChangeCount = changeCount; NextToken = nextToken; Exception = exception;
        }
        /// <summary>Gets the user or shared-drive feed represented by this result.</summary>
        public GoogleWorkspaceChangeSource Source { get; }
        /// <summary>Gets whether this source completed, was initialized, or failed.</summary>
        public GoogleWorkspaceChangeReadStatus Status { get; }
        /// <summary>Gets the number of changes included from this source; failed and initialized sources report zero.</summary>
        public int ChangeCount { get; }
        /// <summary>Gets the new token after completion or initialization, or the retained token after failure.</summary>
        public string? NextToken { get; }
        /// <summary>Gets the exception captured for a failed source, or null otherwise.</summary>
        public Exception? Exception { get; }
    }

    /// <summary>Changes and source outcomes returned by a completed tracking call.</summary>
    public sealed class GoogleWorkspaceChangeReadResult {
        internal GoogleWorkspaceChangeReadResult(IReadOnlyList<GoogleWorkspaceTrackedChange> changes, GoogleWorkspaceSyncCheckpoint checkpoint, IReadOnlyList<GoogleWorkspaceChangeSourceResult> sources, TranslationReport report) {
            Changes = Array.AsReadOnly(changes.ToArray());
            NextCheckpoint = checkpoint;
            Sources = Array.AsReadOnly(sources.ToArray());
            Report = report;
        }
        /// <summary>Gets changes from sources that reached the end of their feed.</summary>
        public IReadOnlyList<GoogleWorkspaceTrackedChange> Changes { get; }
        /// <summary>Gets a separate checkpoint with tokens advanced only for completed or newly initialized sources.</summary>
        /// <remarks>The caller is responsible for persisting it after successfully processing the returned changes.</remarks>
        public GoogleWorkspaceSyncCheckpoint NextCheckpoint { get; }
        /// <summary>Gets one outcome for each source visited, in user-then-shared-drive order.</summary>
        public IReadOnlyList<GoogleWorkspaceChangeSourceResult> Sources { get; }
        /// <summary>Gets diagnostic notices recorded during the read.</summary>
        public TranslationReport Report { get; }
        /// <summary>Gets whether any source produced a captured failure.</summary>
        public bool HasFailures => Sources.Any(source => source.Status == GoogleWorkspaceChangeReadStatus.Failed);
    }

    /// <summary>Consumes complete Drive change pages and advances each source token only after that source succeeds.</summary>
    public sealed class GoogleWorkspaceChangeTracker : IDisposable {
        private readonly GoogleDriveClient _drive;

        /// <summary>Creates a tracker using the supplied session and optional Drive client settings.</summary>
        public GoogleWorkspaceChangeTracker(GoogleWorkspaceSession session, GoogleDriveClientOptions? options = null) {
            _drive = new GoogleDriveClient(session ?? throw new ArgumentNullException(nameof(session)), options);
        }

        /// <summary>Obtains current start tokens for the user feed and distinct supplied shared drives.</summary>
        /// <remarks>Initialization begins at the current cursors; it does not read historical changes.</remarks>
        public async Task<GoogleWorkspaceSyncCheckpoint> InitializeAsync(IEnumerable<string>? sharedDriveIds = null, CancellationToken cancellationToken = default) {
            var checkpoint = new GoogleWorkspaceSyncCheckpoint {
                UserChangeToken = await _drive.GetStartPageTokenAsync(cancellationToken: cancellationToken).ConfigureAwait(false),
            };
            foreach (string driveId in NormalizeDriveIds(sharedDriveIds)) {
                checkpoint.SharedDriveChangeTokens[driveId] = await _drive.GetStartPageTokenAsync(driveId, cancellationToken: cancellationToken).ConfigureAwait(false);
            }
            return checkpoint;
        }

        /// <summary>Reads each feed to its new start token without advancing a failed source in the returned checkpoint.</summary>
        /// <remarks>The input checkpoint is cloned. Missing tokens are initialized at the current cursor. A source's changes are included only when all of its pages succeed; source-initialization failures and cancellation propagate.</remarks>
        public async Task<GoogleWorkspaceChangeReadResult> ReadAsync(GoogleWorkspaceSyncCheckpoint checkpoint, GoogleWorkspaceChangeReadOptions? options = null, CancellationToken cancellationToken = default) {
            if (checkpoint == null) throw new ArgumentNullException(nameof(checkpoint));
            options ??= new GoogleWorkspaceChangeReadOptions();
            if (options.MaxPagesPerSource < 1) throw new ArgumentOutOfRangeException(nameof(options.MaxPagesPerSource));
            if (options.MaxChangesPerSource < 1) throw new ArgumentOutOfRangeException(nameof(options.MaxChangesPerSource));
            if (options.MaxTotalChanges < 1) throw new ArgumentOutOfRangeException(nameof(options.MaxTotalChanges));
            var report = new TranslationReport();
            GoogleWorkspaceSyncCheckpoint next = checkpoint.Clone();
            var changes = new List<GoogleWorkspaceTrackedChange>();
            var outcomes = new List<GoogleWorkspaceChangeSourceResult>();
            var sources = new List<GoogleWorkspaceChangeSource> { new GoogleWorkspaceChangeSource(GoogleWorkspaceChangeSourceKind.User, null) };
            sources.AddRange(NormalizeDriveIds(checkpoint.SharedDriveChangeTokens.Keys.Concat(options.SharedDriveIds))
                .Select(id => new GoogleWorkspaceChangeSource(GoogleWorkspaceChangeSourceKind.SharedDrive, id)));
            var partitionedDriveIds = new HashSet<string>(
                checkpoint.SharedDriveChangeTokens
                    .Where(pair => !string.IsNullOrWhiteSpace(pair.Value))
                    .Select(pair => pair.Key),
                StringComparer.Ordinal);
            bool hasUnpartitionedSharedDrives = sources.Any(source =>
                source.Kind == GoogleWorkspaceChangeSourceKind.SharedDrive
                && !partitionedDriveIds.Contains(source.DriveId!));

            foreach (GoogleWorkspaceChangeSource source in sources) {
                cancellationToken.ThrowIfCancellationRequested();
                string? startToken = TokenFor(next, source);
                if (string.IsNullOrWhiteSpace(startToken)) {
                    string initialized = await _drive.GetStartPageTokenAsync(source.DriveId, report, cancellationToken).ConfigureAwait(false);
                    SetToken(next, source, initialized);
                    outcomes.Add(new GoogleWorkspaceChangeSourceResult(source, GoogleWorkspaceChangeReadStatus.Initialized, 0, initialized, null));
                    report.Add(TranslationSeverity.Info, "ChangeTracking", $"Initialized the {source.Key} Drive change cursor; no historical changes were inferred.", code: "SYNC.CHANGES.SOURCE_INITIALIZED", action: TranslationAction.Preserve, targetId: source.Key);
                    continue;
                }

                try {
                    var sourceChanges = new List<GoogleWorkspaceTrackedChange>();
                    string pageToken = startToken!;
                    string? completedToken = null;
                    for (int page = 0; page < options.MaxPagesPerSource; page++) {
                        bool includeItemsFromAllDrives = source.Kind == GoogleWorkspaceChangeSourceKind.SharedDrive
                            || partitionedDriveIds.Count == 0
                            || hasUnpartitionedSharedDrives;
                        GoogleDriveChangeList response = await _drive.ListChangesAsync(pageToken, new GoogleDriveChangeListOptions {
                            DriveId = source.DriveId,
                            PageSize = options.PageSize,
                            IncludeRemoved = options.IncludeRemoved,
                            IncludeCorpusRemovals = options.IncludeCorpusRemovals,
                            IncludeItemsFromAllDrives = includeItemsFromAllDrives,
                        }, report, cancellationToken).ConfigureAwait(false);
                        IEnumerable<GoogleDriveChange> pageChanges = response.Changes;
                        if (source.Kind == GoogleWorkspaceChangeSourceKind.User && includeItemsFromAllDrives) {
                            pageChanges = pageChanges.Where(change => !partitionedDriveIds.Contains(change.DriveId ?? change.File?.DriveId ?? string.Empty));
                        }
                        GoogleWorkspaceTrackedChange[] materializedPage = pageChanges
                            .Select(change => new GoogleWorkspaceTrackedChange(source, change))
                            .ToArray();
                        if (materializedPage.Length > options.MaxChangesPerSource - sourceChanges.Count) {
                            throw new InvalidDataException($"Google Drive change pagination for {source.Key} exceeded the configured {options.MaxChangesPerSource} change limit.");
                        }
                        if (materializedPage.Length > options.MaxTotalChanges - changes.Count - sourceChanges.Count) {
                            throw new InvalidDataException($"Google Drive change pagination exceeded the configured {options.MaxTotalChanges} total change limit.");
                        }
                        sourceChanges.AddRange(materializedPage);
                        if (!string.IsNullOrWhiteSpace(response.NextPageToken)) {
                            pageToken = response.NextPageToken!;
                            continue;
                        }
                        completedToken = response.NewStartPageToken;
                        if (string.IsNullOrWhiteSpace(completedToken)) throw new InvalidOperationException($"Google Drive completed the {source.Key} change feed without returning a new start page token.");
                        break;
                    }
                    if (string.IsNullOrWhiteSpace(completedToken)) throw new InvalidOperationException($"Google Drive change pagination for {source.Key} exceeded the configured {options.MaxPagesPerSource} page limit.");
                    changes.AddRange(sourceChanges);
                    SetToken(next, source, completedToken!);
                    outcomes.Add(new GoogleWorkspaceChangeSourceResult(source, GoogleWorkspaceChangeReadStatus.Completed, sourceChanges.Count, completedToken, null));
                } catch (Exception exception) when (!(exception is OperationCanceledException)) {
                    outcomes.Add(new GoogleWorkspaceChangeSourceResult(source, GoogleWorkspaceChangeReadStatus.Failed, 0, startToken, exception));
                    report.Add(TranslationSeverity.Error, "ChangeTracking", $"The {source.Key} Drive change feed failed and its checkpoint was not advanced: {exception.Message}", code: "SYNC.CHANGES.SOURCE_FAILED", action: TranslationAction.Fail, targetId: source.Key);
                    if (!options.ContinueOnSourceFailure) throw;
                }
            }
            return new GoogleWorkspaceChangeReadResult(changes, next, outcomes, report);
        }

        /// <summary>Disposes the underlying Drive client without taking ownership of the session.</summary>
        public void Dispose() => _drive.Dispose();

        private static IEnumerable<string> NormalizeDriveIds(IEnumerable<string>? ids) => (ids ?? Array.Empty<string>())
            .Where(id => !string.IsNullOrWhiteSpace(id)).Select(id => id.Trim()).Distinct(StringComparer.Ordinal).OrderBy(id => id, StringComparer.Ordinal);

        private static string? TokenFor(GoogleWorkspaceSyncCheckpoint checkpoint, GoogleWorkspaceChangeSource source) {
            if (source.Kind == GoogleWorkspaceChangeSourceKind.User) return checkpoint.UserChangeToken;
            return source.DriveId != null && checkpoint.SharedDriveChangeTokens.TryGetValue(source.DriveId, out string? token) ? token : null;
        }

        private static void SetToken(GoogleWorkspaceSyncCheckpoint checkpoint, GoogleWorkspaceChangeSource source, string token) {
            if (source.Kind == GoogleWorkspaceChangeSourceKind.User) checkpoint.UserChangeToken = token;
            else checkpoint.SharedDriveChangeTokens[source.DriveId!] = token;
        }
    }
}
