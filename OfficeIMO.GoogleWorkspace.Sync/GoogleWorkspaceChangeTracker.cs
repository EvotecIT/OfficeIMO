using OfficeIMO.GoogleWorkspace.Drive;
using System.IO;

namespace OfficeIMO.GoogleWorkspace.Sync {
    public enum GoogleWorkspaceChangeSourceKind { User = 0, SharedDrive = 1 }
    public enum GoogleWorkspaceChangeReadStatus { Completed = 0, Initialized = 1, Failed = 2 }

    public sealed class GoogleWorkspaceChangeSource {
        internal GoogleWorkspaceChangeSource(GoogleWorkspaceChangeSourceKind kind, string? driveId) { Kind = kind; DriveId = driveId; }
        public GoogleWorkspaceChangeSourceKind Kind { get; }
        public string? DriveId { get; }
        public string Key => Kind == GoogleWorkspaceChangeSourceKind.User ? "user" : "drive:" + DriveId;
    }

    public sealed class GoogleWorkspaceTrackedChange {
        internal GoogleWorkspaceTrackedChange(GoogleWorkspaceChangeSource source, GoogleDriveChange change) { Source = source; Change = change; }
        public GoogleWorkspaceChangeSource Source { get; }
        public GoogleDriveChange Change { get; }
    }

    public sealed class GoogleWorkspaceChangeReadOptions {
        public IList<string> SharedDriveIds { get; } = new List<string>();
        public int PageSize { get; set; } = 100;
        public int MaxPagesPerSource { get; set; } = 10000;
        public int MaxChangesPerSource { get; set; } = 10_000;
        public int MaxTotalChanges { get; set; } = 50_000;
        public bool IncludeRemoved { get; set; } = true;
        public bool IncludeCorpusRemovals { get; set; } = true;
        public bool ContinueOnSourceFailure { get; set; } = true;
    }

    public sealed class GoogleWorkspaceChangeSourceResult {
        internal GoogleWorkspaceChangeSourceResult(GoogleWorkspaceChangeSource source, GoogleWorkspaceChangeReadStatus status, int changeCount, string? nextToken, Exception? exception) {
            Source = source; Status = status; ChangeCount = changeCount; NextToken = nextToken; Exception = exception;
        }
        public GoogleWorkspaceChangeSource Source { get; }
        public GoogleWorkspaceChangeReadStatus Status { get; }
        public int ChangeCount { get; }
        public string? NextToken { get; }
        public Exception? Exception { get; }
    }

    public sealed class GoogleWorkspaceChangeReadResult {
        internal GoogleWorkspaceChangeReadResult(IReadOnlyList<GoogleWorkspaceTrackedChange> changes, GoogleWorkspaceSyncCheckpoint checkpoint, IReadOnlyList<GoogleWorkspaceChangeSourceResult> sources, TranslationReport report) {
            Changes = changes; NextCheckpoint = checkpoint; Sources = sources; Report = report;
        }
        public IReadOnlyList<GoogleWorkspaceTrackedChange> Changes { get; }
        public GoogleWorkspaceSyncCheckpoint NextCheckpoint { get; }
        public IReadOnlyList<GoogleWorkspaceChangeSourceResult> Sources { get; }
        public TranslationReport Report { get; }
        public bool HasFailures => Sources.Any(source => source.Status == GoogleWorkspaceChangeReadStatus.Failed);
    }

    /// <summary>Consumes complete Drive change pages and advances each source token only after that source succeeds.</summary>
    public sealed class GoogleWorkspaceChangeTracker : IDisposable {
        private readonly GoogleDriveClient _drive;

        public GoogleWorkspaceChangeTracker(GoogleWorkspaceSession session, GoogleDriveClientOptions? options = null) {
            _drive = new GoogleDriveClient(session ?? throw new ArgumentNullException(nameof(session)), options);
        }

        public async Task<GoogleWorkspaceSyncCheckpoint> InitializeAsync(IEnumerable<string>? sharedDriveIds = null, CancellationToken cancellationToken = default) {
            var checkpoint = new GoogleWorkspaceSyncCheckpoint {
                UserChangeToken = await _drive.GetStartPageTokenAsync(cancellationToken: cancellationToken).ConfigureAwait(false),
            };
            foreach (string driveId in NormalizeDriveIds(sharedDriveIds)) {
                checkpoint.SharedDriveChangeTokens[driveId] = await _drive.GetStartPageTokenAsync(driveId, cancellationToken: cancellationToken).ConfigureAwait(false);
            }
            return checkpoint;
        }

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
