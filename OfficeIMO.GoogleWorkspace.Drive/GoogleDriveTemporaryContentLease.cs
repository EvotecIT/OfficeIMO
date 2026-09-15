using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.GoogleWorkspace.Drive {
    /// <summary>Cleanup state for one temporary public Drive file.</summary>
    public enum GoogleDriveCleanupStatus {
        /// <summary>Cleanup has not completed.</summary>
        Pending = 0,
        /// <summary>The temporary file was deleted successfully.</summary>
        Deleted = 1,
        /// <summary>Deletion failed and requires follow-up.</summary>
        Failed = 2,
    }

    /// <summary>Cleanup result for one temporary Drive file.</summary>
    public sealed class GoogleDriveCleanupEntry {
        /// <summary>Gets or sets the Drive file identifier.</summary>
        public string FileId { get; set; } = string.Empty;
        /// <summary>Gets or sets the current cleanup state.</summary>
        public GoogleDriveCleanupStatus Status { get; set; }
        /// <summary>Gets or sets the deletion error when cleanup failed.</summary>
        public string? Error { get; set; }
    }

    /// <summary>Tracks cleanup outcomes for temporary Drive content created by an operation.</summary>
    public sealed class GoogleDriveCleanupReport {
        private readonly List<GoogleDriveCleanupEntry> _entries = new List<GoogleDriveCleanupEntry>();

        /// <summary>Gets cleanup entries in creation order.</summary>
        public IReadOnlyList<GoogleDriveCleanupEntry> Entries => _entries;
        /// <summary>Gets whether at least one temporary file could not be deleted.</summary>
        public bool HasFailures => _entries.Any(entry => entry.Status == GoogleDriveCleanupStatus.Failed);

        internal GoogleDriveCleanupEntry Add(string fileId) {
            var entry = new GoogleDriveCleanupEntry { FileId = fileId, Status = GoogleDriveCleanupStatus.Pending };
            _entries.Add(entry);
            return entry;
        }
    }

    /// <summary>Owns a short-lived publicly readable Drive file and its explicit cleanup result.</summary>
    public sealed class GoogleDriveTemporaryContentLease {
        private readonly GoogleDriveClient _client;
        private readonly TranslationReport _report;
        private readonly GoogleDriveCleanupEntry _cleanupEntry;
        private readonly SemaphoreSlim _cleanupGate = new SemaphoreSlim(1, 1);
        private bool _cleaned;

        private GoogleDriveTemporaryContentLease(
            GoogleDriveClient client,
            TranslationReport report,
            GoogleDriveFile file,
            string publicUri,
            GoogleDriveCleanupReport cleanupReport) {
            _client = client;
            _report = report;
            File = file;
            PublicUri = publicUri;
            CleanupReport = cleanupReport;
            _cleanupEntry = cleanupReport.Add(file.Id ?? string.Empty);
        }

        /// <summary>Gets metadata for the temporary Drive file.</summary>
        public GoogleDriveFile File { get; }
        /// <summary>Gets the public download URI exposed while the lease is active.</summary>
        public string PublicUri { get; }
        /// <summary>Gets the cleanup report associated with this lease.</summary>
        public GoogleDriveCleanupReport CleanupReport { get; }

        /// <summary>Uploads content, grants a non-discoverable public-reader permission, and returns a cleanup lease.</summary>
        /// <remarks>If permission creation fails after upload, the method attempts immediate best-effort deletion before rethrowing.</remarks>
        public static async Task<GoogleDriveTemporaryContentLease> CreatePublicReadLeaseAsync(
            GoogleDriveClient client,
            byte[] content,
            GoogleDriveUploadOptions options,
            TranslationReport? report = null,
            CancellationToken cancellationToken = default) {
            if (client == null) throw new ArgumentNullException(nameof(client));
            report ??= new TranslationReport();
            GoogleDriveFile? file = null;
            var cleanupReport = new GoogleDriveCleanupReport();
            try {
                file = content.LongLength <= GoogleDriveClient.MultipartUploadLimitBytes
                    ? await client.UploadMultipartAsync(content, options, report, cancellationToken).ConfigureAwait(false)
                    : await client.UploadResumableAsync(content, options, report, cancellationToken).ConfigureAwait(false);
                if (string.IsNullOrWhiteSpace(file.Id)) {
                    throw new InvalidOperationException("Temporary Drive upload did not return a file id.");
                }

                await client.CreatePermissionAsync(
                    file.Id!,
                    new GoogleDrivePermissionCreateOptions {
                        Type = "anyone",
                        Role = "reader",
                        AllowFileDiscovery = false,
                        SendNotificationEmail = false,
                    },
                    report,
                    cancellationToken).ConfigureAwait(false);
                report.Add(
                    TranslationSeverity.Info,
                    "TemporaryContent",
                    "Created a short-lived public Drive object that must be cleaned after the target Google service fetches it.",
                    code: "DRIVE.TEMPORARY_CONTENT.PUBLIC_LEASE_CREATED",
                    action: TranslationAction.Preserve,
                    targetId: file.Id);
                return new GoogleDriveTemporaryContentLease(
                    client,
                    report,
                    file,
                    "https://drive.google.com/uc?export=download&id=" + Uri.EscapeDataString(file.Id!),
                    cleanupReport);
            } catch {
                if (!string.IsNullOrWhiteSpace(file?.Id)) {
                    var entry = cleanupReport.Add(file!.Id!);
                    await TryDeleteAsync(client, file.Id!, report, entry, CancellationToken.None).ConfigureAwait(false);
                }

                throw;
            }
        }

        /// <summary>Deletes the temporary file once and returns the persistent cleanup result.</summary>
        /// <remarks>Failed deletion may be retried by calling this method again; successful cleanup is idempotent.</remarks>
        public async Task<GoogleDriveCleanupReport> CleanupAsync(CancellationToken cancellationToken = default) {
            await _cleanupGate.WaitAsync(cancellationToken).ConfigureAwait(false);
            try {
                if (_cleaned) return CleanupReport;
                _cleaned = await TryDeleteAsync(
                    _client,
                    File.Id ?? string.Empty,
                    _report,
                    _cleanupEntry,
                    cancellationToken).ConfigureAwait(false);
                return CleanupReport;
            } finally {
                _cleanupGate.Release();
            }
        }

        private static async Task<bool> TryDeleteAsync(
            GoogleDriveClient client,
            string fileId,
            TranslationReport report,
            GoogleDriveCleanupEntry entry,
            CancellationToken cancellationToken) {
            try {
                await client.DeleteFileAsync(fileId, report, cancellationToken).ConfigureAwait(false);
                entry.Status = GoogleDriveCleanupStatus.Deleted;
                entry.Error = null;
                report.Add(
                    TranslationSeverity.Info,
                    "TemporaryContent",
                    "Deleted the temporary public Drive object.",
                    code: "DRIVE.TEMPORARY_CONTENT.CLEANED",
                    action: TranslationAction.Preserve,
                    targetId: fileId);
                return true;
            } catch (Exception exception) when (!(exception is OperationCanceledException)) {
                entry.Status = GoogleDriveCleanupStatus.Failed;
                entry.Error = exception.Message;
                report.Add(
                    TranslationSeverity.Error,
                    "TemporaryContent",
                    $"Temporary public Drive object '{fileId}' could not be deleted: {exception.Message}",
                    code: "DRIVE.TEMPORARY_CONTENT.CLEANUP_FAILED",
                    action: TranslationAction.Fail,
                    targetId: fileId);
                return false;
            }
        }
    }
}
