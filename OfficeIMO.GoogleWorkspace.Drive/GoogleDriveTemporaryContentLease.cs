using OfficeIMO.GoogleWorkspace;

namespace OfficeIMO.GoogleWorkspace.Drive {
    public enum GoogleDriveCleanupStatus {
        Pending = 0,
        Deleted = 1,
        Failed = 2,
    }

    public sealed class GoogleDriveCleanupEntry {
        public string FileId { get; set; } = string.Empty;
        public GoogleDriveCleanupStatus Status { get; set; }
        public string? Error { get; set; }
    }

    public sealed class GoogleDriveCleanupReport {
        private readonly List<GoogleDriveCleanupEntry> _entries = new List<GoogleDriveCleanupEntry>();

        public IReadOnlyList<GoogleDriveCleanupEntry> Entries => _entries;
        public bool HasFailures => _entries.Any(entry => entry.Status == GoogleDriveCleanupStatus.Failed);

        internal GoogleDriveCleanupEntry Add(string fileId) {
            var entry = new GoogleDriveCleanupEntry { FileId = fileId, Status = GoogleDriveCleanupStatus.Pending };
            _entries.Add(entry);
            return entry;
        }
    }

    public sealed class GoogleDriveTemporaryContentLease {
        private readonly GoogleDriveClient _client;
        private readonly TranslationReport _report;
        private readonly GoogleDriveCleanupEntry _cleanupEntry;
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

        public GoogleDriveFile File { get; }
        public string PublicUri { get; }
        public GoogleDriveCleanupReport CleanupReport { get; }

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
                file = await client.UploadMultipartAsync(content, options, report, cancellationToken).ConfigureAwait(false);
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

        public async Task<GoogleDriveCleanupReport> CleanupAsync(CancellationToken cancellationToken = default) {
            if (_cleaned) return CleanupReport;
            _cleaned = true;
            await TryDeleteAsync(_client, File.Id ?? string.Empty, _report, _cleanupEntry, cancellationToken).ConfigureAwait(false);
            return CleanupReport;
        }

        private static async Task TryDeleteAsync(
            GoogleDriveClient client,
            string fileId,
            TranslationReport report,
            GoogleDriveCleanupEntry entry,
            CancellationToken cancellationToken) {
            try {
                await client.DeleteFileAsync(fileId, report, cancellationToken).ConfigureAwait(false);
                entry.Status = GoogleDriveCleanupStatus.Deleted;
                report.Add(
                    TranslationSeverity.Info,
                    "TemporaryContent",
                    "Deleted the temporary public Drive object.",
                    code: "DRIVE.TEMPORARY_CONTENT.CLEANED",
                    action: TranslationAction.Preserve,
                    targetId: fileId);
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
            }
        }
    }
}
