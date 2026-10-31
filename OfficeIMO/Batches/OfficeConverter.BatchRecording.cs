using OfficeIMO.Internal;
using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Core.Internal;

namespace OfficeIMO;

public static partial class OfficeConverter {
    private sealed class RunRecording(string path, string runId, Func<bool>? retain, IOfficePublicationGuard? guard) : IDisposable {
        private FileStream? _lease;
        private string? _expectedHash;
        private bool _started;
        internal async Task BeginAsync(OfficeBatchReport report) {
            if (retain?.Invoke() == false) return;
            if (guard != null && !await guard.CanPublishAsync(path, false, CancellationToken.None).ConfigureAwait(false))
                throw new UnauthorizedAccessException("Run report destination is protected by the host publication policy.");
            if (retain?.Invoke() == false) return;
            OfficeBatchExecutor.EnsureNoLinks(path);
            _lease = new FileStream(path + ".lease", FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
            OfficeRuntimeCompatibility.WriteAll(_lease, Encoding.ASCII.GetBytes(runId)); _lease.Flush(flushToDisk: true);
            byte[] bytes = Encoding.UTF8.GetBytes(OfficeBatchReportStore.Serialize(report));
            OfficeFileCommit.WriteAtomically(path, stream => {
                OfficeRuntimeCompatibility.WriteAll(stream, bytes); if (stream is FileStream file) file.Flush(flushToDisk: true);
            }, conflictPolicy: OfficeFileCommit.ConflictPolicy.FailIfExists);
            _expectedHash = OfficeRuntimeCompatibility.ToHexString(OfficeRuntimeCompatibility.HashSha256(bytes));
            _started = true;
        }

        internal async Task<OfficeBatchReport> FinishAsync(OfficeBatchReport report) {
            try {
                if (!_started) return report;
                if (retain?.Invoke() == false) {
                    OfficeBatchReportStore.DeleteOwned(path, _expectedHash!);
                    return report;
                }
                if (guard != null && !await guard.CanPublishAsync(path, false, CancellationToken.None).ConfigureAwait(false))
                    throw new UnauthorizedAccessException("Run report destination is protected by the host publication policy.");
                if (retain?.Invoke() == false) {
                    OfficeBatchReportStore.DeleteOwned(path, _expectedHash!);
                    return report;
                }
                byte[] bytes = Encoding.UTF8.GetBytes(OfficeBatchReportStore.Serialize(report));
                await OfficeFileCommit.WriteIfUnchangedAsync(path, async (stream, token) => {
                    await stream.WriteAsync(bytes, 0, bytes.Length, token).ConfigureAwait(false);
                    if (stream is FileStream file) file.Flush(flushToDisk: true);
                }, candidate => retain?.Invoke() != false && OfficeBatchReportStore.MatchesSnapshot(candidate, _expectedHash!)).ConfigureAwait(false);
                return report;
            } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
                try {
                    if (retain?.Invoke() == false) {
                        OfficeBatchReportStore.DeleteOwned(path, _expectedHash!);
                        return report;
                    }
                } catch (Exception cleanupError) when (cleanupError is not OutOfMemoryException and not StackOverflowException) {
                    return report with { RecordingError = BoundError(cleanupError) };
                }
                return report with { RecordingError = BoundError(error) };
            }
        }

        private static string BoundError(Exception error) => error.Message.Length <= 4096 ? error.Message : error.Message.Substring(0, 4096);

        public void Dispose() {
            if (_lease == null) return;
            _lease.Dispose();
            try { OfficeBatchReportStore.DeleteLease(path + ".lease", runId); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException or InvalidDataException) { }
        }
    }
}
