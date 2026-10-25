using System.Security.Cryptography;
using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

public static partial class OfficeConversionJobRunner {
    /// <summary>Records one saved-job attempt before executing, retaining bounded observations separately from authoritative checkpoints.</summary>
    /// <remarks>The report destination must be new. A final recording failure is returned in RecordingError without discarding conversion outcomes.
    /// The optional host retention callback can disable recording without cancelling conversion; an already-created unchanged report is then removed.</remarks>
    public static async Task<OfficeConversionJobRunReport> RunRecordedAsync(OfficeConversionJob job, string jobPath, string reportPath,
        IOfficeWorkflowRunner runner, OfficeConversionJobRuntimeOptions? runtime = null,
        IProgress<OfficeConversionBatchItemResult>? progress = null, CancellationToken cancellationToken = default,
        IOfficeWorkflowPublicationGuard? publicationGuard = null, Func<bool>? retainReport = null, bool allowEmpty = false) {
        ArgumentNullException.ThrowIfNull(job);
        ArgumentNullException.ThrowIfNull(runner);
        string source = Path.GetFullPath(jobPath), destination = Path.GetFullPath(reportPath);
        OfficeConversionJob saved = OfficeConversionJobSerializer.ReadSnapshot(source, out string jobSha256);
        saved.RetryFailed = job.RetryFailed;
        if (OfficeConversionJobSerializer.Serialize(saved) != OfficeConversionJobSerializer.Serialize(job))
            throw new InvalidDataException("The execution settings differ from the saved job. Save or review the definition before recording this attempt.");
        job = saved;
        string baseDirectory = Path.GetDirectoryName(source)!;
        OfficeConversionBatchRequest request = job.CreateRequest(baseDirectory);
        ValidateReportPath(request, source, destination);
        var report = new OfficeConversionJobRunReport {
            RunId = Guid.NewGuid().ToString("N"), StartedUtc = DateTimeOffset.UtcNow, JobPath = source, JobSha256 = jobSha256,
            OutputDirectory = request.OutputDirectory, CheckpointDirectory = request.CheckpointDirectory,
            RetryFailed = job.RetryFailed, AllowEmpty = allowEmpty, State = OfficeConversionJobRunState.Running
        };
        var collector = new OfficeConversionBatchReportCollector();
        using var recording = new RunRecording(destination, report.RunId, retainReport, publicationGuard);
        await recording.BeginAsync(report).ConfigureAwait(false);
        try {
            OfficeConversionBatchResult summary = await RunAsync(job, baseDirectory, runner, runtime,
                new RecordedProgress(collector, progress), cancellationToken, publicationGuard).ConfigureAwait(false);
            OfficeConversionBatchReport batch = collector.Complete(summary, allowEmpty);
            report = report with {
                FinishedUtc = DateTimeOffset.UtcNow, Batch = batch,
                State = summary.Cancelled ? OfficeConversionJobRunState.Cancelled :
                    batch.Succeeded ? OfficeConversionJobRunState.Completed : OfficeConversionJobRunState.Failed
            };
        } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
            report = report with { FinishedUtc = DateTimeOffset.UtcNow, State = OfficeConversionJobRunState.Cancelled,
                Batch = collector.Complete(collector.ObservedSummary(cancelled: true), allowEmpty) };
        } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
            string message = "Saved-job execution failed (" + error.GetType().Name + "). Review the original execution error before retrying.";
            report = report with { FinishedUtc = DateTimeOffset.UtcNow, State = OfficeConversionJobRunState.Failed,
                FailureSummary = message, Batch = collector.Complete(collector.ObservedSummary(false, message), allowEmpty) };
            _ = await recording.FinishAsync(report).ConfigureAwait(false);
            throw;
        }
        return await recording.FinishAsync(report).ConfigureAwait(false);
    }

    private static void ValidateReportPath(OfficeConversionBatchRequest request, string jobPath, string reportPath) {
        OfficeConversionBatchExecutor.EnsureNoLinks(reportPath);
        OfficeConversionBatchExecutor.EnsureNoLinks(reportPath + ".lease");
        if (OfficePathIdentity.AreEquivalent(reportPath, jobPath)) throw new InvalidDataException("The report must not replace its saved job.");
        foreach (string root in new[] { request.InputDirectory, request.OutputDirectory, request.CheckpointDirectory }.OfType<string>()) {
            OfficeConversionBatchExecutor.EnsureNoLinks(root);
            if (OfficePathIdentity.IsSameOrDescendant(reportPath, root) || OfficePathIdentity.IsSameOrDescendant(root, reportPath))
                throw new InvalidDataException("Save run reports outside the source, output and checkpoint trees.");
        }
        foreach (string input in request.InputPaths ?? []) {
            OfficeConversionBatchExecutor.EnsureNoLinks(input);
            if (OfficePathIdentity.IsSameOrDescendant(reportPath, input))
                throw new InvalidDataException("A run report must not replace a source or source-package member.");
        }
    }

    private sealed class RecordedProgress(OfficeConversionBatchReportCollector collector, IProgress<OfficeConversionBatchItemResult>? next)
        : IProgress<OfficeConversionBatchItemResult> {
        public void Report(OfficeConversionBatchItemResult value) { collector.Report(value); next?.Report(value); }
    }

    private sealed class RunRecording(string path, string runId, Func<bool>? retain, IOfficeWorkflowPublicationGuard? guard) : IDisposable {
        private FileStream? _lease;
        private string? _expectedHash;
        private bool _started;
        internal async Task BeginAsync(OfficeConversionJobRunReport report) {
            if (retain?.Invoke() == false) return;
            if (guard != null && !await guard.CanPublishAsync(path, false, CancellationToken.None).ConfigureAwait(false))
                throw new UnauthorizedAccessException("Run report destination is protected by the host publication policy.");
            if (retain?.Invoke() == false) return;
            OfficeConversionBatchExecutor.EnsureNoLinks(path);
            _lease = new FileStream(path + ".lease", FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
            _lease.Write(Encoding.ASCII.GetBytes(runId)); _lease.Flush(flushToDisk: true);
            byte[] bytes = Encoding.UTF8.GetBytes(OfficeConversionJobRunReportSerializer.Serialize(report));
            OfficeFileCommit.WriteAtomically(path, stream => {
                stream.Write(bytes); if (stream is FileStream file) file.Flush(flushToDisk: true);
            }, conflictPolicy: OfficeFileCommit.ConflictPolicy.FailIfExists);
            _expectedHash = Convert.ToHexString(SHA256.HashData(bytes));
            _started = true;
        }

        internal async Task<OfficeConversionJobRunReport> FinishAsync(OfficeConversionJobRunReport report) {
            try {
                if (!_started) return report;
                if (retain?.Invoke() == false) {
                    OfficeConversionJobRunReportSerializer.DeleteOwned(path, _expectedHash!);
                    return report;
                }
                if (guard != null && !await guard.CanPublishAsync(path, false, CancellationToken.None).ConfigureAwait(false))
                    throw new UnauthorizedAccessException("Run report destination is protected by the host publication policy.");
                if (retain?.Invoke() == false) {
                    OfficeConversionJobRunReportSerializer.DeleteOwned(path, _expectedHash!);
                    return report;
                }
                byte[] bytes = Encoding.UTF8.GetBytes(OfficeConversionJobRunReportSerializer.Serialize(report));
                await OfficeFileCommit.WriteIfUnchangedAsync(path, async (stream, token) => {
                    await stream.WriteAsync(bytes, token).ConfigureAwait(false);
                    if (stream is FileStream file) file.Flush(flushToDisk: true);
                }, candidate => retain?.Invoke() != false && OfficeConversionJobRunReportSerializer.MatchesSnapshot(candidate, _expectedHash!)).ConfigureAwait(false);
                return report;
            } catch (Exception error) when (error is not OutOfMemoryException and not StackOverflowException) {
                try {
                    if (retain?.Invoke() == false) {
                        OfficeConversionJobRunReportSerializer.DeleteOwned(path, _expectedHash!);
                        return report;
                    }
                } catch (Exception cleanupError) when (cleanupError is not OutOfMemoryException and not StackOverflowException) {
                    return report with { RecordingError = BoundError(cleanupError) };
                }
                return report with { RecordingError = BoundError(error) };
            }
        }

        private static string BoundError(Exception error) => error.Message.Length <= 4096 ? error.Message : error.Message[..4096];

        public void Dispose() {
            if (_lease == null) return;
            _lease.Dispose();
            try { OfficeConversionJobRunReportSerializer.DeleteLease(path + ".lease", runId); }
            catch (Exception error) when (error is IOException or UnauthorizedAccessException or InvalidDataException) { }
        }
    }
}
