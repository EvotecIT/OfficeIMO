using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO.Workflows;

/// <summary>Strict bounded run-report JSON and local lease-aware inspection. Reading never executes or resumes a job.</summary>
public static class OfficeConversionJobRunReportSerializer {
    /// <summary>Maximum UTF-8 size of a recorded attempt, including item observations.</summary>
    public const int MaximumReportBytes = OfficeConversionBatchReportSerializer.MaximumReportBytes;

    /// <summary>Serializes observational state without accessing its source or output files.</summary>
    public static string Serialize(OfficeConversionJobRunReport report) {
        Validate(report);
        return OfficeConversionBatchReportSerializer.SerializeBounded(report,
            OfficeConversionBatchReportSerializer.Context.OfficeConversionJobRunReport);
    }

    /// <summary>Reads strict JSON. Lease interpretation is performed only by ReadFile.</summary>
    public static OfficeConversionJobRunReport Deserialize(string json) {
        OfficeConversionBatchReportSerializer.CheckJson(json);
        var report = JsonSerializer.Deserialize(json, OfficeConversionBatchReportSerializer.Context.OfficeConversionJobRunReport)
            ?? throw new InvalidDataException("The saved-job run report is empty.");
        Validate(report);
        return report;
    }

    /// <summary>Reads a bounded regular file. An unfinished record without a live local lease is returned as Interrupted.</summary>
    public static OfficeConversionJobRunReport ReadFile(string path) {
        var report = ReadSnapshot(path, out _);
        return report.State == OfficeConversionJobRunState.Running && !IsActive(path)
            ? report with { State = OfficeConversionJobRunState.Interrupted } : report;
    }

    /// <summary>Returns whether the local report lease is held or cannot safely be inspected. Timestamps are never activity evidence.</summary>
    public static bool IsActive(string path) {
        string leasePath = Path.GetFullPath(path) + ".lease";
        OfficeConversionBatchExecutor.EnsureNoLinks(leasePath);
        try {
            using var lease = new FileStream(leasePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return false;
        } catch (FileNotFoundException) { return false; }
        catch (DirectoryNotFoundException) { return false; }
        catch (IOException) { return true; }
        catch (UnauthorizedAccessException) { return true; }
    }

    /// <summary>Deletes only a valid inactive report and its matching lease. Missing reports succeed; live reports return false; malformed reports are preserved and rejected.</summary>
    public static bool TryDeleteFinished(string path) {
        string fullPath = Path.GetFullPath(path);
        OfficeConversionBatchExecutor.EnsureNoLinks(fullPath);
        if (!File.Exists(fullPath)) return true;
        var report = ReadSnapshot(fullPath, out string expectedHash);
        string leasePath = fullPath + ".lease";
        OfficeConversionBatchExecutor.EnsureNoLinks(leasePath);
        FileStream lease;
        try { lease = new FileStream(leasePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
        catch (IOException) { return false; }
        bool deleteLease = false;
        try {
            if (lease.Length > 0) {
                if (lease.Length != 32) throw new InvalidDataException("The report lease does not belong to this attempt.");
                byte[] owner = new byte[32]; lease.ReadExactly(owner);
                if (Encoding.ASCII.GetString(owner) != report.RunId) throw new InvalidDataException("The report lease does not belong to this attempt.");
            } else {
                lease.Write(Encoding.ASCII.GetBytes(report.RunId)); lease.Flush(flushToDisk: true);
            }
            deleteLease = true;
            DeleteOwned(fullPath, expectedHash);
            return true;
        } finally {
            lease.Dispose();
            if (deleteLease) DeleteLease(leasePath, report.RunId);
        }
    }

    internal static OfficeConversionJobRunReport ReadSnapshot(string path, out string sha256) {
        string fullPath = Path.GetFullPath(path);
        OfficeConversionBatchExecutor.EnsureNoLinks(fullPath);
        using var stream = OfficePathIdentity.OpenRegularFileForRead(fullPath, Path.GetDirectoryName(fullPath)!, 81920);
        if (stream.Length > MaximumReportBytes) throw new InvalidDataException("The report exceeds its 20 MiB limit.");
        using var bytes = new BoundedReportStream(MaximumReportBytes);
        stream.CopyTo(bytes);
        byte[] content = bytes.ToArray();
        sha256 = Convert.ToHexString(SHA256.HashData(content));
        return Deserialize(new UTF8Encoding(false, true).GetString(content).TrimStart('\uFEFF'));
    }

    internal static bool MatchesSnapshot(string path, string expectedHash) {
        try { _ = ReadSnapshot(path, out string actual); return actual == expectedHash; }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException or JsonException or InvalidDataException or NotSupportedException) { return false; }
    }

    internal static void DeleteOwned(string path, string expectedHash) {
        OfficeConversionBatchExecutor.EnsureNoLinks(path);
        if (!OfficeFileCommit.TryDeleteIfUnchanged(path, candidate => MatchesSnapshot(candidate, expectedHash)))
            throw new IOException("The run report changed; its replacement was preserved.");
    }

    internal static void DeleteLease(string path, string runId) {
        OfficeConversionBatchExecutor.EnsureNoLinks(path);
        _ = OfficeFileCommit.TryDeleteIfUnchanged(path, candidate => {
            OfficeConversionBatchExecutor.EnsureNoLinks(candidate);
            using var stream = OfficePathIdentity.OpenRegularFileForRead(candidate, Path.GetDirectoryName(candidate)!, 32);
            if (stream.Length != 32) return false;
            byte[] owner = new byte[33];
            int count = stream.ReadAtLeast(owner, 32, throwOnEndOfStream: false);
            return count == 32 && Encoding.ASCII.GetString(owner, 0, count) == runId;
        });
    }

    internal static void Validate(OfficeConversionJobRunReport report) {
        ArgumentNullException.ThrowIfNull(report);
        if (report.Schema != OfficeConversionJobRunReport.CurrentSchema) throw new NotSupportedException("Unsupported saved-job run schema.");
        if (!Guid.TryParseExact(report.RunId, "N", out _) || report.JobSha256 == null || report.JobSha256.Length != 64 ||
            report.JobSha256.Any(character => !char.IsAsciiHexDigit(character)) || !Enum.IsDefined(report.State))
            throw new InvalidDataException("The saved-job run identity or state is invalid.");
        CheckPath(report.JobPath); CheckPath(report.OutputDirectory);
        if (report.CheckpointDirectory != null) CheckPath(report.CheckpointDirectory);
        if (report.StartedUtc == default || report.StartedUtc.Offset != TimeSpan.Zero ||
            report.FinishedUtc is { } finished && (finished.Offset != TimeSpan.Zero || finished < report.StartedUtc))
            throw new InvalidDataException("Run report times must be ordered UTC instants.");
        bool unfinished = report.State is OfficeConversionJobRunState.Running or OfficeConversionJobRunState.Interrupted;
        if (unfinished && (report.FinishedUtc != null || report.Batch != null) || !unfinished && report.FinishedUtc == null)
            throw new InvalidDataException("The run report lifecycle fields are inconsistent.");
        if (report.State is OfficeConversionJobRunState.Completed or OfficeConversionJobRunState.Cancelled && report.Batch == null)
            throw new InvalidDataException("A terminal conversion outcome requires its batch summary.");
        if (report.FailureSummary?.Length > 4096 || report.RecordingError?.Length > 4096)
            throw new InvalidDataException("The run report error text is oversized.");
        if (report.Batch != null) OfficeConversionBatchReportSerializer.Validate(report.Batch);
    }

    private static void CheckPath(string? path) {
        if (string.IsNullOrWhiteSpace(path) || path.Length > 4096 || path.Any(char.IsControl) || !Path.IsPathFullyQualified(path))
            throw new InvalidDataException("Run report paths must be bounded absolute local paths.");
    }
}
