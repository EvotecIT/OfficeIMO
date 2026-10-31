using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;

namespace OfficeIMO;

/// <summary>Strict bounded run-report JSON and local lease-aware inspection. Reading never executes or resumes a job.</summary>
internal static class OfficeBatchReportStore {
    /// <summary>Maximum UTF-8 size of a recorded attempt, including item observations.</summary>
    public const int MaximumReportBytes = OfficeBatchReportSerializer.MaximumReportBytes;

    /// <summary>Serializes observational state without accessing its source or output files.</summary>
    public static string Serialize(OfficeBatchReport report) {
        Validate(report);
        return OfficeBatchReportSerializer.SerializeBounded(report,
            OfficeBatchReportSerializer.Context.OfficeBatchReport);
    }

    /// <summary>Reads strict JSON. Lease interpretation is performed only by ReadFile.</summary>
    public static OfficeBatchReport Deserialize(string json) {
        var report = OfficeBatchReportSerializer.DeserializeBounded(json,
            OfficeBatchReportSerializer.Context.OfficeBatchReport);
        Validate(report);
        return report;
    }

    /// <summary>Reads a bounded regular file. An unfinished record without a live local lease is returned as Interrupted.</summary>
    public static OfficeBatchReport ReadFile(string path) => ReadFile(path, IsActive);

    // The lease probe is the single filesystem scheduling boundary needed to exercise
    // final publication between the first snapshot and the activity observation.
    internal static OfficeBatchReport ReadFile(string path, Func<string, bool> leaseIsActive) {
        if (leaseIsActive == null) throw new ArgumentNullException(nameof(leaseIsActive));
        var report = ReadSnapshot(path, out _);
        if (report.State != OfficeBatchRunState.Running || leaseIsActive(path)) return report;
        // Final publication precedes lease release. Re-read after observing a free
        // lease so a newly published terminal result is never relabelled Interrupted.
        report = ReadSnapshot(path, out _);
        return report.State == OfficeBatchRunState.Running
            ? report with { State = OfficeBatchRunState.Interrupted, Succeeded = false } : report;
    }

    /// <summary>Returns whether the local report lease is held or cannot safely be inspected. Timestamps are never activity evidence.</summary>
    public static bool IsActive(string path) {
        string leasePath = Path.GetFullPath(path) + ".lease";
        OfficeBatchExecutor.EnsureNoLinks(leasePath);
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
        OfficeBatchExecutor.EnsureNoLinks(fullPath);
        if (!File.Exists(fullPath)) return true;
        var report = ReadSnapshot(fullPath, out string expectedHash);
        if (report.RunId == null) throw new InvalidDataException("Only recorded reports carry a lease identity for owned deletion.");
        string leasePath = fullPath + ".lease";
        OfficeBatchExecutor.EnsureNoLinks(leasePath);
        FileStream lease;
        try { lease = new FileStream(leasePath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None); }
        catch (IOException) { return false; }
        bool deleteLease = false;
        try {
            if (lease.Length > 0) {
                if (lease.Length != 32) throw new InvalidDataException("The report lease does not belong to this attempt.");
                byte[] owner = new byte[32]; OfficeRuntimeCompatibility.ReadAtLeast(lease, owner, owner.Length);
                if (Encoding.ASCII.GetString(owner) != report.RunId) throw new InvalidDataException("The report lease does not belong to this attempt.");
            } else {
                OfficeRuntimeCompatibility.WriteAll(lease, Encoding.ASCII.GetBytes(report.RunId)); lease.Flush(flushToDisk: true);
            }
            deleteLease = true;
            DeleteOwned(fullPath, expectedHash);
            return true;
        } finally {
            lease.Dispose();
            if (deleteLease) DeleteLease(leasePath, report.RunId);
        }
    }

    internal static OfficeBatchReport ReadSnapshot(string path, out string sha256) {
        string fullPath = Path.GetFullPath(path);
        OfficeBatchExecutor.EnsureNoLinks(fullPath);
        using var stream = OfficePathIdentity.OpenRegularFileForRead(fullPath, Path.GetDirectoryName(fullPath)!, 81920);
        if (stream.Length > MaximumReportBytes) throw new InvalidDataException("The report exceeds its 20 MiB limit.");
        using var bytes = new BoundedReportStream(MaximumReportBytes);
        stream.CopyTo(bytes);
        byte[] content = bytes.ToArray();
        sha256 = OfficeRuntimeCompatibility.ToHexString(OfficeRuntimeCompatibility.HashSha256(content));
        string json;
        try { json = new UTF8Encoding(false, true).GetString(content).TrimStart('\uFEFF'); }
        catch (DecoderFallbackException error) { throw new InvalidDataException("The run report does not contain valid UTF-8.", error); }
        return Deserialize(json);
    }

    internal static bool MatchesSnapshot(string path, string expectedHash) {
        try { _ = ReadSnapshot(path, out string actual); return actual == expectedHash; }
        catch (Exception error) when (error is IOException or UnauthorizedAccessException or JsonException or InvalidDataException or NotSupportedException) { return false; }
    }

    internal static void DeleteOwned(string path, string expectedHash) {
        OfficeBatchExecutor.EnsureNoLinks(path);
        if (!OfficeFileCommit.TryDeleteIfUnchanged(path, candidate => MatchesSnapshot(candidate, expectedHash)))
            throw new IOException("The run report changed; its replacement was preserved.");
    }

    internal static void DeleteLease(string path, string runId) {
        OfficeBatchExecutor.EnsureNoLinks(path);
        _ = OfficeFileCommit.TryDeleteIfUnchanged(path, candidate => {
            OfficeBatchExecutor.EnsureNoLinks(candidate);
            using var stream = OfficePathIdentity.OpenRegularFileForRead(candidate, Path.GetDirectoryName(candidate)!, 32);
            if (stream.Length != 32) return false;
            byte[] owner = new byte[33];
            int count = OfficeRuntimeCompatibility.ReadAtLeast(stream, owner, 32, throwOnEndOfStream: false);
            return count == 32 && Encoding.ASCII.GetString(owner, 0, count) == runId;
        });
    }

    internal static void Validate(OfficeBatchReport report) {
        OfficeBatchReportSerializer.Validate(report);
        if (report.RunId != null && !Guid.TryParseExact(report.RunId, "N", out _)) throw new InvalidDataException("The report attempt identity is invalid.");
        if (report.JobSha256 != null && (report.JobSha256.Length != 64 || report.JobSha256.Any(character => !OfficeRuntimeCompatibility.IsAsciiHexDigit(character))))
            throw new InvalidDataException("The saved-batch hash is invalid.");
        if (report.RunId != null && report.StartedUtc == default || report.StartedUtc.Offset != TimeSpan.Zero ||
            report.FinishedUtc is { } finished && (finished.Offset != TimeSpan.Zero || finished < report.StartedUtc))
            throw new InvalidDataException("Run report times must be ordered UTC instants.");
        bool unfinished = report.State is OfficeBatchRunState.Running or OfficeBatchRunState.Interrupted;
        if (!OfficeRuntimeCompatibility.IsDefined(report.State) || unfinished && (report.FinishedUtc != null || report.Succeeded) ||
            !unfinished && report.RunId != null && report.FinishedUtc == null)
            throw new InvalidDataException("The run report lifecycle fields are inconsistent.");
        if (report.FailureSummary?.Length > 4096 || report.RecordingError?.Length > 4096)
            throw new InvalidDataException("The run report error text is oversized.");
    }
}
