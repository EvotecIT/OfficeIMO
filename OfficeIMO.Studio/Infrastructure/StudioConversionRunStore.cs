using System.Text.Json;
using OfficeIMO.Workflows;

namespace OfficeIMO.Studio.Infrastructure;

/// <summary>Retains bounded, observational saved-job reports under Studio's local history root.</summary>
internal sealed class StudioConversionRunStore {
    internal const int MaximumEntries = 50;
    private readonly string _directory;
    private readonly Func<bool> _enabled;

    internal StudioConversionRunStore(string directory, Func<bool> enabled) {
        _directory = Path.GetFullPath(directory);
        _enabled = enabled;
    }

    /// <summary>Consulted again by the shared writer so opting out during a run stops retention.</summary>
    internal bool CanRetainReport => _enabled();

    /// <summary>Allocates only a filename; the shared recorded runner owns report publication and leases.</summary>
    internal string? CreateReportPath() {
        if (!CanRetainReport) return null;
        EnsureDirectory(create: true);
        return Path.Combine(_directory, Guid.NewGuid().ToString("N") + ".json");
    }

    /// <summary>Reads the most recently updated reports independently, so one malformed file cannot hide others.</summary>
    internal StudioConversionRunHistory ReadLatest() {
        if (!CanRetainReport) return new([], null);
        var entries = new List<StudioConversionRunEntry>();
        string? error = null;
        try {
            foreach (string path in ReportFiles().Take(MaximumEntries)) {
                try {
                    EnsureRegularFile(path);
                    entries.Add(new(path, OfficeBatchReport.Load(path)));
                } catch (Exception failure) when (IsStorageError(failure)) {
                    error ??= BoundedError(path, failure);
                }
            }
        } catch (Exception failure) when (IsStorageError(failure)) {
            error ??= BoundedError(_directory, failure);
        }
        return new(entries.OrderByDescending(entry => entry.Report.StartedUtc).ToArray(), error);
    }

    /// <summary>Refreshes an already admitted card without restoring new history while recording is disabled.</summary>
    internal OfficeBatchReport ReadReport(string reportPath) {
        string path = RequireOwnedPath(reportPath);
        EnsureDirectory(create: false);
        EnsureRegularFile(path);
        return OfficeBatchReport.Load(path);
    }

    /// <summary>Removes only this store's inactive report. Referenced job, output and checkpoint paths are never followed.</summary>
    internal bool RemoveFinished(string reportPath) {
        try {
            string path = RequireOwnedPath(reportPath);
            if (!File.Exists(path)) return true;
            EnsureDirectory(create: false);
            EnsureRegularFile(path);
            return OfficeBatchReport.TryDeleteFinished(path);
        } catch (Exception failure) when (IsStorageError(failure)) { return false; }
    }

    /// <summary>Clears eligible reports while preserving live writer leases and reporting partial removal.</summary>
    internal bool ClearFinished() => RemoveReports(keepLatest: 0);

    /// <summary>Bounds retained completed reports without interrupting any live attempt.</summary>
    internal bool TrimFinished() => RemoveReports(MaximumEntries);

    private bool RemoveReports(int keepLatest) {
        bool succeeded = true;
        try {
            foreach (string path in ReportFiles().Skip(keepLatest)) {
                try {
                    if (OfficeBatchReport.IsActive(path)) continue;
                    if (!RemoveFinished(path) && !OfficeBatchReport.IsActive(path)) succeeded = false;
                } catch (Exception failure) when (IsStorageError(failure)) { succeeded = false; }
            }
        } catch (Exception failure) when (IsStorageError(failure)) { succeeded = false; }
        return succeeded;
    }

    private IEnumerable<string> ReportFiles() {
        EnsureDirectory(create: false);
        if (!Directory.Exists(_directory)) return [];
        // A known filename is necessary but not sufficient for removal: the shared owner also
        // validates the report schema and acquires its lease before deleting it.
        return Directory.EnumerateFiles(_directory, "*.json", SearchOption.TopDirectoryOnly)
            .Where(path => IsReportName(Path.GetFileName(path)))
            .OrderByDescending(path => File.GetLastWriteTimeUtc(path));
    }

    private string RequireOwnedPath(string path) {
        string full = Path.GetFullPath(path);
        if (!string.Equals(Path.GetDirectoryName(full), _directory, OperatingSystem.IsWindows() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal) ||
            !IsReportName(Path.GetFileName(full))) throw new IOException("The conversion report is outside Studio's history store.");
        return full;
    }

    private static bool IsReportName(string name) => name.EndsWith(".json", StringComparison.OrdinalIgnoreCase) &&
        Guid.TryParseExact(Path.GetFileNameWithoutExtension(name), "N", out _);

    private void EnsureDirectory(bool create) {
        if (Directory.Exists(_directory)) {
            if ((File.GetAttributes(_directory) & FileAttributes.ReparsePoint) != 0)
                throw new IOException("The conversion history directory cannot be a link.");
        } else if (File.Exists(_directory)) throw new IOException("The conversion history location is not a directory.");
        else if (create) Directory.CreateDirectory(_directory);
    }

    private static void EnsureRegularFile(string path) {
        if ((File.GetAttributes(path) & (FileAttributes.ReparsePoint | FileAttributes.Directory)) != 0)
            throw new IOException("Conversion history reports cannot be links or directories.");
    }

    private static bool IsStorageError(Exception error) => error is IOException or InvalidDataException or UnauthorizedAccessException or
        JsonException or ArgumentException or NotSupportedException or InvalidOperationException;

    private static string BoundedError(string path, Exception error) {
        string message = Path.GetFileName(path) + ": " + error.Message;
        return message.Length > 1024 ? message[..1024] : message;
    }
}

/// <summary>A report with its app-owned storage path, independent of the locations recorded in its content.</summary>
internal sealed record StudioConversionRunEntry(string ReportPath, OfficeBatchReport Report);

/// <summary>The bounded reports and first storage error available to the Jobs surface.</summary>
internal sealed record StudioConversionRunHistory(IReadOnlyList<StudioConversionRunEntry> Entries, string? Error);
