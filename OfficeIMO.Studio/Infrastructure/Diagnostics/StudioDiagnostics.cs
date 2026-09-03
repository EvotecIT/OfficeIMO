using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace OfficeIMO.Studio.Infrastructure.Diagnostics;

/// <summary>Writes bounded JSON-lines diagnostics with paths and common identifiers removed.</summary>
internal sealed partial class StudioDiagnostics : IStudioDiagnostics {
    private const long MaximumLogBytes = 2L * 1024L * 1024L;
    private const int MaximumArchivedLogs = 4;
    private readonly object _sync = new();
    private readonly string _logPath;

    internal StudioDiagnostics(string directoryPath) {
        DirectoryPath = Path.GetFullPath(directoryPath ?? throw new ArgumentNullException(nameof(directoryPath)));
        _logPath = Path.Combine(DirectoryPath, "studio.log.jsonl");
    }

    public string DirectoryPath { get; }

    public void Write(StudioDiagnosticLevel level, string area, string code, Exception? exception = null) {
        if (string.IsNullOrWhiteSpace(area) || string.IsNullOrWhiteSpace(code)) return;
        try {
            lock (_sync) {
                Directory.CreateDirectory(DirectoryPath);
                RotateIfNeeded();
                var entry = new DiagnosticEntry(
                    DateTimeOffset.UtcNow,
                    level.ToString(),
                    SanitizeToken(area),
                    SanitizeToken(code),
                    exception?.GetType().FullName,
                    exception?.HResult,
                    SanitizeStack(exception?.StackTrace));
                File.AppendAllText(_logPath, JsonSerializer.Serialize(entry) + Environment.NewLine);
            }
        } catch (IOException) {
            // Diagnostics must never replace the original application failure.
        } catch (UnauthorizedAccessException) {
            // A locked-down profile may deny LocalApplicationData writes.
        }
    }

    public StudioSupportSnapshot CreateSupportSnapshot() {
        Assembly assembly = typeof(StudioDiagnostics).Assembly;
        return new StudioSupportSnapshot(
            "OfficeIMO Studio",
            assembly.GetName().Version?.ToString(3) ?? "0.0.0",
            RuntimeInformation.OSDescription,
            RuntimeInformation.FrameworkDescription,
            RuntimeInformation.ProcessArchitecture.ToString(),
            System.Globalization.CultureInfo.CurrentUICulture.Name,
            "LocalApplicationData/OfficeIMO/Studio/Diagnostics",
            "Diagnostics exclude document contents, document names, document paths, and exception messages.");
    }

    private void RotateIfNeeded() {
        if (!File.Exists(_logPath) || new FileInfo(_logPath).Length < MaximumLogBytes) return;
        for (int index = MaximumArchivedLogs; index >= 1; index--) {
            string current = _logPath + "." + index;
            if (!File.Exists(current)) continue;
            if (index == MaximumArchivedLogs) File.Delete(current);
            else File.Move(current, _logPath + "." + (index + 1), overwrite: true);
        }
        File.Move(_logPath, _logPath + ".1", overwrite: true);
    }

    private static string SanitizeToken(string value) => TokenSanitizer().Replace(value.Trim(), "_");

    private static string? SanitizeStack(string? stack) {
        if (string.IsNullOrWhiteSpace(stack)) return null;
        string sanitized = WindowsPathSanitizer().Replace(stack, "<path>");
        sanitized = UnixPathSanitizer().Replace(sanitized, "<path>");
        return sanitized.Length <= 12_000 ? sanitized : sanitized[..12_000];
    }

    [GeneratedRegex("[^A-Za-z0-9_.-]+", RegexOptions.CultureInvariant)]
    private static partial Regex TokenSanitizer();

    [GeneratedRegex(@"(?i)(?:[a-z]:\\|\\\\)[^\r\n:]+", RegexOptions.CultureInvariant)]
    private static partial Regex WindowsPathSanitizer();

    [GeneratedRegex(@"(?<![A-Za-z0-9])/(?:[^\s:]+/)+[^\s:]+", RegexOptions.CultureInvariant)]
    private static partial Regex UnixPathSanitizer();

    private sealed record DiagnosticEntry(
        DateTimeOffset TimestampUtc,
        string Level,
        string Area,
        string Code,
        string? ExceptionType,
        int? HResult,
        string? Stack);
}
