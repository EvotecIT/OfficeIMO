using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Workflows;

public sealed partial class OfficeWorkflowRunner {
    private static async Task<string> PublishDirectoryAsync(
        string stagingDirectory,
        string requestedDirectory,
        OfficeWorkflowConflictPolicy policy,
        ICollection<OfficeWorkflowDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        switch (policy) {
            case OfficeWorkflowConflictPolicy.Fail:
                Directory.Move(stagingDirectory, requestedDirectory);
                return requestedDirectory;
            case OfficeWorkflowConflictPolicy.Rename:
                for (int suffix = 0; suffix < 10_000; suffix++) {
                    cancellationToken.ThrowIfCancellationRequested();
                    string candidate = suffix == 0 ? requestedDirectory : AddDirectorySuffix(requestedDirectory, suffix);
                    try {
                        Directory.Move(stagingDirectory, candidate);
                        return candidate;
                    } catch (IOException) when (Directory.Exists(candidate) || File.Exists(candidate)) {
                        // A concurrent request owns this path. Try the next deterministic suffix.
                    }
                }
                throw new IOException("No available numbered output directory could be reserved.");
            case OfficeWorkflowConflictPolicy.Replace:
                return await ReplaceDirectoryAsync(stagingDirectory, requestedDirectory, diagnostics, cancellationToken).ConfigureAwait(false);
            default:
                throw new ArgumentOutOfRangeException(nameof(policy));
        }
    }

    private static async Task<string> ReplaceDirectoryAsync(
        string stagingDirectory,
        string requestedDirectory,
        ICollection<OfficeWorkflowDiagnostic> diagnostics,
        CancellationToken cancellationToken) {
        await using FileStream publicationLock = await AcquireDirectoryPublicationLockAsync(requestedDirectory, cancellationToken)
            .ConfigureAwait(false);
        RecoverInterruptedDirectoryReplacement(requestedDirectory);

        if (File.Exists(requestedDirectory)) {
            throw new IOException("A file already occupies the requested output directory path.");
        }
        if (!Directory.Exists(requestedDirectory)) {
            Directory.Move(stagingDirectory, requestedDirectory);
            return requestedDirectory;
        }

        string recoveryDirectory = requestedDirectory + ".officeimo-recovery-" + Guid.NewGuid().ToString("N");
        Directory.Move(requestedDirectory, recoveryDirectory);
        try {
            Directory.Move(stagingDirectory, requestedDirectory);
        } catch (Exception publicationException) when (publicationException is not OutOfMemoryException and not StackOverflowException) {
            try {
                Directory.Move(recoveryDirectory, requestedDirectory);
            } catch (Exception rollbackException) when (rollbackException is not OutOfMemoryException and not StackOverflowException) {
                throw new DirectoryPublicationRecoveryException(
                    "The new output could not be published and the previous output could not be restored automatically.",
                    requestedDirectory,
                    [recoveryDirectory],
                    new AggregateException(publicationException, rollbackException));
            }
            throw;
        }

        string retiredDirectory = requestedDirectory + ".officeimo-retired-" + Guid.NewGuid().ToString("N");
        try {
            Directory.Move(recoveryDirectory, retiredDirectory);
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            diagnostics.Add(new OfficeWorkflowDiagnostic(
                "PreviousOutputRetained",
                "The new output was published, but the previous output could not be retired automatically.",
                OfficeWorkflowDiagnosticSeverity.Warning,
                "publish",
                new Dictionary<string, string>(StringComparer.Ordinal) {
                    ["destination"] = requestedDirectory,
                    ["recoveryPaths"] = recoveryDirectory,
                    ["exceptionType"] = exception.GetType().Name
                }));
            return requestedDirectory;
        }
        TryDeleteDirectory(retiredDirectory);
        return requestedDirectory;
    }

    private static void RecoverInterruptedDirectoryReplacement(string requestedDirectory) {
        string parent = Path.GetDirectoryName(requestedDirectory)!;
        string prefix = Path.GetFileName(requestedDirectory) + ".officeimo-recovery-";
        StringComparison pathComparison = OfficeWorkflowPathIdentity.GetComparison(parent);
        StringComparer pathComparer = OfficeWorkflowPathIdentity.GetComparer(parent);
        string[] recoveryDirectories = Directory
            .EnumerateDirectories(parent, "*", SearchOption.TopDirectoryOnly)
            .Where(path => Path.GetFileName(path).StartsWith(prefix, pathComparison))
            .OrderBy(static path => path, pathComparer)
            .ThenBy(static path => path, StringComparer.Ordinal)
            .ToArray();
        if (recoveryDirectories.Length == 0) return;

        if (!Directory.Exists(requestedDirectory) && !File.Exists(requestedDirectory) && recoveryDirectories.Length == 1) {
            Directory.Move(recoveryDirectories[0], requestedDirectory);
            return;
        }

        throw new DirectoryPublicationRecoveryException(
            "A prior interrupted output replacement requires recovery before this destination can be replaced.",
            requestedDirectory,
            recoveryDirectories);
    }

    private static async Task<FileStream> AcquireDirectoryPublicationLockAsync(
        string requestedDirectory,
        CancellationToken cancellationToken) {
        string lockRoot = Path.Combine(Path.GetTempPath(), "OfficeIMO", "directory-publication-locks");
        Directory.CreateDirectory(lockRoot);
        string identity = OfficeWorkflowPathIdentity.Normalize(requestedDirectory);
        string lockName = System.Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(identity))) + ".lock";
        string lockPath = Path.Combine(lockRoot, lockName);
        DateTime deadline = DateTime.UtcNow.AddSeconds(30D);

        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            try {
                return new FileStream(lockPath, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.None, 1, FileOptions.Asynchronous);
            } catch (IOException) when (DateTime.UtcNow < deadline) {
                await Task.Delay(25, cancellationToken).ConfigureAwait(false);
            }
        }
    }

    private static IReadOnlyDictionary<string, string> CreateFailureDetails(Exception exception) {
        var details = new Dictionary<string, string>(StringComparer.Ordinal) {
            ["exceptionType"] = exception.GetType().Name
        };
        if (exception is DirectoryPublicationRecoveryException recoveryException) {
            details["destination"] = recoveryException.Destination;
            details["recoveryPaths"] = string.Join(Path.PathSeparator, recoveryException.RecoveryPaths);
        }
        return details;
    }

    private static string AddDirectorySuffix(string path, int suffix) =>
        path.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) +
        " (" + suffix.ToString(System.Globalization.CultureInfo.InvariantCulture) + ")";

    private static void TryDeleteDirectory(string path) {
        try {
            if (Directory.Exists(path)) Directory.Delete(path, recursive: true);
        } catch (IOException) {
            // Best-effort cleanup of a task-owned staging, extraction, or retired directory.
        } catch (UnauthorizedAccessException) {
            // Best-effort cleanup of a task-owned staging, extraction, or retired directory.
        }
    }

    private sealed class DirectoryPublicationRecoveryException : IOException {
        internal DirectoryPublicationRecoveryException(
            string message,
            string destination,
            IReadOnlyList<string> recoveryPaths,
            Exception? innerException = null)
            : base(message + " Destination: '" + destination + "'. Recovery path(s): " + string.Join(", ", recoveryPaths.Select(static path => "'" + path + "'")) + ".", innerException) {
            Destination = destination;
            RecoveryPaths = recoveryPaths.ToArray();
        }

        internal string Destination { get; }
        internal IReadOnlyList<string> RecoveryPaths { get; }
    }
}
