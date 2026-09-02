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
        cancellationToken.ThrowIfCancellationRequested();
        switch (policy) {
            case OfficeWorkflowConflictPolicy.Fail:
                cancellationToken.ThrowIfCancellationRequested();
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
        cancellationToken.ThrowIfCancellationRequested();
        RecoverInterruptedDirectoryReplacement(requestedDirectory, diagnostics);

        if (File.Exists(requestedDirectory)) {
            throw new IOException("A file already occupies the requested output directory path.");
        }
        if (!Directory.Exists(requestedDirectory)) {
            cancellationToken.ThrowIfCancellationRequested();
            Directory.Move(stagingDirectory, requestedDirectory);
            return requestedDirectory;
        }

        string transactionId = Guid.NewGuid().ToString("N");
        string recoveryDirectory = requestedDirectory + ".officeimo-recovery-" + transactionId;
        string ownershipMarker = CreateDirectoryPublicationOwnershipMarker(requestedDirectory, transactionId);
        try {
            cancellationToken.ThrowIfCancellationRequested();
            Directory.Move(requestedDirectory, recoveryDirectory);
        } catch (Exception exception) when (exception is not OutOfMemoryException and not StackOverflowException) {
            TryDeleteFile(ownershipMarker);
            throw;
        }
        try {
            Directory.Move(stagingDirectory, requestedDirectory);
        } catch (Exception publicationException) when (publicationException is not OutOfMemoryException and not StackOverflowException) {
            try {
                Directory.Move(recoveryDirectory, requestedDirectory);
                TryDeleteFile(ownershipMarker);
            } catch (Exception rollbackException) when (rollbackException is not OutOfMemoryException and not StackOverflowException) {
                throw new DirectoryPublicationRecoveryException(
                    "The new output could not be published and the previous output could not be restored automatically.",
                    requestedDirectory,
                    [recoveryDirectory],
                    new AggregateException(publicationException, rollbackException));
            }
            throw;
        }

        string retiredDirectory = requestedDirectory + ".officeimo-retired-" + transactionId;
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
        Exception? retirementCleanupFailure = TryDeleteDirectory(retiredDirectory);
        if (retirementCleanupFailure is not null) {
            diagnostics.Add(CreateRetainedOutputDiagnostic(requestedDirectory, retiredDirectory, retirementCleanupFailure));
        } else {
            TryDeleteFile(ownershipMarker);
        }
        return requestedDirectory;
    }

    private static void RecoverInterruptedDirectoryReplacement(
        string requestedDirectory,
        ICollection<OfficeWorkflowDiagnostic> diagnostics) {
        string parent = Path.GetDirectoryName(requestedDirectory)!;
        string destinationName = Path.GetFileName(requestedDirectory);
        string recoveryPrefix = destinationName + ".officeimo-recovery-";
        string retiredPrefix = destinationName + ".officeimo-retired-";
        StringComparison pathComparison = OfficeWorkflowPathIdentity.GetComparison(parent);
        StringComparer pathComparer = OfficeWorkflowPathIdentity.GetComparer(parent);
        foreach (string retiredDirectory in Directory
            .EnumerateDirectories(parent, "*", SearchOption.TopDirectoryOnly)
            .Where(path => IsOwnedPublicationDirectory(requestedDirectory, path, retiredPrefix, pathComparison))
            .OrderBy(static path => path, pathComparer)
            .ThenBy(static path => path, StringComparer.Ordinal)) {
            Exception? cleanupFailure = TryDeleteDirectory(retiredDirectory);
            if (cleanupFailure is not null) {
                diagnostics.Add(CreateRetainedOutputDiagnostic(requestedDirectory, retiredDirectory, cleanupFailure));
            } else {
                TryDeleteFile(GetDirectoryPublicationOwnershipMarker(requestedDirectory, GetPublicationTransactionId(retiredDirectory)));
            }
        }

        string[] recoveryDirectories = Directory
            .EnumerateDirectories(parent, "*", SearchOption.TopDirectoryOnly)
            .Where(path => IsOwnedPublicationDirectory(requestedDirectory, path, recoveryPrefix, pathComparison))
            .OrderBy(static path => path, pathComparer)
            .ThenBy(static path => path, StringComparer.Ordinal)
            .ToArray();
        if (recoveryDirectories.Length == 0) return;

        if (!Directory.Exists(requestedDirectory) && !File.Exists(requestedDirectory) && recoveryDirectories.Length == 1) {
            string ownershipMarker = GetDirectoryPublicationOwnershipMarker(
                requestedDirectory,
                GetPublicationTransactionId(recoveryDirectories[0]));
            Directory.Move(recoveryDirectories[0], requestedDirectory);
            TryDeleteFile(ownershipMarker);
            return;
        }

        throw new DirectoryPublicationRecoveryException(
            "A prior interrupted output replacement requires recovery before this destination can be replaced.",
            requestedDirectory,
            recoveryDirectories);
    }

    private static bool IsOwnedPublicationDirectory(
        string requestedDirectory,
        string path,
        string prefix,
        StringComparison comparison) {
        string name = Path.GetFileName(path);
        if (!name.StartsWith(prefix, comparison)) return false;
        string suffix = name[prefix.Length..];
        if (suffix.Length != 32 || !Guid.TryParseExact(suffix, "N", out _)) return false;
        string markerPath = GetDirectoryPublicationOwnershipMarker(requestedDirectory, suffix);
        try {
            if (!File.Exists(markerPath) || new FileInfo(markerPath).Length > 4096L) return false;
            return string.Equals(
                File.ReadAllText(markerPath, Encoding.UTF8),
                GetDirectoryPublicationOwnershipContents(requestedDirectory, suffix),
                comparison);
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            return false;
        }
    }

    internal static string CreateDirectoryPublicationOwnershipMarker(string requestedDirectory, string transactionId) {
        if (!Guid.TryParseExact(transactionId, "N", out _)) {
            throw new ArgumentException("Directory publication transaction id must be a 32-character GUID.", nameof(transactionId));
        }
        string markerPath = GetDirectoryPublicationOwnershipMarker(requestedDirectory, transactionId);
        using var stream = new FileStream(markerPath, FileMode.CreateNew, FileAccess.Write, FileShare.None);
        using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        writer.Write(GetDirectoryPublicationOwnershipContents(requestedDirectory, transactionId));
        return markerPath;
    }

    private static string GetDirectoryPublicationOwnershipMarker(string requestedDirectory, string transactionId) =>
        requestedDirectory + ".officeimo-publication-" + transactionId + ".owner";

    private static string GetDirectoryPublicationOwnershipContents(string requestedDirectory, string transactionId) =>
        "OfficeIMO.Workflows.DirectoryPublication.v1\n" + Path.GetFullPath(requestedDirectory) + "\n" + transactionId;

    private static string GetPublicationTransactionId(string publicationDirectory) =>
        Path.GetFileName(publicationDirectory)[^32..];

    private static void TryDeleteFile(string path) {
        try {
            if (File.Exists(path)) File.Delete(path);
        } catch (IOException) {
            // A leftover marker is inert without its matching transaction directory.
        } catch (UnauthorizedAccessException) {
            // A leftover marker is inert without its matching transaction directory.
        }
    }

    private static OfficeWorkflowDiagnostic CreateRetainedOutputDiagnostic(
        string requestedDirectory,
        string retainedDirectory,
        Exception exception) => new(
            "PreviousOutputRetained",
            "The new output is available, but a retired previous output remains and will be retried during the next replacement.",
            OfficeWorkflowDiagnosticSeverity.Warning,
            "publish",
            new Dictionary<string, string>(StringComparer.Ordinal) {
                ["destination"] = requestedDirectory,
                ["recoveryPaths"] = retainedDirectory,
                ["exceptionType"] = exception.GetType().Name
            });

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

    private static Exception? TryDeleteDirectory(string path) {
        try {
            if (Directory.Exists(path)) Directory.Delete(path, recursive: true);
            return null;
        } catch (IOException exception) {
            return exception;
        } catch (UnauthorizedAccessException exception) {
            return exception;
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
