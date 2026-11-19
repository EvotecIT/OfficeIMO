using System.Security.Cryptography;
using System.Text.Json;
using System.Xml;

namespace OfficeIMO.Reader;

/// <summary>
/// Stable codes for inputs that could not be read. Batch outcomes and folder documents report one of these values so
/// callers can branch on the cause without matching exception types or messages.
/// </summary>
/// <remarks>
/// Codes are lower-case, hyphenated and never renamed. New codes may be added; treat an unrecognized code like
/// <see cref="ReadFailed"/>.
/// </remarks>
public static class ReaderErrorCodes {
    /// <summary>The input is encrypted and no password was supplied, or the empty password was not accepted.</summary>
    public const string PasswordRequired = "password-required";

    /// <summary>The input is encrypted and the supplied password did not open it.</summary>
    public const string InvalidPassword = "invalid-password";

    /// <summary>The input uses an encryption method, or a certificate-based protection, that the reader cannot open.</summary>
    public const string UnsupportedEncryption = "unsupported-encryption";

    /// <summary>The input opened, but its own permissions do not allow content extraction.</summary>
    public const string PermissionDenied = "permission-denied";

    /// <summary>No configured handler reads the input, or the input uses a format feature the handler does not support.</summary>
    public const string UnsupportedFormat = "unsupported-format";

    /// <summary>The input is damaged, truncated or not valid for its detected format.</summary>
    public const string InvalidContent = "invalid-content";

    /// <summary>A configured time limit, such as <see cref="ReaderOptions.LayoutTimeout"/>, was exceeded.</summary>
    public const string Timeout = "timeout";

    /// <summary>The file or directory does not exist.</summary>
    public const string NotFound = "not-found";

    /// <summary>The file system denied access to the input.</summary>
    public const string AccessDenied = "access-denied";

    /// <summary>The input could not be read because of an I/O failure.</summary>
    public const string IOError = "io-error";

    /// <summary>The read failed for a reason without a more specific code.</summary>
    public const string ReadFailed = "read-failed";
}

internal static partial class DocumentReaderEngine {
    private const int MaximumClassifiedExceptionDepth = 8;

    /// <summary>
    /// Maps a read failure to a stable <see cref="ReaderErrorCodes"/> value. Handler classifiers see each exception in the
    /// chain first, so format-specific password and encryption failures keep their meaning; general exception types
    /// supply the fallback.
    /// </summary>
    internal static string ClassifyReadError(Exception exception) {
        ArgumentNullException.ThrowIfNull(exception);
        IReadOnlyList<Exception> chain = GetExceptionChain(exception);
        IReadOnlyList<ReaderHandlerDescriptor> handlers = GetActiveHandlerRegistry().Handlers;
        foreach (Exception current in chain) {
            foreach (ReaderHandlerDescriptor handler in handlers) {
                string? code = handler.ClassifyReadError?.Invoke(current);
                if (!string.IsNullOrWhiteSpace(code)) return code!;
            }
        }
        foreach (Exception current in chain) {
            string? code = ClassifyCommonReadError(current);
            if (code != null) return code;
        }
        return ReaderErrorCodes.ReadFailed;
    }

    /// <summary>
    /// <see cref="Exception.Data"/> key the Office engines set on an exception thrown because the supplied password did
    /// not verify. It keeps a wrong password apart from damaged content that fails an integrity check after the password
    /// was accepted, which surfaces as the same exception type.
    /// </summary>
    internal const string OfficePasswordRejectedDataKey = "OfficeIMO.PasswordRejected";

    /// <summary>
    /// Classifies a failed file read. A general load or decryption failure is refined when the cause is the password: an
    /// Office engine marks a rejected password, and an encrypted Office or legacy Office file read without
    /// <see cref="ReaderOptions.OpenPassword"/> is found by bounded detection of the same file, which follows the read's
    /// detection limits and container-inspection setting. Other codes, such as an unsupported projection, are kept.
    /// </summary>
    internal static string ClassifyReadError(Exception exception, string? path, ReaderOptions? options) {
        string code = ClassifyReadError(exception);
        if (code is not (ReaderErrorCodes.InvalidContent or ReaderErrorCodes.ReadFailed)) return code;
        bool hasPassword = !string.IsNullOrEmpty(options?.OpenPassword);
        if (GetExceptionChain(exception).Any(static current => current.Data.Contains(OfficePasswordRejectedDataKey))) {
            return hasPassword ? ReaderErrorCodes.InvalidPassword : ReaderErrorCodes.PasswordRequired;
        }
        if (path == null || hasPassword) return code;

        try {
            if (!File.Exists(path)) return code;
            ReaderDetectionResult detection = Detect(path, new ReaderDetectionOptions {
                Mode = ReaderDetectionMode.PreferContent,
                MaxProbeBytes = options?.DetectionMaxProbeBytes ?? ReaderOptions.DefaultDetectionMaxProbeBytes,
                MaxContainerEntries = options?.DetectionMaxContainerEntries ?? ReaderOptions.DefaultDetectionMaxContainerEntries,
                InspectContainers = options?.DetectionInspectContainers ?? true
            });
            return detection.IsEncrypted && detection.Evidence.Any(IsOfficeEncryptionEvidence)
                ? ReaderErrorCodes.PasswordRequired
                : code;
        } catch (Exception detectionFailure) when (detectionFailure is IOException or UnauthorizedAccessException
                                                       or InvalidDataException or ReaderResourceLimitException) {
            return code;
        }
    }

    private static bool IsOfficeEncryptionEvidence(string evidence) =>
        evidence.StartsWith("encryption:office-", StringComparison.Ordinal) ||
        evidence.StartsWith("encryption:legacy-", StringComparison.Ordinal);

    private static string? ClassifyCommonReadError(Exception exception) => exception switch {
        TimeoutException => ReaderErrorCodes.Timeout,
        FileNotFoundException or DirectoryNotFoundException => ReaderErrorCodes.NotFound,
        UnauthorizedAccessException => ReaderErrorCodes.AccessDenied,
        InvalidDataException or EndOfStreamException or FormatException or XmlException or JsonException
            or CryptographicException => ReaderErrorCodes.InvalidContent,
        NotSupportedException => ReaderErrorCodes.UnsupportedFormat,
        IOException => ReaderErrorCodes.IOError,
        _ => null
    };

    private static IReadOnlyList<Exception> GetExceptionChain(Exception exception) {
        var chain = new List<Exception>();
        Exception? current = exception;
        while (current != null && chain.Count < MaximumClassifiedExceptionDepth) {
            if (current is AggregateException aggregate && aggregate.InnerExceptions.Count == 1) {
                current = aggregate.InnerExceptions[0];
                continue;
            }
            chain.Add(current);
            current = current.InnerException;
        }
        return chain;
    }
}
