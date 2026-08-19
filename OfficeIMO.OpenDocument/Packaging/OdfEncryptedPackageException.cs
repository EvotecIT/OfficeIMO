namespace OfficeIMO.OpenDocument;

/// <summary>Thrown when an encrypted ODF package cannot be processed safely.</summary>
public sealed class OdfEncryptedPackageException : NotSupportedException {
    /// <summary>Creates the exception.</summary>
    public OdfEncryptedPackageException(string message) : this(message, OdfEncryptionFailureReason.UnsupportedProfile, null, null) {
    }

    /// <summary>Creates a classified encrypted-package exception.</summary>
    public OdfEncryptedPackageException(string message, OdfEncryptionFailureReason reason, string? entryPath = null,
        Exception? innerException = null) : base(message, innerException) {
        Reason = reason;
        EntryPath = entryPath;
    }

    /// <summary>Reason the encrypted package could not be processed.</summary>
    public OdfEncryptionFailureReason Reason { get; }

    /// <summary>Encrypted package entry associated with the failure, when known.</summary>
    public string? EntryPath { get; }
}
