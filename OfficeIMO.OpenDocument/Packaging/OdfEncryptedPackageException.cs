namespace OfficeIMO.OpenDocument;

/// <summary>Thrown when an encrypted ODF package is opened before native encryption support is enabled.</summary>
public sealed class OdfEncryptedPackageException : NotSupportedException {
    /// <summary>Creates the exception.</summary>
    public OdfEncryptedPackageException(string message) : base(message) {
    }
}
