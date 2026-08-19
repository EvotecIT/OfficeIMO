namespace OfficeIMO.OpenDocument;

/// <summary>Classifies why an encrypted OpenDocument package could not be processed.</summary>
public enum OdfEncryptionFailureReason {
    /// <summary>No password was supplied.</summary>
    PasswordRequired,
    /// <summary>The supplied password did not match the package checksum.</summary>
    IncorrectPassword,
    /// <summary>The package uses an encryption profile that this version does not implement.</summary>
    UnsupportedProfile,
    /// <summary>The encryption metadata or encrypted payload is malformed.</summary>
    InvalidEncryptedPackage,
    /// <summary>Decrypted content would exceed the configured resource budget.</summary>
    ResourceLimitExceeded,
    /// <summary>Saving would remove source encryption without explicit authorization.</summary>
    PreservationRequired
}
