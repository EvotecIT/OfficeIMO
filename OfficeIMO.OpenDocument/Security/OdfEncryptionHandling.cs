namespace OfficeIMO.OpenDocument;

/// <summary>Controls save behavior for a package that was encrypted when loaded.</summary>
public enum OdfEncryptionHandling {
    /// <summary>Require encryption settings so an encrypted source cannot be written as plaintext accidentally.</summary>
    Preserve,
    /// <summary>Explicitly allow the decrypted package to be written without encryption.</summary>
    Remove
}
