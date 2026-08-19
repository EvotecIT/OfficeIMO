namespace OfficeIMO.OpenDocument;

/// <summary>Controls OpenDocument serialization.</summary>
public sealed class OdfSaveOptions {
    /// <summary>Compatibility profile used for versioned output.</summary>
    public OdfCompatibilityProfile CompatibilityProfile { get; set; } = OdfCompatibilityProfile.Odf14;

    /// <summary>Controls changed signed-document behavior.</summary>
    public OdfSignatureHandling SignatureHandling { get; set; } = OdfSignatureHandling.RejectInvalidation;

    /// <summary>Encryption settings for the output package.</summary>
    /// <remarks>
    /// Supplying these settings writes the interoperable ODF AES-256-CBC/PBKDF2 profile. An encrypted source
    /// requires either these settings or an explicit <see cref="EncryptionHandling"/> value of
    /// <see cref="OdfEncryptionHandling.Remove"/> so protection cannot be removed accidentally.
    /// </remarks>
    public OdfEncryptionOptions? Encryption { get; set; }

    /// <summary>Controls whether encryption may be removed from a package that was encrypted when loaded.</summary>
    public OdfEncryptionHandling EncryptionHandling { get; set; } = OdfEncryptionHandling.Preserve;

    /// <summary>Use stable timestamps and ordinal entry ordering after preserved source entries.</summary>
    public bool Deterministic { get; set; } = true;
}
