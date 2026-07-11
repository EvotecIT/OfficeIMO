namespace OfficeIMO.OpenDocument;

/// <summary>Controls OpenDocument serialization.</summary>
public sealed class OdfSaveOptions {
    /// <summary>Compatibility profile used for versioned output.</summary>
    public OdfCompatibilityProfile CompatibilityProfile { get; set; } = OdfCompatibilityProfile.Odf14;

    /// <summary>Controls changed signed-document behavior.</summary>
    public OdfSignatureHandling SignatureHandling { get; set; } = OdfSignatureHandling.RejectInvalidation;

    /// <summary>Use stable timestamps and ordinal entry ordering after preserved source entries.</summary>
    public bool Deterministic { get; set; } = true;
}
