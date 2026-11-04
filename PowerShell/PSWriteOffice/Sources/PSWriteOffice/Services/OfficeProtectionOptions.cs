using OfficeIMO.Pdf;
using OfficeIMO.Security;

namespace PSWriteOffice.Services;

/// <summary>Shell options for the Excel protection vocabulary in the cross-format command.</summary>
public sealed class OfficeProtectionExcelOptions {
    /// <summary>Precomputed native workbook protection hash.</summary>
    public string? LegacyPasswordHash { get; set; }
    /// <summary>Precomputed native write-reservation password hash.</summary>
    public string? ModifyPasswordHash { get; set; }
    /// <summary>Name displayed by Excel for a write reservation.</summary>
    public string? ReservedBy { get; set; }
}

/// <summary>
/// Password-free PDF option overrides for hashtable binding. Native options are created with the command's
/// open password or through the native permissions-only factory with its owner password.
/// </summary>
public sealed class OfficeProtectionPdfOptions {
    /// <summary>Native encryption algorithm.</summary>
    public PdfStandardEncryptionAlgorithm Algorithm { get; set; } = PdfStandardEncryptionAlgorithm.Aes256;
    /// <summary>Whether metadata streams are encrypted.</summary>
    public bool EncryptMetadata { get; set; } = true;
    /// <summary>Native PDF permissions mask.</summary>
    public int? Permissions { get; set; }
    /// <summary>Optional typed native permissions override.</summary>
    public PdfStandardPermissions? AllowedPermissions { get; set; }
    /// <summary>Optional already configured native AES provider.</summary>
    public IOfficeAesCryptographyProvider? AesCryptographyProvider { get; set; }

    internal void Apply(PdfStandardEncryptionOptions target) {
        target.Algorithm = Algorithm;
        target.EncryptMetadata = EncryptMetadata;
        if (Permissions.HasValue) target.Permissions = Permissions.Value;
        if (AllowedPermissions.HasValue) target.AllowedPermissions = AllowedPermissions.Value;
        target.AesCryptographyProvider = AesCryptographyProvider;
    }
}
