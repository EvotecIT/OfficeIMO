namespace OfficeIMO.Pdf;

/// <summary>Controls how authenticated PDF Standard-security permission restrictions are handled.</summary>
public enum PdfPermissionPolicy {
    /// <summary>Enforces the permission bits associated with user-password authorization.</summary>
    Enforce,

    /// <summary>
    /// Ignores authenticated user-password permission restrictions after the PDF has been successfully decrypted.
    /// This does not discover, bypass, or crack an unknown document-open password.
    /// </summary>
    IgnoreRestrictions
}
