namespace OfficeIMO.Pdf;

/// <summary>Raised when an authenticated PDF permission restriction blocks an extraction or mutation.</summary>
public sealed class PdfPermissionDeniedException : InvalidOperationException {
    internal PdfPermissionDeniedException(PdfStandardPermissions permission, PdfPasswordAuthenticationRole authenticationRole, string message)
        : base(message) {
        Permission = permission;
        AuthenticationRole = authenticationRole;
    }

    /// <summary>Permission required by the blocked operation.</summary>
    public PdfStandardPermissions Permission { get; }

    /// <summary>Password role established while opening the PDF.</summary>
    public PdfPasswordAuthenticationRole AuthenticationRole { get; }
}
