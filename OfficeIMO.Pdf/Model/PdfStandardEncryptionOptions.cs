using OfficeIMO.Security;

namespace OfficeIMO.Pdf;

/// <summary>
/// Options for generated PDF Standard password security.
/// </summary>
public sealed class PdfStandardEncryptionOptions {
    /// <summary>PDF permission bit mask allowing all standard operations.</summary>
    public const int AllowAllPermissions = -4;

    private string _userPassword = string.Empty;
    private string? _ownerPassword;
    private int _permissions = AllowAllPermissions;

    /// <summary>Creates Standard security options using the supplied document-open password.</summary>
    public PdfStandardEncryptionOptions(string userPassword) {
        UserPassword = userPassword;
    }

    private PdfStandardEncryptionOptions() { }

    /// <summary>Creates permissions-only security: readers can open without a password, while changing
    /// restrictions requires the supplied owner password.</summary>
    /// <param name="ownerPassword">Non-empty password authorizing permission changes.</param>
    /// <param name="allowedPermissions">Operations permitted to readers without the owner password.</param>
    /// <returns>Explicit permissions-only encryption options using AES-256 by default.</returns>
    public static PdfStandardEncryptionOptions ForPermissions(string ownerPassword, PdfStandardPermissions allowedPermissions) {
        Guard.NotNullOrWhiteSpace(ownerPassword, nameof(ownerPassword));
        if ((allowedPermissions & ~PdfStandardPermissions.All) != 0)
            throw new ArgumentOutOfRangeException(nameof(allowedPermissions));
        return new PdfStandardEncryptionOptions { OwnerPassword = ownerPassword, AllowedPermissions = allowedPermissions };
    }

    /// <summary>Password required to open the generated PDF, or empty for options created by
    /// <see cref="ForPermissions"/>. AES-256 accepts at most 4096 UTF-16 characters before normalization.</summary>
    public string UserPassword {
        get => _userPassword;
        set {
            Guard.NotNullOrWhiteSpace(value, nameof(UserPassword));
            _userPassword = value;
        }
    }

    /// <summary>Optional owner password. When omitted, the user password is reused. AES-256 accepts at most 4096 UTF-16 characters before normalization.</summary>
    public string? OwnerPassword {
        get => _ownerPassword;
        set {
            if (value == null && _userPassword.Length == 0)
                throw new ArgumentException("Permissions-only security requires an owner password.", nameof(OwnerPassword));
            if (value != null) {
                Guard.NotNullOrWhiteSpace(value, nameof(OwnerPassword));
            }

            _ownerPassword = value;
        }
    }

    /// <summary>Whether the owner password differs from the open password after the selected PDF
    /// algorithm's normalization and byte limit. False means an open-password holder also has owner access.</summary>
    public bool HasDistinctOwnerPassword => OwnerPassword != null &&
        !PdfStandardSecurityWriter.PasswordsAreEquivalent(UserPassword, OwnerPassword, Algorithm, default);

    /// <summary>Encryption algorithm. AES-256 revision 6 is the default; RC4 requires explicit legacy selection.</summary>
    public PdfStandardEncryptionAlgorithm Algorithm { get; set; } = PdfStandardEncryptionAlgorithm.Aes256;

    /// <summary>Whether document metadata streams are encrypted. Defaults to true.</summary>
    public bool EncryptMetadata { get; set; } = true;

    /// <summary>
    /// Optional synchronous AES provider used when the host does not expose platform AES-CBC, such as browser WebAssembly.
    /// Native platform cryptography remains the default when this property is null.
    /// </summary>
    public IOfficeAesCryptographyProvider? AesCryptographyProvider { get; set; }

    /// <summary>Raw PDF Standard security permission bit mask. Defaults to allowing all standard operations.</summary>
    public int Permissions {
        get => _permissions;
        set => _permissions = value;
    }

    /// <summary>Typed view over the Standard security permission mask.</summary>
    public PdfStandardPermissions AllowedPermissions {
        get => FromRawPermissions(_permissions);
        set => _permissions = ToRawPermissions(value);
    }

    /// <summary>Creates a deep copy of the encryption options.</summary>
    public PdfStandardEncryptionOptions Clone() {
        var copy = UserPassword.Length == 0
            ? ForPermissions(OwnerPassword!, AllowedPermissions)
            : new PdfStandardEncryptionOptions(UserPassword);
        copy.OwnerPassword = OwnerPassword;
        copy.Permissions = Permissions;
        copy.Algorithm = Algorithm;
        copy.EncryptMetadata = EncryptMetadata;
        copy.AesCryptographyProvider = AesCryptographyProvider;
        return copy;
    }

    internal static int ToRawPermissions(PdfStandardPermissions permissions) {
        const int requiredBits = unchecked((int)0xFFFFF0C0);
        int raw = requiredBits;
        if ((permissions & PdfStandardPermissions.Print) != 0) raw |= 1 << 2;
        if ((permissions & PdfStandardPermissions.ModifyContents) != 0) raw |= 1 << 3;
        if ((permissions & PdfStandardPermissions.CopyContents) != 0) raw |= 1 << 4;
        if ((permissions & PdfStandardPermissions.ModifyAnnotations) != 0) raw |= 1 << 5;
        if ((permissions & PdfStandardPermissions.FillForms) != 0) raw |= 1 << 8;
        if ((permissions & PdfStandardPermissions.Accessibility) != 0) raw |= 1 << 9;
        if ((permissions & PdfStandardPermissions.AssembleDocument) != 0) raw |= 1 << 10;
        if ((permissions & PdfStandardPermissions.HighQualityPrint) != 0) raw |= 1 << 11;
        return raw;
    }

    internal static PdfStandardPermissions FromRawPermissions(int raw) {
        PdfStandardPermissions permissions = PdfStandardPermissions.None;
        if ((raw & (1 << 2)) != 0) permissions |= PdfStandardPermissions.Print;
        if ((raw & (1 << 3)) != 0) permissions |= PdfStandardPermissions.ModifyContents;
        if ((raw & (1 << 4)) != 0) permissions |= PdfStandardPermissions.CopyContents;
        if ((raw & (1 << 5)) != 0) permissions |= PdfStandardPermissions.ModifyAnnotations;
        if ((raw & (1 << 8)) != 0) permissions |= PdfStandardPermissions.FillForms;
        if ((raw & (1 << 9)) != 0) permissions |= PdfStandardPermissions.Accessibility;
        if ((raw & (1 << 10)) != 0) permissions |= PdfStandardPermissions.AssembleDocument;
        if ((raw & (1 << 11)) != 0) permissions |= PdfStandardPermissions.HighQualityPrint;
        return permissions;
    }
}
