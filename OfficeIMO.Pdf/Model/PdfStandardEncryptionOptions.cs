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

    /// <summary>Password required to open the generated PDF.</summary>
    public string UserPassword {
        get => _userPassword;
        set {
            Guard.NotNullOrWhiteSpace(value, nameof(UserPassword));
            _userPassword = value;
        }
    }

    /// <summary>Optional owner password. When omitted, the user password is reused as the owner password.</summary>
    public string? OwnerPassword {
        get => _ownerPassword;
        set {
            if (value != null) {
                Guard.NotNullOrWhiteSpace(value, nameof(OwnerPassword));
            }

            _ownerPassword = value;
        }
    }

    /// <summary>Encryption algorithm. AES-256 revision 6 is the default; RC4 requires explicit legacy selection.</summary>
    public PdfStandardEncryptionAlgorithm Algorithm { get; set; } = PdfStandardEncryptionAlgorithm.Aes256;

    /// <summary>Whether document metadata streams are encrypted. Defaults to true.</summary>
    public bool EncryptMetadata { get; set; } = true;

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
        return new PdfStandardEncryptionOptions(UserPassword) {
            OwnerPassword = OwnerPassword,
            Permissions = Permissions,
            Algorithm = Algorithm,
            EncryptMetadata = EncryptMetadata
        };
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
