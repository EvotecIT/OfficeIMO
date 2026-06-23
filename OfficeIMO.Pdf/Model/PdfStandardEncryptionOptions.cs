namespace OfficeIMO.Pdf;

/// <summary>
/// Options for generated PDF Standard password security.
/// </summary>
public sealed class PdfStandardEncryptionOptions {
    /// <summary>PDF permission bit mask allowing all standard operations.</summary>
    public const int AllowAllPermissions = -4;

    private string _userPassword = string.Empty;
    private string? _ownerPassword;

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

    /// <summary>Raw PDF Standard security permission bit mask. Defaults to allowing all standard operations.</summary>
    public int Permissions { get; set; } = AllowAllPermissions;

    /// <summary>Creates a deep copy of the encryption options.</summary>
    public PdfStandardEncryptionOptions Clone() {
        return new PdfStandardEncryptionOptions(UserPassword) {
            OwnerPassword = OwnerPassword,
            Permissions = Permissions
        };
    }
}
