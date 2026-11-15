using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

namespace OfficeIMO.Pdf;

/// <summary>
/// Encrypts a PDF for certificate recipients with the public-key security handler (/Filter /Adobe.PubSec,
/// /SubFilter /adbe.pkcs7.s5, AES-256). Each recipient opens the file with its private key; no password exists.
/// </summary>
/// <remarks>
/// The security provider creates the CMS EnvelopedData that carries the file key seed and permissions to every recipient.
/// Certificates and the provider are not exposed as public properties, so option snapshots never copy certificate material.
/// </remarks>
public sealed class PdfPublicKeyEncryptionOptions {
    private int _permissions = PdfStandardEncryptionOptions.AllowAllPermissions;

    /// <summary>Creates public-key encryption options for one or more recipient certificates.</summary>
    /// <param name="securityProvider">Provider that creates CMS EnvelopedData, for example <c>OfficeSecurityProvider</c> from OfficeIMO.Security.</param>
    /// <param name="recipientCertificates">Recipient certificates. Only their public keys are used.</param>
    public PdfPublicKeyEncryptionOptions(IOfficeSecurityProvider securityProvider, IEnumerable<X509Certificate2> recipientCertificates) {
        Guard.NotNull(securityProvider, nameof(securityProvider));
        Guard.NotNull(recipientCertificates, nameof(recipientCertificates));
        X509Certificate2[] certificates = recipientCertificates.ToArray();
        if (certificates.Length == 0) {
            throw new ArgumentException("At least one recipient certificate is required.", nameof(recipientCertificates));
        }

        foreach (X509Certificate2 certificate in certificates) {
            if (certificate is null) {
                throw new ArgumentException("Recipient certificates cannot contain null entries.", nameof(recipientCertificates));
            }
        }

        SecurityProvider = securityProvider;
        RecipientCertificates = Array.AsReadOnly(certificates);
    }

    /// <summary>Creates public-key encryption options for one recipient certificate.</summary>
    public PdfPublicKeyEncryptionOptions(IOfficeSecurityProvider securityProvider, X509Certificate2 recipientCertificate)
        : this(securityProvider, new[] { recipientCertificate ?? throw new ArgumentNullException(nameof(recipientCertificate)) }) {
    }

    internal IOfficeSecurityProvider SecurityProvider { get; }

    internal IReadOnlyList<X509Certificate2> RecipientCertificates { get; }

    /// <summary>Number of recipient certificates the document is encrypted to.</summary>
    public int RecipientCount => RecipientCertificates.Count;

    /// <summary>Operations every recipient may perform. Defaults to all standard operations.</summary>
    public PdfStandardPermissions AllowedPermissions {
        get => PdfStandardEncryptionOptions.FromRawPermissions(_permissions);
        set {
            if ((value & ~PdfStandardPermissions.All) != 0) throw new ArgumentOutOfRangeException(nameof(value));
            _permissions = PdfStandardEncryptionOptions.ToRawPermissions(value);
        }
    }

    internal int Permissions => _permissions;

    /// <summary>Whether document metadata streams are encrypted. Defaults to true.</summary>
    public bool EncryptMetadata { get; set; } = true;

    /// <summary>
    /// Optional synchronous AES provider used when the host does not expose platform AES-CBC, such as browser WebAssembly.
    /// </summary>
    public IOfficeAesCryptographyProvider? AesCryptographyProvider { get; set; }
}
