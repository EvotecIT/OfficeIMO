using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Security;

namespace OfficeIMO.Pdf;

/// <summary>
/// Opens PDFs encrypted with the public-key security handler (/Filter /Adobe.PubSec).
/// The security provider decrypts the recipient's CMS EnvelopedData; OfficeIMO.Pdf derives the file key from it.
/// </summary>
/// <remarks>
/// The provider and certificates are deliberately not exposed as public properties, so serializers that snapshot
/// <see cref="PdfLoadOptions"/> (for example batch checkpoints) never copy certificate or key material.
/// </remarks>
public sealed class PdfPublicKeyDecryptionOptions {
    /// <summary>Creates public-key decryption options.</summary>
    /// <param name="securityProvider">Provider that decrypts CMS EnvelopedData, for example <c>OfficeSecurityProvider</c> from OfficeIMO.Security.</param>
    /// <param name="recipientCertificates">Certificates with their private keys. Each is tried against every recipient entry in the file.</param>
    public PdfPublicKeyDecryptionOptions(IOfficeSecurityProvider securityProvider, IEnumerable<X509Certificate2> recipientCertificates) {
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

    /// <summary>Creates public-key decryption options for one recipient certificate.</summary>
    public PdfPublicKeyDecryptionOptions(IOfficeSecurityProvider securityProvider, X509Certificate2 recipientCertificate)
        : this(securityProvider, new[] { recipientCertificate ?? throw new ArgumentNullException(nameof(recipientCertificate)) }) {
    }

    internal IOfficeSecurityProvider SecurityProvider { get; }

    internal IReadOnlyList<X509Certificate2> RecipientCertificates { get; }

    /// <summary>Number of recipient certificates that will be tried.</summary>
    public int CertificateCount => RecipientCertificates.Count;
}
