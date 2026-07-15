using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Pdf.Cryptography;

#pragma warning disable CA1510 // Cross-target guard code supports netstandard2.0 and net472.

/// <summary>First-party detached CMS signer backed by an RSA certificate available to the current process.</summary>
public sealed class PdfPkcsExternalSigner : IPdfExternalSigner, IDisposable {
    private const string DataOid = "1.2.840.113549.1.7.1";
    private const string SignedDataOid = "1.2.840.113549.1.7.2";
    private const string ContentTypeAttributeOid = "1.2.840.113549.1.9.3";
    private const string MessageDigestAttributeOid = "1.2.840.113549.1.9.4";
    private const string SigningTimeAttributeOid = "1.2.840.113549.1.9.5";
    private const string Sha256Oid = "2.16.840.1.101.3.4.2.1";
    private const string RsaEncryptionOid = "1.2.840.113549.1.1.1";
    private readonly X509Certificate2 _certificate;
    private bool _disposed;

    /// <summary>Creates a signer that embeds the end certificate and emits SHA-256 signed attributes.</summary>
    public PdfPkcsExternalSigner(X509Certificate2 certificate, string? name = null) {
        if (certificate == null) throw new ArgumentNullException(nameof(certificate));
        if (!certificate.HasPrivateKey) throw new ArgumentException("The signing certificate must include a private key.", nameof(certificate));
        _certificate = new X509Certificate2(certificate);
        Name = string.IsNullOrWhiteSpace(name) ? "OfficeIMO.Pdf managed CMS signer" : name!;
    }

    /// <inheritdoc />
    public string Name { get; }

    /// <summary>Whether to include the CMS signing-time attribute. Defaults to true.</summary>
    public bool IncludeSigningTime { get; set; } = true;

    /// <summary>Optional signing time used when <see cref="IncludeSigningTime"/> is enabled. Current UTC time is used when omitted.</summary>
    public DateTimeOffset? SigningTime { get; set; }

    /// <inheritdoc />
    public byte[] Sign(PdfExternalSignatureRequest request) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        ThrowIfDisposed();
        using RSA? rsa = _certificate.GetRSAPrivateKey();
        if (rsa == null) throw new NotSupportedException("The managed CMS signer currently supports RSA certificates.");

        byte[] content = request.SignedContent;
        byte[] digest;
#pragma warning disable CA1850
        using (SHA256 sha256 = SHA256.Create()) digest = sha256.ComputeHash(content);
#pragma warning restore CA1850
        var attributes = new List<byte[]> {
            Attribute(ContentTypeAttributeOid, PdfDerCodec.ObjectIdentifier(DataOid)),
            Attribute(MessageDigestAttributeOid, PdfDerCodec.OctetString(digest))
        };
        if (IncludeSigningTime) attributes.Add(Attribute(SigningTimeAttributeOid, PdfDerCodec.UtcTime(SigningTime ?? DateTimeOffset.UtcNow)));
        byte[] signedAttributes = PdfDerCodec.Set(attributes.ToArray());
        byte[] signature = rsa.SignData(signedAttributes, HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1);

        byte[] serial = _certificate.GetSerialNumber();
        Array.Reverse(serial);
        byte[] signerIdentifier = PdfDerCodec.Sequence(_certificate.IssuerName.RawData, PdfDerCodec.Integer(serial));
        byte[] signerInfo = PdfDerCodec.Sequence(
            PdfDerCodec.Integer(1),
            signerIdentifier,
            PdfDerCodec.AlgorithmIdentifier(Sha256Oid),
            PdfDerCodec.ReplaceTag(signedAttributes, 0xA0),
            PdfDerCodec.AlgorithmIdentifier(RsaEncryptionOid),
            PdfDerCodec.OctetString(signature));
        byte[] signedData = PdfDerCodec.Sequence(
            PdfDerCodec.Integer(1),
            PdfDerCodec.Set(PdfDerCodec.AlgorithmIdentifier(Sha256Oid)),
            PdfDerCodec.Sequence(PdfDerCodec.ObjectIdentifier(DataOid)),
            PdfDerCodec.Context(0, _certificate.RawData),
            PdfDerCodec.Set(signerInfo));
        return PdfDerCodec.Sequence(
            PdfDerCodec.ObjectIdentifier(SignedDataOid),
            PdfDerCodec.Context(0, signedData));
    }

    /// <summary>Releases the cloned certificate and private-key handle.</summary>
    public void Dispose() {
        if (_disposed) return;
        _disposed = true;
        _certificate.Dispose();
    }

    private static byte[] Attribute(string oid, byte[] value) =>
        PdfDerCodec.Sequence(PdfDerCodec.ObjectIdentifier(oid), PdfDerCodec.Set(value));

#pragma warning disable CA1513 // Newer helper is unavailable on netstandard2.0 and net472.
    private void ThrowIfDisposed() {
        if (_disposed) throw new ObjectDisposedException(nameof(PdfPkcsExternalSigner));
    }
#pragma warning restore CA1513
}
