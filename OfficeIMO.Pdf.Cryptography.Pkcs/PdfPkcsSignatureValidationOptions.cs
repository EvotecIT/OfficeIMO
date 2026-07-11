using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Pdf.Cryptography;

/// <summary>Certificate-chain, revocation, and timestamp policy for the optional PKCS provider.</summary>
public sealed class PdfPkcsSignatureValidationOptions {
    /// <summary>Whether to build a signer or TSA certificate chain. Defaults to true.</summary>
    public bool ValidateCertificateChain { get; set; } = true;

    /// <summary>Revocation policy used by X509Chain. Defaults to NoCheck to avoid implicit network access.</summary>
    public X509RevocationMode RevocationMode { get; set; } = X509RevocationMode.NoCheck;

    /// <summary>Certificate portion covered by revocation checking.</summary>
    public X509RevocationFlag RevocationFlag { get; set; } = X509RevocationFlag.ExcludeRoot;

    /// <summary>Additional certificate-chain verification flags.</summary>
    public X509VerificationFlags VerificationFlags { get; set; } = X509VerificationFlags.NoFlag;

    /// <summary>Optional verification time. Current system time is used when omitted.</summary>
    public DateTime? VerificationTime { get; set; }

    /// <summary>Timeout used for certificate retrieval by platform chain policy.</summary>
    public TimeSpan UrlRetrievalTimeout { get; set; } = TimeSpan.FromSeconds(15);

    /// <summary>Whether RFC 3161 document and signature timestamps should be validated.</summary>
    public bool ValidateTimestamps { get; set; } = true;

    /// <summary>Additional intermediate, root, signer, or TSA certificate candidates.</summary>
    public X509Certificate2Collection ExtraCertificates { get; } = new X509Certificate2Collection();

    /// <summary>Optional application trust callback. Return true to accept the built chain under caller policy.</summary>
    public Func<X509Certificate2, X509Chain, bool>? ChainEvaluator { get; set; }
}
