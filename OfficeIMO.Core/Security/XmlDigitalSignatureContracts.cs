using System;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security;

/// <summary>Well-known XML Digital Signature algorithm identifiers used by document formats.</summary>
public static class XmlDigitalSignatureAlgorithms {
    /// <summary>XML Digital Signature namespace.</summary>
    public const string Namespace = "http://www.w3.org/2000/09/xmldsig#";
    /// <summary>Inclusive XML canonicalization without comments.</summary>
    public const string CanonicalXml = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315";
    /// <summary>Inclusive XML canonicalization with comments.</summary>
    public const string CanonicalXmlWithComments = "http://www.w3.org/TR/2001/REC-xml-c14n-20010315#WithComments";
    /// <summary>Exclusive XML canonicalization without comments.</summary>
    public const string ExclusiveCanonicalXml = "http://www.w3.org/2001/10/xml-exc-c14n#";
    /// <summary>Exclusive XML canonicalization with comments.</summary>
    public const string ExclusiveCanonicalXmlWithComments = "http://www.w3.org/2001/10/xml-exc-c14n#WithComments";
    /// <summary>Enveloped-signature transform.</summary>
    public const string EnvelopedSignatureTransform = "http://www.w3.org/2000/09/xmldsig#enveloped-signature";
    /// <summary>RSA with SHA-1 signature method.</summary>
    public const string RsaSha1 = "http://www.w3.org/2000/09/xmldsig#rsa-sha1";
    /// <summary>RSA with SHA-256 signature method.</summary>
    public const string RsaSha256 = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256";
    /// <summary>RSA with SHA-384 signature method.</summary>
    public const string RsaSha384 = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha384";
    /// <summary>RSA with SHA-512 signature method.</summary>
    public const string RsaSha512 = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha512";
    /// <summary>SHA-1 digest method retained for validation of legacy Office signatures.</summary>
    public const string Sha1 = "http://www.w3.org/2000/09/xmldsig#sha1";
    /// <summary>SHA-256 digest method.</summary>
    public const string Sha256 = "http://www.w3.org/2001/04/xmlenc#sha256";
    /// <summary>SHA-384 digest method.</summary>
    public const string Sha384 = "http://www.w3.org/2001/04/xmldsig-more#sha384";
    /// <summary>SHA-512 digest method.</summary>
    public const string Sha512 = "http://www.w3.org/2001/04/xmlenc#sha512";
}

/// <summary>Request for creating one bounded enveloping XML signature over a caller-supplied object.</summary>
public sealed class XmlDigitalSignatureCreationRequest {
    /// <summary>Creates a request over the children of the supplied wrapper XML document.</summary>
    public XmlDigitalSignatureCreationRequest(
        byte[] objectXml,
        X509Certificate2 signingCertificate,
        string signatureId,
        string objectId,
        string objectReferenceType,
        string canonicalizationMethod,
        string signatureMethod,
        string digestMethod) {
        ObjectXml = objectXml ?? throw new ArgumentNullException(nameof(objectXml));
        SigningCertificate = signingCertificate ?? throw new ArgumentNullException(nameof(signingCertificate));
        SignatureId = signatureId ?? throw new ArgumentNullException(nameof(signatureId));
        ObjectId = objectId ?? throw new ArgumentNullException(nameof(objectId));
        ObjectReferenceType = objectReferenceType ?? throw new ArgumentNullException(nameof(objectReferenceType));
        CanonicalizationMethod = canonicalizationMethod ?? throw new ArgumentNullException(nameof(canonicalizationMethod));
        SignatureMethod = signatureMethod ?? throw new ArgumentNullException(nameof(signatureMethod));
        DigestMethod = digestMethod ?? throw new ArgumentNullException(nameof(digestMethod));
    }

    /// <summary>Gets the bounded XML wrapper whose child nodes become the signed DataObject content.</summary>
    public byte[] ObjectXml { get; }
    /// <summary>Gets the certificate whose private key creates the signature.</summary>
    public X509Certificate2 SigningCertificate { get; }
    /// <summary>Gets the XML Signature Id.</summary>
    public string SignatureId { get; }
    /// <summary>Gets the signed DataObject Id.</summary>
    public string ObjectId { get; }
    /// <summary>Gets the SignedInfo Reference Type for the DataObject.</summary>
    public string ObjectReferenceType { get; }
    /// <summary>Gets the SignedInfo canonicalization algorithm.</summary>
    public string CanonicalizationMethod { get; }
    /// <summary>Gets the signature algorithm.</summary>
    public string SignatureMethod { get; }
    /// <summary>Gets the DataObject digest algorithm.</summary>
    public string DigestMethod { get; }
    /// <summary>Gets additional certificates embedded after the signer certificate.</summary>
    public IReadOnlyCollection<X509Certificate2>? AdditionalCertificates { get; set; }
    /// <summary>Gets or sets the maximum object XML bytes accepted. Defaults to 16 MiB.</summary>
    public long MaxObjectBytes { get; set; } = 16L * 1024L * 1024L;
    /// <summary>Gets or sets the maximum encoded signature XML bytes returned. Defaults to 16 MiB.</summary>
    public long MaxOutputBytes { get; set; } = 16L * 1024L * 1024L;
}

/// <summary>Request for bounded verification of XML signature math and local signed-object references.</summary>
public sealed class XmlDigitalSignatureVerificationRequest {
    /// <summary>Creates a request for one encoded XML Signature document.</summary>
    public XmlDigitalSignatureVerificationRequest(
        byte[] signatureXml,
        IReadOnlyCollection<X509Certificate2> certificateCandidates) {
        SignatureXml = signatureXml ?? throw new ArgumentNullException(nameof(signatureXml));
        CertificateCandidates = certificateCandidates ?? throw new ArgumentNullException(nameof(certificateCandidates));
    }

    /// <summary>Gets the encoded XML Signature document.</summary>
    public byte[] SignatureXml { get; }
    /// <summary>Gets bounded signer-certificate candidates.</summary>
    public IReadOnlyCollection<X509Certificate2> CertificateCandidates { get; }
    /// <summary>Gets or sets the maximum signature XML bytes accepted. Defaults to 16 MiB.</summary>
    public long MaxSignatureBytes { get; set; } = 16L * 1024L * 1024L;
    /// <summary>Gets or sets the maximum SignedInfo Reference count. Defaults to 4,096.</summary>
    public int MaxReferences { get; set; } = 4096;
    /// <summary>Gets or sets the maximum aggregate local-reference work across candidates. Defaults to 512 MiB.</summary>
    public long MaxTotalDigestWorkBytes { get; set; } = 512L * 1024L * 1024L;
    /// <summary>Gets or sets the accepted SignedInfo canonicalization algorithms.</summary>
    public IReadOnlyCollection<string> AllowedCanonicalizationMethods { get; set; } = new[] {
        XmlDigitalSignatureAlgorithms.CanonicalXml,
        XmlDigitalSignatureAlgorithms.CanonicalXmlWithComments,
        XmlDigitalSignatureAlgorithms.ExclusiveCanonicalXml,
        XmlDigitalSignatureAlgorithms.ExclusiveCanonicalXmlWithComments
    };
    /// <summary>Gets or sets the accepted transforms on local SignedInfo references.</summary>
    public IReadOnlyCollection<string> AllowedReferenceTransforms { get; set; } = new[] {
        XmlDigitalSignatureAlgorithms.CanonicalXml,
        XmlDigitalSignatureAlgorithms.CanonicalXmlWithComments,
        XmlDigitalSignatureAlgorithms.ExclusiveCanonicalXml,
        XmlDigitalSignatureAlgorithms.ExclusiveCanonicalXmlWithComments,
        XmlDigitalSignatureAlgorithms.EnvelopedSignatureTransform
    };
    /// <summary>Gets or sets the accepted signature methods. Caller policy may narrow, but cannot expand, the provider's supported set.</summary>
    public IReadOnlyCollection<string> AllowedSignatureMethods { get; set; } = new[] {
        XmlDigitalSignatureAlgorithms.RsaSha1,
        XmlDigitalSignatureAlgorithms.RsaSha256,
        XmlDigitalSignatureAlgorithms.RsaSha384,
        XmlDigitalSignatureAlgorithms.RsaSha512
    };
    /// <summary>Gets or sets the accepted digest methods. Caller policy may narrow, but cannot expand, the provider's supported set.</summary>
    public IReadOnlyCollection<string> AllowedDigestMethods { get; set; } = new[] {
        XmlDigitalSignatureAlgorithms.Sha1,
        XmlDigitalSignatureAlgorithms.Sha256,
        XmlDigitalSignatureAlgorithms.Sha384,
        XmlDigitalSignatureAlgorithms.Sha512
    };
}

/// <summary>Mathematical XML signature verification evidence independent of certificate trust.</summary>
public sealed class XmlDigitalSignatureVerificationResult {
    /// <summary>Creates XML signature verification evidence for a provider implementation.</summary>
    public XmlDigitalSignatureVerificationResult(
        SecurityValidationStatus status,
        IReadOnlyList<X509Certificate2> matchingCertificates,
        IReadOnlyList<SecurityFinding> findings) {
        Status = status;
        MatchingCertificates = matchingCertificates;
        Findings = findings;
    }

    /// <summary>Gets the mathematical XML signature outcome.</summary>
    public SecurityValidationStatus Status { get; }
    /// <summary>Gets supplied certificates whose public keys validate the signature.</summary>
    public IReadOnlyList<X509Certificate2> MatchingCertificates { get; }
    /// <summary>Gets structured validation findings.</summary>
    public IReadOnlyList<SecurityFinding> Findings { get; }
}
