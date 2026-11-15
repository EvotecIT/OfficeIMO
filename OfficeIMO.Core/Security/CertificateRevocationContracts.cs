using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security;

/// <summary>Where signature validation may obtain OCSP responses and CRLs.</summary>
public enum CertificateRevocationPolicy {
    /// <summary>Use only revocation data carried by the signature or document. Never starts network retrieval.</summary>
    EmbeddedOnly = 0,

    /// <summary>Use embedded revocation data first and fetch online data for certificates it does not settle.</summary>
    EmbeddedThenOnline = 1,

    /// <summary>Ignore embedded revocation data and decide from freshly fetched online data only.</summary>
    Online = 2
}

/// <summary>Retrieves OCSP responses and CRLs for one certificate.</summary>
public interface ICertificateRevocationFetcher {
    /// <summary>
    /// Fetches revocation data for <paramref name="certificate"/>, issued by <paramref name="issuer"/>.
    /// Returned data is untrusted input; validators check its signature, issuer and freshness themselves.
    /// </summary>
    Task<CertificateRevocationEvidence> FetchAsync(
        X509Certificate2 certificate,
        X509Certificate2 issuer,
        CancellationToken cancellationToken = default);
}

/// <summary>DER-encoded certificates, OCSP responses and CRLs gathered for certificate validation.</summary>
public sealed class CertificateRevocationEvidence {
    private readonly byte[][] _certificates;
    private readonly byte[][] _ocspResponses;
    private readonly byte[][] _certificateRevocationLists;

    /// <summary>Creates revocation evidence. Values are copied.</summary>
    public CertificateRevocationEvidence(
        IEnumerable<byte[]>? certificates = null,
        IEnumerable<byte[]>? ocspResponses = null,
        IEnumerable<byte[]>? certificateRevocationLists = null,
        IEnumerable<SecurityFinding>? findings = null) {
        _certificates = Snapshot(certificates, nameof(certificates));
        _ocspResponses = Snapshot(ocspResponses, nameof(ocspResponses));
        _certificateRevocationLists = Snapshot(certificateRevocationLists, nameof(certificateRevocationLists));
        Findings = findings?.ToArray() ?? Array.Empty<SecurityFinding>();
    }

    /// <summary>Empty evidence without findings.</summary>
    public static CertificateRevocationEvidence Empty { get; } = new CertificateRevocationEvidence();

    /// <summary>DER-encoded certificates, such as the issuer path used for the request.</summary>
    public IReadOnlyList<byte[]> Certificates => Clone(_certificates);

    /// <summary>DER-encoded OCSPResponse values.</summary>
    public IReadOnlyList<byte[]> OcspResponses => Clone(_ocspResponses);

    /// <summary>DER-encoded X.509 CRLs.</summary>
    public IReadOnlyList<byte[]> CertificateRevocationLists => Clone(_certificateRevocationLists);

    /// <summary>Why a source could not be used, such as a refused address, a timeout or an oversized response.</summary>
    public IReadOnlyList<SecurityFinding> Findings { get; }

    internal IReadOnlyList<byte[]> CertificateValues => _certificates;
    internal IReadOnlyList<byte[]> OcspValues => _ocspResponses;
    internal IReadOnlyList<byte[]> CrlValues => _certificateRevocationLists;

    private static byte[][] Snapshot(IEnumerable<byte[]>? values, string parameterName) {
        if (values is null) return Array.Empty<byte[]>();
        byte[][] result = values.Select(static value => value is null ? null! : (byte[])value.Clone()).ToArray();
        if (result.Any(static value => value is null || value.Length == 0)) {
            throw new ArgumentException("Revocation evidence cannot contain null or empty values.", parameterName);
        }
        return result;
    }

    private static IReadOnlyList<byte[]> Clone(byte[][] values) =>
        Array.AsReadOnly(values.Select(static value => (byte[])value.Clone()).ToArray());
}

/// <summary>Origin of one revocation token considered during validation.</summary>
internal enum OfficeRevocationTokenSource {
    /// <summary>Carried inside the CMS signature or time-stamp token.</summary>
    Signature = 0,
    /// <summary>Carried by the document, such as a PDF Document Security Store.</summary>
    Document = 1,
    /// <summary>Fetched online during this validation.</summary>
    Online = 2
}

/// <summary>One OCSP response or CRL supplied by a document, with its proof-of-existence time when known.</summary>
internal sealed class OfficeRevocationToken {
    internal OfficeRevocationToken(byte[] encoded, bool isOcspResponse, DateTimeOffset? proofOfExistence) {
        Encoded = encoded;
        IsOcspResponse = isOcspResponse;
        ProofOfExistence = proofOfExistence;
    }

    internal byte[] Encoded { get; }
    internal bool IsOcspResponse { get; }

    /// <summary>Earliest trusted time stamp proving the token existed, such as a covering PDF document time-stamp.</summary>
    internal DateTimeOffset? ProofOfExistence { get; }
}

/// <summary>Revocation validation request for one CMS signature or RFC 3161 time-stamp token.</summary>
internal sealed class OfficeRevocationRequest {
    internal OfficeRevocationRequest(byte[] encodedCms, bool isTimestampToken) {
        EncodedCms = encodedCms;
        IsTimestampToken = isTimestampToken;
    }

    internal byte[] EncodedCms { get; }
    internal bool IsTimestampToken { get; }

    /// <summary>Time at which the signer certificate path must not have been revoked.</summary>
    internal DateTimeOffset? SignerValidationTime { get; set; }

    internal IReadOnlyList<byte[]> DocumentCertificates { get; set; } = Array.Empty<byte[]>();
    internal IReadOnlyList<OfficeRevocationToken> DocumentTokens { get; set; } = Array.Empty<OfficeRevocationToken>();
    internal CertificateRevocationPolicy Policy { get; set; }
    internal ICertificateRevocationFetcher? Fetcher { get; set; }
    internal X509Certificate2Collection? ExtraCertificates { get; set; }
    internal DateTimeOffset? Now { get; set; }
    internal CancellationToken CancellationToken { get; set; }
}

/// <summary>Revocation outcome for one CMS signature and its signature time-stamp tokens.</summary>
internal sealed class OfficeRevocationOutcome {
    internal OfficeRevocationOutcome(
        SecurityValidationStatus status,
        SecurityValidationStatus timestampAuthorityStatus,
        bool hasSigningCertificateAttribute,
        bool signingCertificateMismatch,
        bool settledByDocument,
        bool documentEvidenceHasProofOfExistence,
        IReadOnlyList<SecurityFinding> findings) {
        Status = status;
        TimestampAuthorityStatus = timestampAuthorityStatus;
        HasSigningCertificateAttribute = hasSigningCertificateAttribute;
        SigningCertificateMismatch = signingCertificateMismatch;
        SettledByDocument = settledByDocument;
        DocumentEvidenceHasProofOfExistence = documentEvidenceHasProofOfExistence;
        Findings = findings;
    }

    /// <summary>Aggregate status for the signer path and every signature time-stamp authority path.</summary>
    internal SecurityValidationStatus Status { get; }

    /// <summary>Status for signature time-stamp authority paths only.</summary>
    internal SecurityValidationStatus TimestampAuthorityStatus { get; }

    /// <summary>True when every CMS signer carries an ESS signing-certificate(-v2) attribute that matches its certificate.</summary>
    internal bool HasSigningCertificateAttribute { get; }

    /// <summary>True when an ESS signing-certificate attribute identifies a different certificate than the signer's.</summary>
    internal bool SigningCertificateMismatch { get; }

    /// <summary>True when every checked certificate was settled by document-supplied tokens.</summary>
    internal bool SettledByDocument { get; }

    /// <summary>True when every document token used to settle a certificate has a proof-of-existence time.</summary>
    internal bool DocumentEvidenceHasProofOfExistence { get; }

    internal IReadOnlyList<SecurityFinding> Findings { get; }
}

/// <summary>Revocation engine implemented by the optional OfficeIMO.Security provider.</summary>
internal interface IOfficeRevocationValidator {
    /// <summary>Checks the signer and signature time-stamp authority paths of one CMS container.</summary>
    OfficeRevocationOutcome ValidateRevocation(OfficeRevocationRequest request);

    /// <summary>Fetches revocation data for the signer and time-stamp authority paths of one CMS container.</summary>
    Task<CertificateRevocationEvidence> FetchRevocationEvidenceAsync(
        byte[] encodedCms,
        bool isTimestampToken,
        ICertificateRevocationFetcher fetcher,
        X509Certificate2Collection? extraCertificates,
        CancellationToken cancellationToken);
}
