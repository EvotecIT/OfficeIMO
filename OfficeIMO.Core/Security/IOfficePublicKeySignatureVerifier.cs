namespace OfficeIMO.Security;

/// <summary>Raw public-key signature schemes used by message-level formats such as DKIM and ARC.</summary>
public enum OfficePublicKeySignatureAlgorithm {
    /// <summary>RSASSA-PKCS1-v1_5 with SHA-256.</summary>
    RsaPkcs1Sha256 = 0,

    /// <summary>RSASSA-PKCS1-v1_5 with SHA-1. Callers decide whether the result is acceptable.</summary>
    RsaPkcs1Sha1 = 1,

    /// <summary>PureEdDSA over Ed25519 (RFC 8032). The supplied data is signed as-is.</summary>
    Ed25519 = 2
}

/// <summary>Outcome of one raw public-key signature check.</summary>
public enum OfficePublicKeySignatureStatus {
    /// <summary>The signature matches the supplied data and key.</summary>
    Valid = 0,

    /// <summary>The key and signature were well formed, but the signature does not match.</summary>
    Invalid = 1,

    /// <summary>The public key could not be decoded for the requested algorithm or is outside supported bounds.</summary>
    MalformedKey = 2,

    /// <summary>The signature value is structurally unusable for the key (for example, the wrong length).</summary>
    MalformedSignature = 3,

    /// <summary>The verifier does not implement the requested algorithm.</summary>
    UnsupportedAlgorithm = 4
}

/// <summary>Result of verifying one raw public-key signature.</summary>
public sealed class OfficePublicKeySignatureCheck {
    /// <summary>Creates a signature check result.</summary>
    public OfficePublicKeySignatureCheck(OfficePublicKeySignatureStatus status, int keySizeBits = 0, string? message = null) {
        Status = status;
        KeySizeBits = keySizeBits;
        Message = message;
    }

    /// <summary>Verification outcome.</summary>
    public OfficePublicKeySignatureStatus Status { get; }

    /// <summary>Public key size in bits when the key was decoded; otherwise zero.</summary>
    public int KeySizeBits { get; }

    /// <summary>Optional diagnostic detail. It never contains key or signature material.</summary>
    public string? Message { get; }

    /// <summary>True only when the signature matched.</summary>
    public bool IsValid => Status == OfficePublicKeySignatureStatus.Valid;
}

/// <summary>
/// Verifies raw public-key signatures over caller-supplied bytes. The contract is dependency-free; implementations
/// such as <c>OfficeIMO.Security</c> supply the cryptography. Implementations must not throw for hostile input.
/// </summary>
public interface IOfficePublicKeySignatureVerifier {
    /// <summary>Gets a stable provider name for diagnostics.</summary>
    string Name { get; }

    /// <summary>Returns true when the verifier implements <paramref name="algorithm"/>.</summary>
    bool Supports(OfficePublicKeySignatureAlgorithm algorithm);

    /// <summary>Verifies <paramref name="signature"/> over <paramref name="data"/>.</summary>
    /// <param name="algorithm">Signature scheme.</param>
    /// <param name="publicKey">
    /// For RSA, a DER SubjectPublicKeyInfo or PKCS #1 RSAPublicKey. For Ed25519, the raw 32-byte public key.
    /// </param>
    /// <param name="data">Exact signed bytes. RSA schemes hash them; Ed25519 signs them directly.</param>
    /// <param name="signature">Signature value.</param>
    OfficePublicKeySignatureCheck Verify(
        OfficePublicKeySignatureAlgorithm algorithm,
        byte[] publicKey,
        byte[] data,
        byte[] signature);
}
