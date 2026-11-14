namespace OfficeIMO.Security;

/// <summary>
/// Keys supplied by the caller for one OpenPGP operation. OpenPGP providers never read a user keyring, agent, or
/// key server on their own; only the key material in this set is used.
/// </summary>
public sealed class OfficeOpenPgpKeySet {
    /// <summary>Creates a key set.</summary>
    /// <param name="publicKeys">Transferable public keys, binary or ASCII-armored. One buffer may hold several keys.</param>
    /// <param name="secretKeys">Transferable secret keys, binary or ASCII-armored, used for decryption.</param>
    /// <param name="passphraseProvider">
    /// Returns the passphrase for a secret key, identified by its 16-hex-digit key ID, or null when none is available.
    /// </param>
    public OfficeOpenPgpKeySet(
        IEnumerable<byte[]>? publicKeys = null,
        IEnumerable<byte[]>? secretKeys = null,
        Func<string, char[]?>? passphraseProvider = null) {
        PublicKeys = publicKeys?.Where(key => key != null).ToArray() ?? Array.Empty<byte[]>();
        SecretKeys = secretKeys?.Where(key => key != null).ToArray() ?? Array.Empty<byte[]>();
        PassphraseProvider = passphraseProvider;
    }

    /// <summary>Transferable public keys.</summary>
    public IReadOnlyList<byte[]> PublicKeys { get; }

    /// <summary>Transferable secret keys.</summary>
    public IReadOnlyList<byte[]> SecretKeys { get; }

    /// <summary>Passphrase callback keyed by 16-hex-digit key ID.</summary>
    public Func<string, char[]?>? PassphraseProvider { get; }
}

/// <summary>Resource limits and policy for OpenPGP operations.</summary>
public sealed class OfficeOpenPgpOptions {
    /// <summary>Maximum encoded message, signature, or key buffer size.</summary>
    public long MaxInputBytes { get; set; } = 64L * 1024 * 1024;

    /// <summary>Maximum decrypted and decompressed output size.</summary>
    public long MaxOutputBytes { get; set; } = 256L * 1024 * 1024;

    /// <summary>Maximum number of signatures evaluated in one operation.</summary>
    public int MaxSignatures { get; set; } = 32;

    /// <summary>Maximum nesting of compressed packets.</summary>
    public int MaxNestingDepth { get; set; } = 8;

    /// <summary>
    /// When false (the default), encrypted data without a modification detection code is refused and its plaintext
    /// is not released.
    /// </summary>
    public bool AllowMissingIntegrityProtection { get; set; }

    /// <summary>When false (the default), signatures using MD5, SHA-1, or RIPEMD-160 are reported as weak.</summary>
    public bool AllowWeakHashAlgorithms { get; set; }

    /// <summary>Minimum RSA, DSA, or ElGamal key size accepted for signatures.</summary>
    public int MinimumKeyBits { get; set; } = 2048;

    /// <summary>Time used for signature and key expiration. Defaults to the current UTC time.</summary>
    public DateTimeOffset? VerificationTime { get; set; }
}

/// <summary>Outcome of one OpenPGP signature check.</summary>
public enum OfficeOpenPgpSignatureStatus {
    /// <summary>The signature matches the data and the signing key is usable for signing.</summary>
    Valid = 0,
    /// <summary>The signature does not match the data.</summary>
    Invalid = 1,
    /// <summary>No supplied public key matches the issuer key ID.</summary>
    MissingPublicKey = 2,
    /// <summary>The signing key or its primary key is revoked.</summary>
    KeyRevoked = 3,
    /// <summary>The signing key had expired, or did not exist yet, when the signature was made.</summary>
    KeyExpired = 4,
    /// <summary>The signature has expired.</summary>
    SignatureExpired = 5,
    /// <summary>The signing subkey has no valid binding, or its key flags exclude signing.</summary>
    KeyNotValidForSigning = 6,
    /// <summary>The hash algorithm or key size is refused by policy.</summary>
    WeakAlgorithm = 7,
    /// <summary>The signature packet version, type, or algorithm is not supported.</summary>
    Unsupported = 8,
    /// <summary>The signature data could not be parsed.</summary>
    Malformed = 9
}

/// <summary>Evaluation of one OpenPGP signature. Trust in the signer's identity remains the caller's decision.</summary>
public sealed class OfficeOpenPgpSignatureInfo {
    /// <summary>Creates a signature evaluation.</summary>
    public OfficeOpenPgpSignatureInfo(
        OfficeOpenPgpSignatureStatus status,
        string message,
        string? keyId = null,
        string? signingKeyFingerprint = null,
        string? primaryKeyFingerprint = null,
        IEnumerable<string>? userIds = null,
        DateTimeOffset? createdAt = null,
        string? hashAlgorithm = null,
        string? publicKeyAlgorithm = null,
        int signatureVersion = 0,
        int keyVersion = 0,
        bool isTextSignature = false) {
        Status = status;
        Message = message ?? string.Empty;
        KeyId = keyId;
        SigningKeyFingerprint = signingKeyFingerprint;
        PrimaryKeyFingerprint = primaryKeyFingerprint;
        UserIds = userIds?.ToArray() ?? Array.Empty<string>();
        CreatedAt = createdAt;
        HashAlgorithm = hashAlgorithm;
        PublicKeyAlgorithm = publicKeyAlgorithm;
        SignatureVersion = signatureVersion;
        KeyVersion = keyVersion;
        IsTextSignature = isTextSignature;
    }

    /// <summary>Signature status.</summary>
    public OfficeOpenPgpSignatureStatus Status { get; }

    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }

    /// <summary>Issuer key ID as 16 uppercase hex digits.</summary>
    public string? KeyId { get; }

    /// <summary>Fingerprint of the key that made the signature, when it was found.</summary>
    public string? SigningKeyFingerprint { get; }

    /// <summary>Fingerprint of the primary key of the signer's certificate, when it was found.</summary>
    public string? PrimaryKeyFingerprint { get; }

    /// <summary>User IDs on the signer's certificate. They are claims, not verified identities.</summary>
    public IReadOnlyList<string> UserIds { get; }

    /// <summary>Signature creation time.</summary>
    public DateTimeOffset? CreatedAt { get; }

    /// <summary>Hash algorithm name.</summary>
    public string? HashAlgorithm { get; }

    /// <summary>Public-key algorithm name.</summary>
    public string? PublicKeyAlgorithm { get; }

    /// <summary>Signature packet version.</summary>
    public int SignatureVersion { get; }

    /// <summary>Signing key packet version.</summary>
    public int KeyVersion { get; }

    /// <summary>True for canonical text signatures (type 0x01).</summary>
    public bool IsTextSignature { get; }

    /// <summary>True only when <see cref="Status"/> is <see cref="OfficeOpenPgpSignatureStatus.Valid"/>.</summary>
    public bool IsValid => Status == OfficeOpenPgpSignatureStatus.Valid;
}

/// <summary>Result of verifying a detached or cleartext OpenPGP signature.</summary>
public sealed class OfficeOpenPgpVerificationResult {
    /// <summary>Creates a verification result.</summary>
    public OfficeOpenPgpVerificationResult(IEnumerable<OfficeOpenPgpSignatureInfo>? signatures, string? message = null,
        byte[]? signedContent = null) {
        Signatures = signatures?.ToArray() ?? Array.Empty<OfficeOpenPgpSignatureInfo>();
        Message = message;
        SignedContent = signedContent;
    }

    /// <summary>Every signature found, in packet order.</summary>
    public IReadOnlyList<OfficeOpenPgpSignatureInfo> Signatures { get; }

    /// <summary>Explanation when no signature could be read.</summary>
    public string? Message { get; }

    /// <summary>For cleartext-signed messages, the signed text with dash-escaping removed and CRLF line endings.</summary>
    public byte[]? SignedContent { get; }

    /// <summary>True when at least one signature was found and every signature is valid.</summary>
    public bool IsValid => Signatures.Count > 0 && Signatures.All(signature => signature.IsValid);
}

/// <summary>Outcome of opening an OpenPGP message.</summary>
public enum OfficeOpenPgpMessageStatus {
    /// <summary>The message was decrypted (when encrypted) and its literal data is available.</summary>
    Opened = 0,
    /// <summary>No supplied secret key matches a recipient of the message.</summary>
    NoMatchingSecretKey = 1,
    /// <summary>The matching secret key is protected and no passphrase was supplied.</summary>
    PassphraseRequired = 2,
    /// <summary>The supplied passphrase did not unlock the matching secret key.</summary>
    WrongPassphrase = 3,
    /// <summary>The modification detection code did not match; the plaintext is not released.</summary>
    IntegrityCheckFailed = 4,
    /// <summary>The encrypted data has no modification detection code; the plaintext is not released.</summary>
    MissingIntegrityProtection = 5,
    /// <summary>The message uses a packet or algorithm the provider does not support.</summary>
    Unsupported = 6,
    /// <summary>The message could not be parsed or decrypted.</summary>
    Malformed = 7,
    /// <summary>A size, signature, or nesting limit was reached.</summary>
    LimitExceeded = 8
}

/// <summary>Result of decrypting and/or reading a signed OpenPGP message.</summary>
public sealed class OfficeOpenPgpMessageResult {
    /// <summary>Creates a message result.</summary>
    public OfficeOpenPgpMessageResult(
        OfficeOpenPgpMessageStatus status,
        string message,
        bool isEncrypted = false,
        bool isIntegrityProtected = false,
        byte[]? content = null,
        string? fileName = null,
        string? symmetricAlgorithm = null,
        IEnumerable<string>? recipientKeyIds = null,
        IEnumerable<OfficeOpenPgpSignatureInfo>? signatures = null) {
        Status = status;
        Message = message ?? string.Empty;
        IsEncrypted = isEncrypted;
        IsIntegrityProtected = isIntegrityProtected;
        Content = status == OfficeOpenPgpMessageStatus.Opened ? content : null;
        FileName = fileName;
        SymmetricAlgorithm = symmetricAlgorithm;
        RecipientKeyIds = recipientKeyIds?.ToArray() ?? Array.Empty<string>();
        Signatures = signatures?.ToArray() ?? Array.Empty<OfficeOpenPgpSignatureInfo>();
    }

    /// <summary>Overall status.</summary>
    public OfficeOpenPgpMessageStatus Status { get; }

    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }

    /// <summary>True when the message was encrypted.</summary>
    public bool IsEncrypted { get; }

    /// <summary>True when the encrypted data carried a modification detection code that matched.</summary>
    public bool IsIntegrityProtected { get; }

    /// <summary>Literal data. Set only when <see cref="Status"/> is <see cref="OfficeOpenPgpMessageStatus.Opened"/>.</summary>
    public byte[]? Content { get; }

    /// <summary>File name recorded in the literal data packet.</summary>
    public string? FileName { get; }

    /// <summary>Symmetric algorithm used for the session key.</summary>
    public string? SymmetricAlgorithm { get; }

    /// <summary>Recipient key IDs listed in the message, as 16 uppercase hex digits.</summary>
    public IReadOnlyList<string> RecipientKeyIds { get; }

    /// <summary>Signatures over the literal data (signed-then-encrypted or signed-only messages).</summary>
    public IReadOnlyList<OfficeOpenPgpSignatureInfo> Signatures { get; }

    /// <summary>True when the message carried at least one signature and all of them are valid.</summary>
    public bool IsSignatureValid => Signatures.Count > 0 && Signatures.All(signature => signature.IsValid);
}

/// <summary>
/// OpenPGP (RFC 4880) operations over caller-supplied keys. The contract is dependency-free; implementations such as
/// <c>OfficeIMO.Security</c> supply the cryptography. Implementations must not throw for hostile message input.
/// </summary>
public interface IOfficeOpenPgpProvider {
    /// <summary>Gets a stable provider name for diagnostics.</summary>
    string Name { get; }

    /// <summary>Verifies a detached signature (binary or armored) over the exact supplied content.</summary>
    OfficeOpenPgpVerificationResult VerifyDetached(byte[] signature, byte[] signedContent, OfficeOpenPgpKeySet keys,
        OfficeOpenPgpOptions? options = null);

    /// <summary>Verifies an inline cleartext-signed message ("BEGIN PGP SIGNED MESSAGE").</summary>
    OfficeOpenPgpVerificationResult VerifyCleartext(byte[] armoredMessage, OfficeOpenPgpKeySet keys,
        OfficeOpenPgpOptions? options = null);

    /// <summary>
    /// Opens an OpenPGP message ("BEGIN PGP MESSAGE" or binary): decrypts it with a supplied secret key when it is
    /// encrypted, checks integrity, and verifies any signatures over the literal data.
    /// </summary>
    OfficeOpenPgpMessageResult OpenMessage(byte[] message, OfficeOpenPgpKeySet keys, OfficeOpenPgpOptions? options = null);
}
