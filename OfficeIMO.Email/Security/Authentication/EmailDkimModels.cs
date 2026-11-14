namespace OfficeIMO.Email;

/// <summary>Result of evaluating one DKIM-style signature, using the RFC 8601 result vocabulary.</summary>
public enum EmailDkimResult {
    /// <summary>The signature, body hash, and key checks all succeeded.</summary>
    Pass = 0,

    /// <summary>The signature was well formed but the body hash or signature did not match, or it has expired.</summary>
    Fail = 1,

    /// <summary>The signature verified or could be verified, but local policy refuses it (weak algorithm, small key).</summary>
    Policy = 2,

    /// <summary>The signature could not be processed with the configured verifier.</summary>
    Neutral = 3,

    /// <summary>A transient key lookup failure prevented verification; retrying later may succeed.</summary>
    TempError = 4,

    /// <summary>The signature or its key record is malformed, missing, revoked, or inconsistent.</summary>
    PermError = 5
}

/// <summary>Specific reason for a non-passing signature result.</summary>
public enum EmailDkimFailureReason {
    /// <summary>No failure.</summary>
    None = 0,
    /// <summary>The tag list is malformed or contains duplicate tags.</summary>
    MalformedSignature,
    /// <summary>A required tag is missing or empty.</summary>
    MissingRequiredTag,
    /// <summary>A tag value is not valid for its tag.</summary>
    InvalidTagValue,
    /// <summary>The v= version is not supported.</summary>
    UnsupportedVersion,
    /// <summary>The a= algorithm is not supported.</summary>
    UnsupportedAlgorithm,
    /// <summary>The c= canonicalization is not supported.</summary>
    UnsupportedCanonicalization,
    /// <summary>The q= query methods do not include dns/txt.</summary>
    UnsupportedQueryMethod,
    /// <summary>rsa-sha1 signatures are refused per RFC 8301.</summary>
    WeakAlgorithm,
    /// <summary>The h= list does not include From, or includes a header the signature type forbids.</summary>
    InvalidSignedHeaders,
    /// <summary>The i= identity is not the d= domain or a subdomain of it.</summary>
    IdentityMismatch,
    /// <summary>The x= expiration time has passed.</summary>
    Expired,
    /// <summary>The x= expiration is earlier than the t= timestamp.</summary>
    InvalidTimestamps,
    /// <summary>The l= body length exceeds the canonical body.</summary>
    BodyLengthExceedsBody,
    /// <summary>The l= body length was refused by caller policy.</summary>
    BodyLengthLimitRefused,
    /// <summary>No key record was published for the selector.</summary>
    KeyNotFound,
    /// <summary>The key lookup failed temporarily.</summary>
    KeyLookupFailed,
    /// <summary>The key record is malformed, or more than one key record was returned.</summary>
    KeyRecordMalformed,
    /// <summary>The key record has an empty p= tag.</summary>
    KeyRevoked,
    /// <summary>The key record k= type does not match the signature algorithm.</summary>
    KeyTypeMismatch,
    /// <summary>The key record h= list does not allow the signature hash.</summary>
    KeyHashNotAllowed,
    /// <summary>The key record s= service types exclude email.</summary>
    KeyServiceNotAllowed,
    /// <summary>The key record t=s flag requires i= to use exactly the d= domain.</summary>
    KeyStrictIdentity,
    /// <summary>The public key is shorter than the configured minimum.</summary>
    KeyTooSmall,
    /// <summary>The public key data could not be decoded.</summary>
    KeyMalformed,
    /// <summary>The computed body hash does not match bh=.</summary>
    BodyHashMismatch,
    /// <summary>The signature does not match the signed header data.</summary>
    SignatureMismatch,
    /// <summary>The configured signature verifier does not implement the algorithm.</summary>
    VerifierUnavailable,
    /// <summary>A configured resource limit was reached.</summary>
    ResourceLimit
}

/// <summary>Kind of signature header evaluated.</summary>
public enum EmailDkimSignatureKind {
    /// <summary>A DKIM-Signature header field (RFC 6376).</summary>
    DkimSignature = 0,
    /// <summary>An ARC-Message-Signature header field (RFC 8617).</summary>
    ArcMessageSignature = 1,
    /// <summary>An ARC-Seal header field (RFC 8617).</summary>
    ArcSeal = 2
}

/// <summary>Typed evaluation of one DKIM-Signature, ARC-Message-Signature, or ARC-Seal header field.</summary>
public sealed class EmailDkimSignatureResult {
    internal EmailDkimSignatureResult(EmailDkimSignatureKind kind, int headerIndex) {
        Kind = kind;
        HeaderIndex = headerIndex;
    }

    /// <summary>Signature header kind.</summary>
    public EmailDkimSignatureKind Kind { get; }

    /// <summary>Zero-based position of the signature header among all header fields, top first.</summary>
    public int HeaderIndex { get; }

    /// <summary>Evaluation result.</summary>
    public EmailDkimResult Result { get; internal set; } = EmailDkimResult.PermError;

    /// <summary>Reason for a non-passing result.</summary>
    public EmailDkimFailureReason FailureReason { get; internal set; }

    /// <summary>Human-readable explanation. It never contains key material.</summary>
    public string Message { get; internal set; } = string.Empty;

    /// <summary>Signing domain (d=).</summary>
    public string? Domain { get; internal set; }

    /// <summary>Selector (s=).</summary>
    public string? Selector { get; internal set; }

    /// <summary>Agent or user identifier (i=) for DKIM-Signature; null for ARC headers.</summary>
    public string? Identity { get; internal set; }

    /// <summary>ARC instance (i=) for ARC headers; zero for DKIM-Signature.</summary>
    public int Instance { get; internal set; }

    /// <summary>Signature algorithm as written (a=).</summary>
    public string? Algorithm { get; internal set; }

    /// <summary>Header canonicalization.</summary>
    public EmailDkimCanonicalization HeaderCanonicalization { get; internal set; }

    /// <summary>Body canonicalization.</summary>
    public EmailDkimCanonicalization BodyCanonicalization { get; internal set; }

    /// <summary>Signed header names (h=) in signature order, including oversigned names.</summary>
    public IReadOnlyList<string> SignedHeaders { get; internal set; } = Array.Empty<string>();

    /// <summary>Body length limit (l=) when present.</summary>
    public long? BodyLengthLimit { get; internal set; }

    /// <summary>
    /// Canonical body bytes that follow the l= limit and are therefore not covered by the signature. Content there
    /// can be changed or appended by anyone without breaking the signature.
    /// </summary>
    public long UnsignedBodyBytes { get; internal set; }

    /// <summary>True when the body hash matched; null when it was not computed.</summary>
    public bool? BodyHashMatched { get; internal set; }

    /// <summary>Signature timestamp (t=).</summary>
    public DateTimeOffset? SignedAt { get; internal set; }

    /// <summary>Signature expiration (x=).</summary>
    public DateTimeOffset? ExpiresAt { get; internal set; }

    /// <summary>Public key size in bits when the key was decoded.</summary>
    public int KeySizeBits { get; internal set; }

    /// <summary>True when the key record carries the t=y testing flag.</summary>
    public bool IsTestingKey { get; internal set; }

    /// <summary>ARC-Seal chain validation tag (cv=) for ARC-Seal results.</summary>
    public string? ChainValidation { get; internal set; }

    /// <summary>Non-fatal observations such as a body length limit or a future timestamp.</summary>
    public IReadOnlyList<string> Warnings => _warnings;

    internal readonly List<string> _warnings = new List<string>();

    /// <summary>True only when <see cref="Result"/> is <see cref="EmailDkimResult.Pass"/>.</summary>
    public bool IsPass => Result == EmailDkimResult.Pass;

    internal EmailDkimSignatureResult Set(EmailDkimResult result, EmailDkimFailureReason reason, string message) {
        Result = result;
        FailureReason = reason;
        Message = message;
        return this;
    }
}

/// <summary>Result of verifying every DKIM-Signature header field in one message.</summary>
public sealed class EmailDkimVerificationResult {
    internal EmailDkimVerificationResult(IReadOnlyList<EmailDkimSignatureResult> signatures, IReadOnlyList<EmailDiagnostic> diagnostics) {
        Signatures = signatures;
        Diagnostics = diagnostics;
    }

    /// <summary>Per-signature results in header order, top first.</summary>
    public IReadOnlyList<EmailDkimSignatureResult> Signatures { get; }

    /// <summary>Message-level diagnostics such as missing raw source or normalized line endings.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }

    /// <summary>True when at least one signature passed.</summary>
    public bool HasPassingSignature => Signatures.Any(signature => signature.IsPass);
}

/// <summary>Resource limits and policy for DKIM and ARC verification.</summary>
public sealed class EmailAuthenticationOptions {
    /// <summary>Maximum raw message size processed. Larger messages are not verified.</summary>
    public long MaxMessageBytes { get; set; } = 64L * 1024 * 1024;

    /// <summary>Maximum number of header fields parsed.</summary>
    public int MaxHeaderFields { get; set; } = 2000;

    /// <summary>Maximum DKIM-Signature header fields evaluated; later ones are reported as not evaluated.</summary>
    public int MaxSignatures { get; set; } = 16;

    /// <summary>Maximum names in one h= list.</summary>
    public int MaxSignedHeaderFields { get; set; } = 128;

    /// <summary>Maximum distinct key lookups per verification.</summary>
    public int MaxKeyLookups { get; set; } = 32;

    /// <summary>Minimum accepted RSA key size. RFC 8301 requires at least 1024 bits.</summary>
    public int MinimumRsaKeyBits { get; set; } = 1024;

    /// <summary>
    /// When true, signatures carrying an l= body length limit are refused with <see cref="EmailDkimResult.Policy"/>.
    /// When false (the default), they are evaluated and reported with a warning and <see cref="EmailDkimSignatureResult.UnsignedBodyBytes"/>.
    /// </summary>
    public bool RejectBodyLengthLimit { get; set; }

    /// <summary>Verification time used for x= expiration. Defaults to the current UTC time.</summary>
    public DateTimeOffset? VerificationTime { get; set; }

    /// <summary>Allowed clock skew for t= timestamps in the future before a warning is recorded.</summary>
    public TimeSpan AllowedClockSkew { get; set; } = TimeSpan.FromMinutes(5);

    internal void Validate() {
        if (MaxMessageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(MaxMessageBytes));
        if (MaxHeaderFields <= 0) throw new ArgumentOutOfRangeException(nameof(MaxHeaderFields));
        if (MaxSignatures <= 0) throw new ArgumentOutOfRangeException(nameof(MaxSignatures));
        if (MaxSignedHeaderFields <= 0) throw new ArgumentOutOfRangeException(nameof(MaxSignedHeaderFields));
        if (MaxKeyLookups <= 0) throw new ArgumentOutOfRangeException(nameof(MaxKeyLookups));
        if (MinimumRsaKeyBits < 512) throw new ArgumentOutOfRangeException(nameof(MinimumRsaKeyBits));
        if (AllowedClockSkew < TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(AllowedClockSkew));
    }
}

/// <summary>Outcome of one DKIM key TXT lookup.</summary>
public enum EmailDkimKeyLookupStatus {
    /// <summary>One or more TXT records were returned.</summary>
    Found = 0,
    /// <summary>The name does not exist or has no TXT records.</summary>
    NotFound = 1,
    /// <summary>The lookup failed transiently (timeout, SERVFAIL).</summary>
    TemporaryFailure = 2
}

/// <summary>TXT records returned for one selector._domainkey query name.</summary>
public sealed class EmailDkimKeyLookupResult {
    private EmailDkimKeyLookupResult(EmailDkimKeyLookupStatus status, IReadOnlyList<string> records, string? message) {
        Status = status;
        Records = records;
        Message = message;
    }

    /// <summary>Lookup status.</summary>
    public EmailDkimKeyLookupStatus Status { get; }

    /// <summary>TXT records, each with its character-strings already concatenated.</summary>
    public IReadOnlyList<string> Records { get; }

    /// <summary>Optional resolver diagnostic.</summary>
    public string? Message { get; }

    /// <summary>Creates a result carrying TXT records.</summary>
    public static EmailDkimKeyLookupResult Found(IEnumerable<string> records) {
        ArgumentNullException.ThrowIfNull(records);
        string[] values = records.ToArray();
        return values.Length == 0
            ? NotFound()
            : new EmailDkimKeyLookupResult(EmailDkimKeyLookupStatus.Found, values, null);
    }

    /// <summary>Creates a result for a name without TXT records.</summary>
    public static EmailDkimKeyLookupResult NotFound(string? message = null) =>
        new EmailDkimKeyLookupResult(EmailDkimKeyLookupStatus.NotFound, Array.Empty<string>(), message);

    /// <summary>Creates a result for a transient lookup failure.</summary>
    public static EmailDkimKeyLookupResult TemporaryFailure(string? message = null) =>
        new EmailDkimKeyLookupResult(EmailDkimKeyLookupStatus.TemporaryFailure, Array.Empty<string>(), message);
}

/// <summary>
/// Supplies DKIM public key TXT records for <c>selector._domainkey.domain</c> names. OfficeIMO does not perform DNS
/// queries itself; hosts plug in their resolver, cache, or a fixed key set.
/// </summary>
public interface IEmailDkimKeyResolver {
    /// <summary>Returns the TXT records published at <paramref name="queryName"/>.</summary>
    Task<EmailDkimKeyLookupResult> ResolveTxtAsync(string queryName, CancellationToken cancellationToken);
}

/// <summary>Resolver backed by a fixed, case-insensitive map of query names to TXT records.</summary>
public sealed class EmailDkimStaticKeyResolver : IEmailDkimKeyResolver {
    private readonly Dictionary<string, List<string>> _records = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

    /// <summary>Adds a TXT record for <c>selector._domainkey.domain</c>.</summary>
    public EmailDkimStaticKeyResolver Add(string queryName, string record) {
        ArgumentException.ThrowIfNullOrWhiteSpace(queryName);
        ArgumentNullException.ThrowIfNull(record);
        string key = queryName.TrimEnd('.');
        if (!_records.TryGetValue(key, out List<string>? list)) {
            list = new List<string>();
            _records.Add(key, list);
        }

        list.Add(record);
        return this;
    }

    /// <summary>Adds a TXT record for a selector and domain.</summary>
    public EmailDkimStaticKeyResolver Add(string selector, string domain, string record) =>
        Add(selector + "._domainkey." + domain, record);

    /// <inheritdoc />
    public Task<EmailDkimKeyLookupResult> ResolveTxtAsync(string queryName, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(queryName != null && _records.TryGetValue(queryName.TrimEnd('.'), out List<string>? list)
            ? EmailDkimKeyLookupResult.Found(list)
            : EmailDkimKeyLookupResult.NotFound());
    }
}
