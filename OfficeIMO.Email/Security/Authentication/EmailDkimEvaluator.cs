using System.Security.Cryptography;
using OfficeIMO.Security;

namespace OfficeIMO.Email;

/// <summary>
/// Evaluates DKIM-Signature, ARC-Message-Signature, and ARC-Seal header fields over one parsed message. Every
/// ambiguous or malformed input produces a non-passing result; nothing here throws for hostile content.
/// </summary>
internal sealed class EmailDkimEvaluator {
    internal const int MaxArcInstance = 50;
    private const long MaxTimestampSeconds = 253402300799L; // 9999-12-31T23:59:59Z

    private readonly EmailAuthenticationMessage _message;
    private readonly IEmailDkimKeyResolver _resolver;
    private readonly IOfficePublicKeySignatureVerifier _verifier;
    private readonly EmailAuthenticationOptions _options;
    private readonly DateTimeOffset _now;
    private readonly Dictionary<string, EmailDkimKeyLookupResult> _lookups = new Dictionary<string, EmailDkimKeyLookupResult>(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<EmailDkimCanonicalization, byte[]> _canonicalBodies = new Dictionary<EmailDkimCanonicalization, byte[]>();
    private readonly Dictionary<string, byte[]> _bodyHashes = new Dictionary<string, byte[]>(StringComparer.Ordinal);

    internal EmailDkimEvaluator(EmailAuthenticationMessage message, IEmailDkimKeyResolver resolver,
        IOfficePublicKeySignatureVerifier verifier, EmailAuthenticationOptions options) {
        _message = message;
        _resolver = resolver;
        _verifier = verifier;
        _options = options;
        _now = options.VerificationTime ?? DateTimeOffset.UtcNow;
    }

    internal EmailAuthenticationMessage Message => _message;

    /// <summary>Parses the ARC instance (i=) of an ARC-Seal or ARC-Message-Signature; zero when invalid.</summary>
    internal static int ReadArcInstance(EmailDkimTagList? tags) =>
        tags != null && TryParseInstance(tags["i"], out int instance) ? instance : 0;

    internal static bool TryParseInstance(string? value, out int instance) {
        instance = 0;
        if (string.IsNullOrEmpty(value) || value!.Length > 2 || !value.All(IsDigit)) return false;
        instance = int.Parse(value, NumberStyles.None, CultureInfo.InvariantCulture);
        return instance >= 1 && instance <= MaxArcInstance;
    }

    /// <summary>Evaluates one signature header.</summary>
    /// <param name="fieldIndex">Index of the signature header field.</param>
    /// <param name="kind">Signature type.</param>
    /// <param name="sealedFields">For ARC-Seal only: the AAR, AMS, AS fields of instances 1..i in signing order.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    internal async Task<EmailDkimSignatureResult> EvaluateAsync(int fieldIndex, EmailDkimSignatureKind kind,
        IReadOnlyList<int>? sealedFields, CancellationToken cancellationToken) {
        var result = new EmailDkimSignatureResult(kind, fieldIndex);
        EmailAuthenticationMessage.Field field = _message.Fields[fieldIndex];
        EmailDkimTagList? tags = EmailDkimTagList.Parse(field.Value);
        if (tags == null) {
            return result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.MalformedSignature,
                "The signature tag list is malformed or contains duplicate tags.");
        }

        Signature? signature = ReadSignature(tags, kind, result);
        if (signature == null) return result;

        if (kind != EmailDkimSignatureKind.ArcSeal) {
            EmailDkimSignatureResult? bodyFailure = CheckBody(signature, result);
            if (bodyFailure != null) return bodyFailure;
        }

        if (result.Result == EmailDkimResult.Policy) return result;

        if (!_verifier.Supports(signature.Algorithm)) {
            return result.Set(EmailDkimResult.Neutral, EmailDkimFailureReason.VerifierUnavailable,
                "The configured signature verifier '" + _verifier.Name + "' does not implement " + result.Algorithm + ".");
        }

        EmailDkimKeyRecord? key = await ResolveKeyAsync(signature, result, cancellationToken).ConfigureAwait(false);
        if (key == null) return result;

        byte[] signedData = BuildSignedHeaderData(field, tags, kind, sealedFields, result);
        byte[] verifiedData = signature.Algorithm == OfficePublicKeySignatureAlgorithm.Ed25519 ? Sha256(signedData) : signedData;
        OfficePublicKeySignatureCheck check;
        try {
            check = _verifier.Verify(signature.Algorithm, key.PublicKey, verifiedData, signature.SignatureValue);
        } catch (Exception ex) when (ex is ArgumentException or InvalidOperationException or CryptographicException) {
            return result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.KeyMalformed,
                "The signature verifier rejected the key or signature.");
        }

        result.KeySizeBits = check.KeySizeBits;
        switch (check.Status) {
            case OfficePublicKeySignatureStatus.MalformedKey:
                return result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.KeyMalformed,
                    "The published public key could not be decoded for " + result.Algorithm + ".");
            case OfficePublicKeySignatureStatus.UnsupportedAlgorithm:
                return result.Set(EmailDkimResult.Neutral, EmailDkimFailureReason.VerifierUnavailable,
                    "The configured signature verifier does not implement " + result.Algorithm + ".");
        }

        if (signature.Algorithm != OfficePublicKeySignatureAlgorithm.Ed25519 && check.KeySizeBits < _options.MinimumRsaKeyBits) {
            return result.Set(EmailDkimResult.Policy, EmailDkimFailureReason.KeyTooSmall,
                "The RSA key has " + check.KeySizeBits.ToString(CultureInfo.InvariantCulture) +
                " bits; at least " + _options.MinimumRsaKeyBits.ToString(CultureInfo.InvariantCulture) + " are required.");
        }

        if (!check.IsValid) {
            return result.Set(EmailDkimResult.Fail, EmailDkimFailureReason.SignatureMismatch,
                "The signature does not match the signed header fields.");
        }

        if (signature.ExpiresAt is long expires && _now.ToUnixTimeSeconds() > expires) {
            return result.Set(EmailDkimResult.Fail, EmailDkimFailureReason.Expired, "The signature has expired (x=).");
        }

        return result.Set(EmailDkimResult.Pass, EmailDkimFailureReason.None, "The signature verified.");
    }

    private Signature? ReadSignature(EmailDkimTagList tags, EmailDkimSignatureKind kind, EmailDkimSignatureResult result) {
        string[] required = kind switch {
            EmailDkimSignatureKind.DkimSignature => new[] { "v", "a", "b", "bh", "d", "h", "s" },
            EmailDkimSignatureKind.ArcMessageSignature => new[] { "i", "a", "b", "bh", "d", "h", "s" },
            _ => new[] { "i", "a", "b", "d", "s", "cv" }
        };
        foreach (string name in required) {
            string? value = tags[name];
            if (value == null || (value.Length == 0 && name != "h")) {
                result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.MissingRequiredTag,
                    "The required " + name + "= tag is missing or empty.");
                return null;
            }
        }

        result.Domain = tags["d"];
        result.Selector = tags["s"];
        result.Algorithm = tags["a"];

        if (kind == EmailDkimSignatureKind.DkimSignature && tags["v"] != "1") {
            result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.UnsupportedVersion, "Only DKIM version 1 is supported.");
            return null;
        }

        if (kind != EmailDkimSignatureKind.DkimSignature) {
            if (!TryParseInstance(tags["i"], out int instance)) {
                result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.InvalidTagValue, "The ARC instance (i=) must be 1 to 50.");
                return null;
            }

            result.Instance = instance;
        }

        if (!IsDomainName(result.Domain!) || !IsSelector(result.Selector!)) {
            result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.InvalidTagValue,
                "The d= domain or s= selector is not a valid domain name.");
            return null;
        }

        OfficePublicKeySignatureAlgorithm algorithm;
        switch (result.Algorithm) {
            case "rsa-sha256":
                algorithm = OfficePublicKeySignatureAlgorithm.RsaPkcs1Sha256;
                break;
            case "ed25519-sha256":
                algorithm = OfficePublicKeySignatureAlgorithm.Ed25519;
                break;
            case "rsa-sha1":
                result.Set(EmailDkimResult.Policy, EmailDkimFailureReason.WeakAlgorithm,
                    "rsa-sha1 signatures are not accepted (RFC 8301).");
                return null;
            default:
                result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.UnsupportedAlgorithm,
                    "The signature algorithm is not supported.");
                return null;
        }

        if (kind == EmailDkimSignatureKind.ArcSeal) {
            result.HeaderCanonicalization = EmailDkimCanonicalization.Relaxed;
            result.BodyCanonicalization = EmailDkimCanonicalization.Relaxed;
            if (tags.Contains("h")) {
                result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.InvalidSignedHeaders, "ARC-Seal must not carry an h= tag.");
                return null;
            }

            string cv = tags["cv"]!;
            if (cv != "none" && cv != "pass" && cv != "fail") {
                result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.InvalidTagValue, "The ARC-Seal cv= value is not valid.");
                return null;
            }

            result.ChainValidation = cv;
        } else {
            if (!TryParseCanonicalization(tags["c"], out EmailDkimCanonicalization header, out EmailDkimCanonicalization body)) {
                result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.UnsupportedCanonicalization,
                    "The c= canonicalization is not supported.");
                return null;
            }

            result.HeaderCanonicalization = header;
            result.BodyCanonicalization = body;
            string[] signedHeaders = tags["h"]!.Split(':').Select(name => EmailDkimTagList.RemoveWhitespace(name)).ToArray();
            if (signedHeaders.Length == 1 && signedHeaders[0].Length == 0 && kind == EmailDkimSignatureKind.ArcMessageSignature) {
                signedHeaders = Array.Empty<string>();
            }

            if (signedHeaders.Length > _options.MaxSignedHeaderFields) {
                result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.ResourceLimit, "The h= list exceeds the configured limit.");
                return null;
            }

            if (signedHeaders.Any(name => !IsFieldName(name))) {
                result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.InvalidSignedHeaders, "The h= list contains an invalid header name.");
                return null;
            }

            result.SignedHeaders = signedHeaders;
            if (kind == EmailDkimSignatureKind.DkimSignature &&
                !signedHeaders.Any(name => name.Equals("From", StringComparison.OrdinalIgnoreCase))) {
                result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.InvalidSignedHeaders, "The h= list does not include From.");
                return null;
            }

            if (kind == EmailDkimSignatureKind.ArcMessageSignature &&
                signedHeaders.Any(name => name.Equals("ARC-Seal", StringComparison.OrdinalIgnoreCase))) {
                result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.InvalidSignedHeaders, "ARC-Message-Signature must not sign ARC-Seal.");
                return null;
            }
        }

        if (kind == EmailDkimSignatureKind.DkimSignature) {
            if (!ValidateDkimOnlyTags(tags, result)) return null;
        }

        long? signedAt = null;
        if (tags["t"] is string timestamp) {
            if (!TryParseSeconds(timestamp, out long seconds)) {
                result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.InvalidTagValue, "The t= timestamp is not valid.");
                return null;
            }

            signedAt = seconds;
            result.SignedAt = DateTimeOffset.FromUnixTimeSeconds(seconds);
            if (result.SignedAt > _now + _options.AllowedClockSkew) result._warnings.Add("The t= timestamp is in the future.");
        }

        long? expiresAt = null;
        if (kind != EmailDkimSignatureKind.ArcSeal && tags["x"] is string expiration) {
            if (!TryParseSeconds(expiration, out long seconds)) {
                result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.InvalidTagValue, "The x= expiration is not valid.");
                return null;
            }

            if (signedAt is long start && seconds <= start) {
                result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.InvalidTimestamps, "The x= expiration is not after t=.");
                return null;
            }

            expiresAt = seconds;
            result.ExpiresAt = DateTimeOffset.FromUnixTimeSeconds(seconds);
        }

        if (!TryDecodeBase64(tags["b"]!, out byte[] signatureValue)) {
            result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.InvalidTagValue, "The b= value is not valid base64.");
            return null;
        }

        byte[] bodyHash = Array.Empty<byte>();
        if (kind != EmailDkimSignatureKind.ArcSeal && !TryDecodeBase64(tags["bh"]!, out bodyHash)) {
            result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.InvalidTagValue, "The bh= value is not valid base64.");
            return null;
        }

        long? length = null;
        if (kind != EmailDkimSignatureKind.ArcSeal && tags["l"] is string bodyLength) {
            if (bodyLength.Length == 0 || bodyLength.Length > 18 || !bodyLength.All(IsDigit)) {
                result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.InvalidTagValue, "The l= body length is not valid.");
                return null;
            }

            length = long.Parse(bodyLength, NumberStyles.None, CultureInfo.InvariantCulture);
            result.BodyLengthLimit = length;
        }

        return new Signature(algorithm, signatureValue, bodyHash, length, expiresAt);
    }

    private static bool ValidateDkimOnlyTags(EmailDkimTagList tags, EmailDkimSignatureResult result) {
        string domain = tags["d"]!;
        string identity = tags["i"] ?? "@" + domain;
        int at = identity.LastIndexOf('@');
        string identityDomain = at < 0 ? string.Empty : identity.Substring(at + 1);
        if (at < 0 || !IsDomainName(identityDomain) || !IsSameOrSubdomain(identityDomain, domain)) {
            result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.IdentityMismatch,
                "The i= identity is not in the d= domain or one of its subdomains.");
            return false;
        }

        result.Identity = identity;
        if (tags["q"] is string query &&
            !EmailDkimKeyRecord.SplitColonList(query).Any(method => method.Equals("dns/txt", StringComparison.OrdinalIgnoreCase))) {
            result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.UnsupportedQueryMethod, "The q= tag does not allow dns/txt.");
            return false;
        }

        return true;
    }

    private EmailDkimSignatureResult? CheckBody(Signature signature, EmailDkimSignatureResult result) {
        string cacheKey = ((int)result.BodyCanonicalization).ToString(CultureInfo.InvariantCulture) + ":" +
            (signature.BodyLength?.ToString(CultureInfo.InvariantCulture) ?? "-");
        if (!_canonicalBodies.TryGetValue(result.BodyCanonicalization, out byte[]? canonical)) {
            canonical = EmailDkimCanonicalizer.CanonicalizeBody(_message.Body, result.BodyCanonicalization);
            _canonicalBodies[result.BodyCanonicalization] = canonical;
        }

        int hashedLength = canonical.Length;
        if (signature.BodyLength is long limit) {
            if (limit > canonical.LongLength) {
                return result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.BodyLengthExceedsBody,
                    "The l= body length is larger than the canonical body.");
            }

            hashedLength = (int)limit;
            result.UnsignedBodyBytes = canonical.LongLength - limit;
            result._warnings.Add(result.UnsignedBodyBytes > 0
                ? "The signature covers only the first " + limit.ToString(CultureInfo.InvariantCulture) + " canonical body bytes; " +
                  result.UnsignedBodyBytes.ToString(CultureInfo.InvariantCulture) + " bytes after them are not signed and may have been added by anyone."
                : "The signature uses an l= body length limit; content appended later would not break it.");
            if (_options.RejectBodyLengthLimit) {
                result.Set(EmailDkimResult.Policy, EmailDkimFailureReason.BodyLengthLimitRefused,
                    "Signatures with an l= body length limit are refused by the verification options.");
            }
        }

        if (!_bodyHashes.TryGetValue(cacheKey, out byte[]? hash)) {
            using SHA256 sha256 = SHA256.Create();
            hash = sha256.ComputeHash(canonical, 0, hashedLength);
            _bodyHashes[cacheKey] = hash;
        }

        result.BodyHashMatched = FixedTimeEquals(hash, signature.BodyHash);
        if (result.BodyHashMatched != true) {
            return result.Set(EmailDkimResult.Fail, EmailDkimFailureReason.BodyHashMismatch,
                "The body hash does not match bh=; the body was changed after signing.");
        }

        return null;
    }

    private async Task<EmailDkimKeyRecord?> ResolveKeyAsync(Signature signature, EmailDkimSignatureResult result,
        CancellationToken cancellationToken) {
        string queryName = result.Selector + "._domainkey." + result.Domain;
        if (!_lookups.TryGetValue(queryName, out EmailDkimKeyLookupResult? lookup)) {
            if (_lookups.Count >= _options.MaxKeyLookups) {
                result.Set(EmailDkimResult.Neutral, EmailDkimFailureReason.ResourceLimit, "The key lookup limit was reached.");
                return null;
            }

            try {
                lookup = await _resolver.ResolveTxtAsync(queryName, cancellationToken).ConfigureAwait(false)
                    ?? EmailDkimKeyLookupResult.TemporaryFailure("The resolver returned no result.");
            } catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested) {
                throw;
            } catch (Exception ex) when (ex is not OutOfMemoryException) {
                lookup = EmailDkimKeyLookupResult.TemporaryFailure("The resolver failed: " + ex.GetType().Name + ".");
            }

            _lookups[queryName] = lookup;
        }

        if (lookup.Status == EmailDkimKeyLookupStatus.TemporaryFailure) {
            result.Set(EmailDkimResult.TempError, EmailDkimFailureReason.KeyLookupFailed,
                "The key lookup for " + queryName + " failed temporarily.");
            return null;
        }

        if (lookup.Status == EmailDkimKeyLookupStatus.NotFound || lookup.Records.Count == 0) {
            result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.KeyNotFound, "No key record is published at " + queryName + ".");
            return null;
        }

        if (lookup.Records.Count > 1) {
            result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.KeyRecordMalformed,
                "More than one TXT record is published at " + queryName + "; the key is ambiguous.");
            return null;
        }

        EmailDkimKeyRecord? key = EmailDkimKeyRecord.Parse(lookup.Records[0] ?? string.Empty, out string error);
        if (key == null) {
            result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.KeyRecordMalformed, error);
            return null;
        }

        result.IsTestingKey = key.Testing;
        if (key.Testing) result._warnings.Add("The key record is flagged as testing (t=y).");
        if (key.IsRevoked) {
            result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.KeyRevoked, "The key record has been revoked (empty p=).");
            return null;
        }

        string expectedType = signature.Algorithm == OfficePublicKeySignatureAlgorithm.Ed25519 ? "ed25519" : "rsa";
        if (!key.KeyType.Equals(expectedType, StringComparison.OrdinalIgnoreCase)) {
            result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.KeyTypeMismatch,
                "The key record type does not match the signature algorithm.");
            return null;
        }

        if (key.Hashes != null && !key.Hashes.Any(hash => hash.Equals("sha256", StringComparison.OrdinalIgnoreCase))) {
            result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.KeyHashNotAllowed, "The key record does not allow sha256.");
            return null;
        }

        if (!key.Services.Any(service => service == "*" || service.Equals("email", StringComparison.OrdinalIgnoreCase))) {
            result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.KeyServiceNotAllowed, "The key record is not for email.");
            return null;
        }

        if (key.Strict && result.Kind == EmailDkimSignatureKind.DkimSignature && result.Identity != null) {
            string identityDomain = result.Identity.Substring(result.Identity.LastIndexOf('@') + 1);
            if (!identityDomain.Equals(result.Domain, StringComparison.OrdinalIgnoreCase)) {
                result.Set(EmailDkimResult.PermError, EmailDkimFailureReason.KeyStrictIdentity,
                    "The key record (t=s) requires i= to use exactly the d= domain.");
                return null;
            }
        }

        return key;
    }

    private byte[] BuildSignedHeaderData(EmailAuthenticationMessage.Field signatureField, EmailDkimTagList tags,
        EmailDkimSignatureKind kind, IReadOnlyList<int>? sealedFields, EmailDkimSignatureResult result) {
        var builder = new StringBuilder();
        EmailDkimCanonicalization canonicalization = kind == EmailDkimSignatureKind.ArcSeal
            ? EmailDkimCanonicalization.Relaxed
            : HeaderCanonicalizationOf(tags);
        if (kind == EmailDkimSignatureKind.ArcSeal) {
            if (sealedFields != null) {
                for (int index = 0; index < sealedFields.Count - 1; index++) {
                    builder.Append(EmailDkimCanonicalizer.CanonicalizeHeader(_message.Fields[sealedFields[index]].Raw, canonicalization));
                }
            }
        } else {
            var used = new HashSet<int>();
            foreach (string name in tags["h"]!.Split(':').Select(EmailDkimTagList.RemoveWhitespace)) {
                if (name.Length == 0) continue;
                for (int index = _message.Fields.Count - 1; index >= 0; index--) {
                    if (used.Contains(index) || !_message.Fields[index].Is(name)) continue;
                    used.Add(index);
                    builder.Append(EmailDkimCanonicalizer.CanonicalizeHeader(_message.Fields[index].Raw, canonicalization));
                    break;
                }
            }

            foreach (string name in result.SignedHeaders.Distinct(StringComparer.OrdinalIgnoreCase)) {
                int unsigned = 0;
                for (int index = 0; index < _message.Fields.Count; index++) {
                    if (!used.Contains(index) && index != result.HeaderIndex && _message.Fields[index].Is(name)) unsigned++;
                }

                if (unsigned > 0) {
                    result._warnings.Add(unsigned.ToString(CultureInfo.InvariantCulture) + " instance(s) of " + name +
                        " are not covered by the signature; they could have been added after signing.");
                }
            }
        }

        string withoutSignature = RemoveSignatureValue(signatureField, tags);
        string canonicalSignature = EmailDkimCanonicalizer.CanonicalizeHeader(withoutSignature, canonicalization);
        if (canonicalSignature.EndsWith("\r\n", StringComparison.Ordinal)) {
            canonicalSignature = canonicalSignature.Substring(0, canonicalSignature.Length - 2);
        }

        builder.Append(canonicalSignature);
        return EmailAuthenticationMessage.ToBytes(builder.ToString());
    }

    private static EmailDkimCanonicalization HeaderCanonicalizationOf(EmailDkimTagList tags) =>
        TryParseCanonicalization(tags["c"], out EmailDkimCanonicalization header, out _) ? header : EmailDkimCanonicalization.Simple;

    private static string RemoveSignatureValue(EmailAuthenticationMessage.Field field, EmailDkimTagList tags) {
        string raw = field.Raw;
        int colon = raw.IndexOf(':');
        if (!tags.TryGetSpan("b", out int start, out int length)) return raw;
        int offset = colon + 1 + start;
        return raw.Substring(0, offset) + raw.Substring(offset + length);
    }

    private static bool TryParseCanonicalization(string? value, out EmailDkimCanonicalization header, out EmailDkimCanonicalization body) {
        header = EmailDkimCanonicalization.Simple;
        body = EmailDkimCanonicalization.Simple;
        if (value == null) return true;
        string[] parts = value.Split('/');
        if (parts.Length > 2 || !TryParseCanonicalizationName(parts[0], out header)) return false;
        return parts.Length == 1 || TryParseCanonicalizationName(parts[1], out body);
    }

    private static bool TryParseCanonicalizationName(string value, out EmailDkimCanonicalization canonicalization) {
        canonicalization = EmailDkimCanonicalization.Simple;
        if (value == "simple") return true;
        if (value != "relaxed") return false;
        canonicalization = EmailDkimCanonicalization.Relaxed;
        return true;
    }

    private static bool TryParseSeconds(string value, out long seconds) {
        seconds = 0;
        if (value.Length == 0 || value.Length > 12 || !value.All(IsDigit)) return false;
        seconds = long.Parse(value, NumberStyles.None, CultureInfo.InvariantCulture);
        return seconds <= MaxTimestampSeconds;
    }

    private static bool TryDecodeBase64(string value, out byte[] bytes) {
        bytes = Array.Empty<byte>();
        string compact = EmailDkimTagList.RemoveWhitespace(value);
        if (compact.Length == 0) return false;
        try {
            bytes = Convert.FromBase64String(compact);
            return bytes.Length > 0;
        } catch (FormatException) {
            return false;
        }
    }

    internal static bool IsDomainName(string value) {
        if (value.Length == 0 || value.Length > 253) return false;
        string trimmed = value.EndsWith(".", StringComparison.Ordinal) ? value.Substring(0, value.Length - 1) : value;
        string[] labels = trimmed.Split('.');
        return labels.All(label => IsLabel(label, allowUnderscore: false));
    }

    private static bool IsSelector(string value) =>
        value.Length > 0 && value.Length <= 253 && value.Split('.').All(label => IsLabel(label, allowUnderscore: true));

    private static bool IsLabel(string label, bool allowUnderscore) {
        if (label.Length == 0 || label.Length > 63 || label[0] == '-' || label[label.Length - 1] == '-') return false;
        foreach (char character in label) {
            bool valid = (character >= 'a' && character <= 'z') || (character >= 'A' && character <= 'Z') ||
                IsDigit(character) || character == '-' || (allowUnderscore && character == '_');
            if (!valid) return false;
        }

        return true;
    }

    private static bool IsFieldName(string name) {
        if (name.Length == 0) return false;
        foreach (char character in name) {
            if (character < 0x21 || character > 0x7E || character == ':') return false;
        }

        return true;
    }

    private static bool IsSameOrSubdomain(string candidate, string domain) {
        string left = candidate.TrimEnd('.');
        string right = domain.TrimEnd('.');
        return left.Equals(right, StringComparison.OrdinalIgnoreCase) ||
            left.EndsWith("." + right, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsDigit(char character) => character >= '0' && character <= '9';

    private static byte[] Sha256(byte[] data) {
        using SHA256 sha256 = SHA256.Create();
        return sha256.ComputeHash(data);
    }

    private static bool FixedTimeEquals(byte[] left, byte[] right) {
        if (left.Length != right.Length) return false;
        int difference = 0;
        for (int index = 0; index < left.Length; index++) difference |= left[index] ^ right[index];
        return difference == 0;
    }

    private sealed class Signature {
        internal Signature(OfficePublicKeySignatureAlgorithm algorithm, byte[] signatureValue, byte[] bodyHash, long? bodyLength,
            long? expiresAt) {
            Algorithm = algorithm;
            SignatureValue = signatureValue;
            BodyHash = bodyHash;
            BodyLength = bodyLength;
            ExpiresAt = expiresAt;
        }

        internal OfficePublicKeySignatureAlgorithm Algorithm { get; }

        internal byte[] SignatureValue { get; }

        internal byte[] BodyHash { get; }

        internal long? BodyLength { get; }

        internal long? ExpiresAt { get; }
    }
}
