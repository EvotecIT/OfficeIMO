using OfficeIMO.Security;

namespace OfficeIMO.Email;

/// <summary>ARC chain validation status (RFC 8617 section 4.4).</summary>
public enum EmailArcChainStatus {
    /// <summary>The message carries no ARC sets.</summary>
    None = 0,

    /// <summary>Every ARC-Seal and the newest ARC-Message-Signature verified and the chain is well formed.</summary>
    Pass = 1,

    /// <summary>The chain is malformed, marked failed, or a required signature did not verify.</summary>
    Fail = 2,

    /// <summary>
    /// The chain could not be evaluated: a key lookup failed temporarily, the verifier lacks an algorithm, or a
    /// resource limit was reached. This is never a pass.
    /// </summary>
    NotVerified = 3
}

/// <summary>Reason for a non-passing ARC chain status.</summary>
public enum EmailArcFailureReason {
    /// <summary>No failure.</summary>
    None = 0,
    /// <summary>An ARC header field could not be parsed or has no valid instance.</summary>
    MalformedHeader,
    /// <summary>Instances are missing, duplicated, or not contiguous from 1.</summary>
    IncompleteSet,
    /// <summary>More than 50 ARC sets are present.</summary>
    TooManyInstances,
    /// <summary>The newest ARC-Seal records cv=fail.</summary>
    ChainMarkedFailed,
    /// <summary>An ARC-Seal cv= value is inconsistent with its instance.</summary>
    ChainValidationMismatch,
    /// <summary>The newest ARC-Message-Signature did not verify.</summary>
    MessageSignatureFailed,
    /// <summary>An ARC-Seal did not verify.</summary>
    SealFailed,
    /// <summary>A component could not be evaluated (temporary key failure, verifier or resource limits).</summary>
    NotEvaluated
}

/// <summary>One ARC set: the ARC-Authentication-Results, ARC-Message-Signature, and ARC-Seal of one instance.</summary>
public sealed class EmailArcSet {
    internal EmailArcSet(int instance) => Instance = instance;

    /// <summary>ARC instance number.</summary>
    public int Instance { get; }

    /// <summary>ARC-Authentication-Results payload after the i= prefix, as written.</summary>
    public string? AuthenticationResults { get; internal set; }

    /// <summary>ARC-Message-Signature evaluation, when it was evaluated.</summary>
    public EmailDkimSignatureResult? MessageSignature { get; internal set; }

    /// <summary>ARC-Seal evaluation, when it was evaluated.</summary>
    public EmailDkimSignatureResult? Seal { get; internal set; }
}

/// <summary>Result of ARC chain validation for one message.</summary>
public sealed class EmailArcVerificationResult {
    internal EmailArcVerificationResult(EmailArcChainStatus status, EmailArcFailureReason reason, string message,
        IReadOnlyList<EmailArcSet> sets, IReadOnlyList<EmailDiagnostic> diagnostics) {
        Status = status;
        FailureReason = reason;
        Message = message;
        Sets = sets;
        Diagnostics = diagnostics;
    }

    /// <summary>Chain validation status (the cv value a validator would record).</summary>
    public EmailArcChainStatus Status { get; }

    /// <summary>Reason for a non-passing status.</summary>
    public EmailArcFailureReason FailureReason { get; }

    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }

    /// <summary>ARC sets ordered by instance, oldest first. Empty when the structure could not be established.</summary>
    public IReadOnlyList<EmailArcSet> Sets { get; }

    /// <summary>Message-level diagnostics.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }

    /// <summary>
    /// Lowest instance whose ARC-Message-Signature and every newer one verified, or zero when the chain did not pass.
    /// Older ARC-Message-Signatures are expected to break as intermediaries modify the message.
    /// </summary>
    public int OldestPassingInstance { get; internal set; }
}

/// <summary>ARC (RFC 8617) chain validation over the exact received message bytes.</summary>
public static class EmailArc {
    /// <summary>Validates the ARC chain of a raw RFC 5322 message.</summary>
    public static async Task<EmailArcVerificationResult> VerifyAsync(
        byte[] rawMessage,
        IEmailDkimKeyResolver keyResolver,
        IOfficePublicKeySignatureVerifier signatureVerifier,
        EmailAuthenticationOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(rawMessage);
        ArgumentNullException.ThrowIfNull(keyResolver);
        ArgumentNullException.ThrowIfNull(signatureVerifier);
        options ??= new EmailAuthenticationOptions();
        options.Validate();
        var diagnostics = new List<EmailDiagnostic>();
        EmailAuthenticationMessage? message = EmailAuthenticationParsing.Parse(rawMessage, options, diagnostics, "arc");
        if (message == null) {
            return new EmailArcVerificationResult(EmailArcChainStatus.NotVerified, EmailArcFailureReason.NotEvaluated,
                "The message was not evaluated.", Array.Empty<EmailArcSet>(), diagnostics);
        }

        var seals = new List<(int Instance, int Index)>();
        var messageSignatures = new List<(int Instance, int Index)>();
        var results = new List<(int Instance, int Index, string? Payload)>();
        int arcFields = 0;
        for (int index = 0; index < message.Fields.Count; index++) {
            EmailAuthenticationMessage.Field field = message.Fields[index];
            if (field.Is("ARC-Seal")) {
                seals.Add((EmailDkimEvaluator.ReadArcInstance(EmailDkimTagList.Parse(field.Value)), index));
            } else if (field.Is("ARC-Message-Signature")) {
                messageSignatures.Add((EmailDkimEvaluator.ReadArcInstance(EmailDkimTagList.Parse(field.Value)), index));
            } else if (field.Is("ARC-Authentication-Results")) {
                int instance = ReadAuthenticationResultsInstance(field.Value, out string? payload);
                results.Add((instance, index, payload));
            } else {
                continue;
            }

            arcFields++;
        }

        if (arcFields == 0) {
            return new EmailArcVerificationResult(EmailArcChainStatus.None, EmailArcFailureReason.None,
                "The message carries no ARC sets.", Array.Empty<EmailArcSet>(), diagnostics);
        }

        if (seals.Any(item => item.Instance == 0) || messageSignatures.Any(item => item.Instance == 0) ||
            results.Any(item => item.Instance == 0)) {
            return Failed(EmailArcFailureReason.MalformedHeader, "An ARC header field is malformed or has no valid i= instance.");
        }

        int highest = Math.Max(seals.Count == 0 ? 0 : seals.Max(item => item.Instance),
            Math.Max(messageSignatures.Count == 0 ? 0 : messageSignatures.Max(item => item.Instance),
                results.Count == 0 ? 0 : results.Max(item => item.Instance)));
        if (highest > EmailDkimEvaluator.MaxArcInstance) {
            return Failed(EmailArcFailureReason.TooManyInstances, "More than 50 ARC sets are present.");
        }

        if (!IsComplete(seals.Select(item => item.Instance), highest) ||
            !IsComplete(messageSignatures.Select(item => item.Instance), highest) ||
            !IsComplete(results.Select(item => item.Instance), highest)) {
            return Failed(EmailArcFailureReason.IncompleteSet,
                "Each instance from 1 to " + highest.ToString(CultureInfo.InvariantCulture) +
                " must have exactly one ARC-Seal, ARC-Message-Signature, and ARC-Authentication-Results.");
        }

        var sets = new EmailArcSet[highest];
        for (int instance = 1; instance <= highest; instance++) {
            sets[instance - 1] = new EmailArcSet(instance) {
                AuthenticationResults = results.Single(item => item.Instance == instance).Payload
            };
        }

        int SealIndex(int instance) => seals.Single(item => item.Instance == instance).Index;
        int SignatureIndex(int instance) => messageSignatures.Single(item => item.Instance == instance).Index;
        int ResultIndex(int instance) => results.Single(item => item.Instance == instance).Index;

        for (int instance = 1; instance <= highest; instance++) {
            string? cv = EmailDkimTagList.Parse(message.Fields[SealIndex(instance)].Value)?["cv"];
            if (instance == highest && cv == "fail") {
                return Failed(EmailArcFailureReason.ChainMarkedFailed, "The newest ARC-Seal records that the chain had already failed (cv=fail).", sets);
            }

            string expected = instance == 1 ? "none" : "pass";
            if (cv != expected) {
                return Failed(EmailArcFailureReason.ChainValidationMismatch,
                    "ARC-Seal i=" + instance.ToString(CultureInfo.InvariantCulture) + " must carry cv=" + expected + ".", sets);
            }
        }

        var evaluator = new EmailDkimEvaluator(message, keyResolver, signatureVerifier, options);
        EmailDkimSignatureResult newest = await evaluator.EvaluateAsync(SignatureIndex(highest),
            EmailDkimSignatureKind.ArcMessageSignature, null, cancellationToken).ConfigureAwait(false);
        sets[highest - 1].MessageSignature = newest;
        if (!newest.IsPass) {
            return NotPassing(newest, EmailArcFailureReason.MessageSignatureFailed,
                "The newest ARC-Message-Signature did not verify: " + newest.Message, sets);
        }

        for (int instance = highest; instance >= 1; instance--) {
            var sealedFields = new List<int>(instance * 3);
            for (int covered = 1; covered <= instance; covered++) {
                sealedFields.Add(ResultIndex(covered));
                sealedFields.Add(SignatureIndex(covered));
                sealedFields.Add(SealIndex(covered));
            }

            EmailDkimSignatureResult seal = await evaluator.EvaluateAsync(SealIndex(instance), EmailDkimSignatureKind.ArcSeal,
                sealedFields, cancellationToken).ConfigureAwait(false);
            sets[instance - 1].Seal = seal;
            if (!seal.IsPass) {
                return NotPassing(seal, EmailArcFailureReason.SealFailed,
                    "ARC-Seal i=" + instance.ToString(CultureInfo.InvariantCulture) + " did not verify: " + seal.Message, sets);
            }
        }

        int oldestPassing = highest;
        for (int instance = highest - 1; instance >= 1; instance--) {
            EmailDkimSignatureResult older = await evaluator.EvaluateAsync(SignatureIndex(instance),
                EmailDkimSignatureKind.ArcMessageSignature, null, cancellationToken).ConfigureAwait(false);
            sets[instance - 1].MessageSignature = older;
            if (!older.IsPass || oldestPassing != instance + 1) continue;
            oldestPassing = instance;
        }

        return new EmailArcVerificationResult(EmailArcChainStatus.Pass, EmailArcFailureReason.None,
            "The ARC chain verified.", sets, diagnostics) { OldestPassingInstance = oldestPassing };

        EmailArcVerificationResult Failed(EmailArcFailureReason reason, string text, IReadOnlyList<EmailArcSet>? known = null) =>
            new EmailArcVerificationResult(EmailArcChainStatus.Fail, reason, text, known ?? Array.Empty<EmailArcSet>(), diagnostics);

        EmailArcVerificationResult NotPassing(EmailDkimSignatureResult component, EmailArcFailureReason reason, string text,
            IReadOnlyList<EmailArcSet> known) {
            bool indeterminate = component.Result == EmailDkimResult.TempError || component.Result == EmailDkimResult.Neutral;
            return indeterminate
                ? new EmailArcVerificationResult(EmailArcChainStatus.NotVerified, EmailArcFailureReason.NotEvaluated, text, known, diagnostics)
                : new EmailArcVerificationResult(EmailArcChainStatus.Fail, reason, text, known, diagnostics);
        }
    }

    /// <summary>Validates the ARC chain over the source bytes retained by the reader.</summary>
    public static Task<EmailArcVerificationResult> VerifyAsync(
        EmailDocument document,
        IEmailDkimKeyResolver keyResolver,
        IOfficePublicKeySignatureVerifier signatureVerifier,
        EmailAuthenticationOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(document);
        if (document.RawSource == null) {
            return Task.FromResult(new EmailArcVerificationResult(EmailArcChainStatus.NotVerified, EmailArcFailureReason.NotEvaluated,
                "The document does not retain its original bytes.", Array.Empty<EmailArcSet>(),
                new[] { EmailAuthenticationParsing.RawSourceMissing("arc") }));
        }

        return VerifyAsync(document.RawSource, keyResolver, signatureVerifier, options, cancellationToken);
    }

    private static bool IsComplete(IEnumerable<int> instances, int highest) {
        int[] sorted = instances.OrderBy(value => value).ToArray();
        if (sorted.Length != highest) return false;
        for (int index = 0; index < sorted.Length; index++) {
            if (sorted[index] != index + 1) return false;
        }

        return true;
    }

    /// <summary>Reads the mandatory leading "i=N;" of an ARC-Authentication-Results value; zero when absent or invalid.</summary>
    private static int ReadAuthenticationResultsInstance(string value, out string? payload) {
        payload = null;
        int position = 0;
        while (position < value.Length && EmailDkimTagList.IsFws(value[position])) position++;
        if (position + 2 > value.Length || value[position] != 'i' || value[position + 1] != '=') return 0;
        position += 2;
        while (position < value.Length && EmailDkimTagList.IsFws(value[position])) position++;
        int digitsStart = position;
        while (position < value.Length && value[position] >= '0' && value[position] <= '9') position++;
        string digits = value.Substring(digitsStart, position - digitsStart);
        while (position < value.Length && EmailDkimTagList.IsFws(value[position])) position++;
        if (position >= value.Length || value[position] != ';') return 0;
        if (!EmailDkimEvaluator.TryParseInstance(digits, out int instance)) return 0;
        payload = value.Substring(position + 1).Trim();
        return instance;
    }
}
