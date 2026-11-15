using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.Cms;
using Org.BouncyCastle.Asn1.Ocsp;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Cms;
using Org.BouncyCastle.Ocsp;
using Org.BouncyCastle.Tsp;
using Org.BouncyCastle.X509;
using BcX509Certificate = Org.BouncyCastle.X509.X509Certificate;

namespace OfficeIMO.Security;

/// <summary>
/// Checks OCSP responses and CRLs for the signer and time-stamp authority paths of one CMS container
/// (RFC 5280 section 6.3, RFC 6960, ETSI EN 319 102-1 past-signature validation).
/// </summary>
internal static class CertificateRevocationEvaluator {
    private const string SigningCertificateV2Oid = "1.2.840.113549.1.9.16.2.47";
    private const string SigningCertificateOid = "1.2.840.113549.1.9.16.2.12";
    private const string SignatureTimestampTokenOid = "1.2.840.113549.1.9.16.2.14";
    private const string AdobeRevocationInfoArchivalOid = "1.2.840.113583.1.1.8";
    private const string OcspResponseRevocationInfoFormatOid = "1.3.6.1.5.5.7.16.2";
    private const string OcspSigningExtendedKeyUsageOid = "1.3.6.1.5.5.7.3.9";
    private static readonly DerObjectIdentifier OcspArchiveCutoffOid = new("1.3.6.1.5.5.7.48.1.6");
    private static readonly DerObjectIdentifier ExpiredCertificatesOnCrlOid = new("2.5.29.60");
    private const int MaxPathLength = 16;
    private const int MaxTokens = 512;
    private static readonly TimeSpan ClockSkew = TimeSpan.FromMinutes(5);
    private static readonly HashSet<string> SupportedCriticalCrlExtensions = new(StringComparer.Ordinal) {
        X509Extensions.IssuingDistributionPoint.Id,
        X509Extensions.CrlNumber.Id,
        X509Extensions.AuthorityKeyIdentifier.Id,
        "2.5.29.60"
    };
    private static readonly HashSet<string> SupportedCriticalCrlEntryExtensions = new(StringComparer.Ordinal) {
        X509Extensions.ReasonCode.Id,
        X509Extensions.InvalidityDate.Id
    };

    internal static OfficeRevocationOutcome Validate(OfficeRevocationRequest request) {
        var findings = new List<SecurityFinding>();
        ParsedCms parsed;
        try {
            parsed = ParsedCms.Parse(request.EncodedCms, request.IsTimestampToken, ToBouncyCastle(request.ExtraCertificates));
        } catch (Exception exception) when (IsValidationException(exception)) {
            findings.Add(Finding(SecurityFindingSeverity.Error, "RevocationContainerUnreadable",
                "Revocation could not be checked because the CMS container could not be read: " + exception.Message));
            return new OfficeRevocationOutcome(SecurityValidationStatus.Indeterminate, SecurityValidationStatus.NotPerformed,
                false, false, false, false, findings);
        }
        if (parsed.SigningCertificateMismatch) {
            findings.Add(Finding(SecurityFindingSeverity.Error, "SigningCertificateMismatch",
                "The ESS signing-certificate attribute identifies a different certificate than the CMS signer's (RFC 5035)."));
        }

        DateTime now = (request.Now ?? DateTimeOffset.UtcNow).UtcDateTime;
        var pool = new CertificatePool();
        pool.AddRange(parsed.Certificates);
        foreach (byte[] encoded in request.DocumentCertificates) pool.TryAdd(encoded);
        if (request.ExtraCertificates is not null) {
            foreach (X509Certificate2 certificate in request.ExtraCertificates) pool.TryAdd(certificate.RawData);
        }

        var tokens = new List<RevocationToken>();
        if (request.Policy != CertificateRevocationPolicy.Online) {
            foreach (EmbeddedValue value in parsed.EmbeddedTokens) {
                AddToken(tokens, value.Encoded, value.IsOcsp, OfficeRevocationTokenSource.Signature, null, findings);
            }
            foreach (OfficeRevocationToken token in request.DocumentTokens) {
                AddToken(tokens, token.Encoded, token.IsOcspResponse, OfficeRevocationTokenSource.Document,
                    token.ProofOfExistence?.UtcDateTime, findings);
            }
        }

        var context = new EvaluationContext(pool, tokens, now, request, findings);
        var signerStatus = new StatusAccumulator();
        var timestampStatus = new StatusAccumulator();
        foreach (RevocationSubject subject in parsed.Subjects) {
            DateTime validationTime = subject.IsTimestampAuthority && subject.Time.HasValue
                ? subject.Time.Value
                : request.SignerValidationTime?.UtcDateTime ?? subject.Time ?? now;
            StatusAccumulator target = subject.IsTimestampAuthority ? timestampStatus : signerStatus;
            if (subject.Certificate is null) {
                findings.Add(Finding(SecurityFindingSeverity.Warning, "RevocationSubjectCertificateMissing",
                    subject.Role + " certificate is not available, so its revocation status cannot be checked."));
                target.Add(SecurityValidationStatus.Indeterminate);
                continue;
            }
            target.Add(CheckPath(context, subject, validationTime));
        }

        SecurityValidationStatus status = StatusAccumulator.Combine(signerStatus, timestampStatus);
        if (status == SecurityValidationStatus.NotPerformed) {
            findings.Add(Finding(SecurityFindingSeverity.Info, "RevocationNotApplicable",
                "Every certificate path ends at a self-signed certificate directly, so there is no issued certificate to check."));
        }
        return new OfficeRevocationOutcome(
            status,
            timestampStatus.Result,
            parsed.HasSigningCertificateAttribute,
            parsed.SigningCertificateMismatch,
            context.CheckedCertificates > 0 && context.CheckedCertificates == context.SettledByDocument,
            context.SettledByDocument > 0 && context.SettledByDocument == context.SettledByDocumentWithProof,
            findings);
    }

    internal static async Task<CertificateRevocationEvidence> FetchAsync(
        byte[] encodedCms,
        bool isTimestampToken,
        ICertificateRevocationFetcher fetcher,
        X509Certificate2Collection? extraCertificates,
        CancellationToken cancellationToken) {
        ParsedCms parsed = ParsedCms.Parse(encodedCms, isTimestampToken, ToBouncyCastle(extraCertificates));
        var pool = new CertificatePool();
        pool.AddRange(parsed.Certificates);
        if (extraCertificates is not null) {
            foreach (X509Certificate2 certificate in extraCertificates) pool.TryAdd(certificate.RawData);
        }

        var findings = new List<SecurityFinding>();
        var certificates = new EncodedSet();
        var ocsp = new EncodedSet();
        var crls = new EncodedSet();
        var requested = new HashSet<string>(StringComparer.Ordinal);
        foreach (RevocationSubject subject in parsed.Subjects) {
            cancellationToken.ThrowIfCancellationRequested();
            if (subject.Certificate is null) continue;
            List<PathLink> path = BuildPath(pool, subject.Certificate, out BcX509Certificate? unresolved);
            foreach (PathLink link in path) {
                certificates.Add(link.Certificate.GetEncoded());
                if (link.Issuer is null) continue;
                certificates.Add(link.Issuer.GetEncoded());
                if (!requested.Add(Identity(link.Certificate))) continue;
                using X509Certificate2 certificate = PlatformCertificateLoader.Load(link.Certificate.GetEncoded());
                using X509Certificate2 issuer = PlatformCertificateLoader.Load(link.Issuer.GetEncoded());
                CertificateRevocationEvidence evidence = await fetcher.FetchAsync(certificate, issuer, cancellationToken)
                    .ConfigureAwait(false) ?? CertificateRevocationEvidence.Empty;
                foreach (byte[] value in evidence.CertificateValues) certificates.Add(value);
                foreach (byte[] value in evidence.OcspValues) ocsp.Add(value);
                foreach (byte[] value in evidence.CrlValues) crls.Add(value);
                findings.AddRange(evidence.Findings);
            }
            if (unresolved is not null) {
                findings.Add(Finding(SecurityFindingSeverity.Warning, "RevocationIssuerMissing",
                    subject.Role + " path stops at '" + unresolved.SubjectDN + "' because its issuer certificate is not available."));
            }
        }
        return new CertificateRevocationEvidence(certificates.Values, ocsp.Values, crls.Values, findings);
    }

    private static SecurityValidationStatus CheckPath(EvaluationContext context, RevocationSubject subject, DateTime validationTime) {
        List<PathLink> path = BuildPath(context.Pool, subject.Certificate!, out BcX509Certificate? unresolved);
        var status = new StatusAccumulator();
        foreach (PathLink link in path) {
            if (link.Issuer is null) continue;
            status.Add(CheckCertificate(context, subject.Role, link.Certificate, link.Issuer, validationTime));
        }
        if (unresolved is not null) {
            context.Findings.Add(Finding(SecurityFindingSeverity.Warning, "RevocationIssuerMissing",
                subject.Role + " certificate '" + unresolved.SubjectDN + "' cannot be checked because its issuer certificate is not available."));
            status.Add(SecurityValidationStatus.Indeterminate);
        }
        return status.Result;
    }

    private static SecurityValidationStatus CheckCertificate(
        EvaluationContext context,
        string role,
        BcX509Certificate certificate,
        BcX509Certificate issuer,
        DateTime validationTime) {
        context.CheckedCertificates++;
        string label = role + " certificate '" + certificate.SubjectDN + "'";
        CertificateRevocationPolicy policy = context.Request.Policy;
        Decision decision = Decide(context, label, certificate, issuer, validationTime, context.Tokens);
        if (decision.Status == SecurityValidationStatus.Indeterminate && policy != CertificateRevocationPolicy.EmbeddedOnly) {
            List<RevocationToken> online = FetchOnline(context, label, certificate, issuer);
            if (online.Count > 0) {
                Decision onlineDecision = Decide(context, label, certificate, issuer, validationTime, online);
                if (onlineDecision.Status != SecurityValidationStatus.Indeterminate) decision = onlineDecision;
                else decision.Reasons.AddRange(onlineDecision.Reasons);
            }
        }

        switch (decision.Status) {
            case SecurityValidationStatus.Invalid:
                context.Findings.Add(Finding(SecurityFindingSeverity.Error, "RevocationCertificateRevoked",
                    label + " was revoked at " + Format(decision.RevocationTime) + ", before the validation time " +
                    Format(validationTime) + " (" + decision.SourceText + ")."));
                if (decision.Source == OfficeRevocationTokenSource.Document) context.SettledByDocument++;
                break;
            case SecurityValidationStatus.Valid:
                if (decision.RevocationTime.HasValue) {
                    context.Findings.Add(Finding(SecurityFindingSeverity.Warning, "RevocationAfterValidationTime",
                        label + " was revoked at " + Format(decision.RevocationTime) +
                        ", after the validation time " + Format(validationTime) + "."));
                }
                if (decision.Source == OfficeRevocationTokenSource.Document) {
                    context.SettledByDocument++;
                    if (decision.HasProofOfExistence) context.SettledByDocumentWithProof++;
                }
                break;
            default:
                string reason = decision.Reasons.Count == 0
                    ? "no OCSP response or CRL for it is available"
                    : string.Join("; ", decision.Reasons.Distinct(StringComparer.Ordinal));
                context.Findings.Add(Finding(SecurityFindingSeverity.Warning,
                    decision.Reasons.Count == 0 ? "RevocationDataMissing" : "RevocationDataUnusable",
                    "Revocation status of " + label + " at " + Format(validationTime) + " is unknown: " + reason + "."));
                break;
        }
        return decision.Status;
    }

    private static List<RevocationToken> FetchOnline(
        EvaluationContext context,
        string label,
        BcX509Certificate certificate,
        BcX509Certificate issuer) {
        var result = new List<RevocationToken>();
        ICertificateRevocationFetcher? fetcher = context.Request.Fetcher;
        if (fetcher is null) {
            context.Findings.Add(Finding(SecurityFindingSeverity.Warning, "RevocationFetcherMissing",
                "Online revocation checking was requested for " + label + ", but no fetcher was supplied."));
            return result;
        }

        try {
            using X509Certificate2 subject = PlatformCertificateLoader.Load(certificate.GetEncoded());
            using X509Certificate2 issuerCertificate = PlatformCertificateLoader.Load(issuer.GetEncoded());
            CancellationToken cancellationToken = context.Request.CancellationToken;
            // Validation is synchronous; run the fetch on the thread pool so a caller's synchronization
            // context cannot deadlock against continuations captured by the fetcher.
            CertificateRevocationEvidence evidence = Task.Run(
                () => fetcher.FetchAsync(subject, issuerCertificate, cancellationToken), cancellationToken)
                .GetAwaiter().GetResult() ?? CertificateRevocationEvidence.Empty;
            context.Findings.AddRange(evidence.Findings);
            foreach (byte[] value in evidence.CertificateValues) context.Pool.TryAdd(value);
            foreach (byte[] value in evidence.OcspValues) {
                AddToken(result, value, isOcsp: true, OfficeRevocationTokenSource.Online, null, context.Findings);
            }
            foreach (byte[] value in evidence.CrlValues) {
                AddToken(result, value, isOcsp: false, OfficeRevocationTokenSource.Online, null, context.Findings);
            }
        } catch (OperationCanceledException) when (context.Request.CancellationToken.IsCancellationRequested) {
            throw;
        } catch (OperationCanceledException exception) {
            // A fetcher's own timeout surfaces as cancellation; it is a source failure, not a cancelled validation.
            context.Findings.Add(Finding(SecurityFindingSeverity.Warning, "RevocationFetchFailed",
                "Online revocation data for " + label + " could not be fetched: " + exception.Message));
        } catch (Exception exception) when (IsValidationException(exception)) {
            context.Findings.Add(Finding(SecurityFindingSeverity.Warning, "RevocationFetchFailed",
                "Online revocation data for " + label + " could not be fetched: " + exception.Message));
        }
        return result;
    }

    private static Decision Decide(
        EvaluationContext context,
        string label,
        BcX509Certificate certificate,
        BcX509Certificate issuer,
        DateTime validationTime,
        IReadOnlyList<RevocationToken> tokens) {
        var decision = new Decision();
        Candidate? best = null;
        foreach (RevocationToken token in tokens) {
            Candidate? candidate = token.Ocsp is not null
                ? EvaluateOcsp(context, token, certificate, issuer, validationTime, decision.Reasons)
                : token.Crl is not null
                    ? EvaluateCrl(context, token, certificate, issuer, validationTime, decision.Reasons)
                    : null;
            if (candidate is null) continue;
            if (candidate.Revoked) {
                decision.Status = SecurityValidationStatus.Invalid;
                decision.RevocationTime = candidate.RevocationTime;
                decision.Source = token.Source;
                decision.SourceText = Describe(token);
                return decision;
            }
            if (best is null || Rank(candidate, token) > Rank(best, best.Token)) best = candidate;
        }

        if (best is null) {
            decision.Status = SecurityValidationStatus.Indeterminate;
            return decision;
        }
        decision.Status = SecurityValidationStatus.Valid;
        decision.RevocationTime = best.RevocationTime;
        decision.Source = best.Token.Source;
        decision.SourceText = Describe(best.Token);
        decision.HasProofOfExistence = best.Token.ProofOfExistence.HasValue;
        return decision;

        static int Rank(Candidate candidate, RevocationToken token) =>
            (token.Source == OfficeRevocationTokenSource.Document ? 4 : 0) +
            (token.ProofOfExistence.HasValue ? 2 : 0) +
            (candidate.RevocationTime.HasValue ? 0 : 1);
    }

    private static Candidate? EvaluateOcsp(
        EvaluationContext context,
        RevocationToken token,
        BcX509Certificate certificate,
        BcX509Certificate issuer,
        DateTime validationTime,
        List<string> reasons) {
        BasicOcspResp response = token.Ocsp!;
        SingleResp? single = null;
        foreach (SingleResp candidate in response.Responses) {
            CertificateID id = candidate.GetCertID();
            if (!id.SerialNumber.Equals(certificate.SerialNumber)) continue;
            bool matches;
            try {
                matches = id.MatchesIssuer(issuer);
            } catch (OcspException) {
                matches = false;
            }
            if (!matches) continue;
            single = candidate;
            break;
        }
        if (single is null) return null;

        BcX509Certificate? responder = FindOcspResponder(response, issuer, context.Pool.All);
        if (responder is null) {
            reasons.Add("the OCSP response (" + Describe(token) + ") is not signed by the issuer or an authorized responder");
            return null;
        }
        bool signatureValid;
        try {
            signatureValid = response.Verify(responder.GetPublicKey());
        } catch (Exception exception) when (IsValidationException(exception)) {
            signatureValid = false;
        }
        if (!signatureValid) {
            reasons.Add("the OCSP response signature (" + Describe(token) + ") does not verify");
            return null;
        }
        DateTime producedAt = response.ProducedAt.ToUniversalTime();
        if (!IsSignerUsable(responder, producedAt, token.ProofOfExistence, context.Now, out string? signerProblem)) {
            reasons.Add("the OCSP responder certificate " + signerProblem + " (" + Describe(token) + ")");
            return null;
        }

        DateTime thisUpdate = single.ThisUpdate.ToUniversalTime();
        DateTime? nextUpdate = single.NextUpdate?.ToUniversalTime();
        object? status = single.GetCertStatus();
        if (status is RevokedStatus revoked) {
            DateTime revocationTime = revoked.RevocationTime.ToUniversalTime();
            if (revocationTime <= validationTime) {
                return new Candidate(token) { Revoked = true, RevocationTime = revocationTime };
            }
            if (!IsFresh(thisUpdate, nextUpdate, validationTime, context.Now)) {
                reasons.Add(StaleReason("OCSP response", token, thisUpdate, nextUpdate));
                return null;
            }
            return new Candidate(token) { RevocationTime = revocationTime };
        }
        if (status is not null) {
            reasons.Add("the OCSP responder reports the status as unknown (" + Describe(token) + ")");
            return null;
        }
        if (!IsFresh(thisUpdate, nextUpdate, validationTime, context.Now)) {
            reasons.Add(StaleReason("OCSP response", token, thisUpdate, nextUpdate));
            return null;
        }
        // RFC 6960 section 4.4.4: after expiry a "good" answer only proves history back to the archive cutoff.
        if (thisUpdate > certificate.NotAfter.ToUniversalTime() &&
            !(ReadGeneralizedTime(single.GetExtensionValue(OcspArchiveCutoffOid)) is DateTime cutoff &&
              cutoff <= certificate.NotAfter.ToUniversalTime())) {
            reasons.Add("the OCSP response (" + Describe(token) + ") was produced after the certificate expired and has no archive cutoff covering its validity");
            return null;
        }
        return new Candidate(token);
    }

    /// <summary>True when an OCSP response carries a status for the certificate signed by its issuer or an authorized responder.</summary>
    internal static bool HasVerifiedStatus(BasicOcspResp response, BcX509Certificate certificate, BcX509Certificate issuer) {
        bool covered = response.Responses.Any(single => {
            CertificateID id = single.GetCertID();
            try {
                return id.SerialNumber.Equals(certificate.SerialNumber) && id.MatchesIssuer(issuer) &&
                    single.GetCertStatus() is not UnknownStatus;
            } catch (OcspException) {
                return false;
            }
        });
        if (!covered) return false;
        BcX509Certificate? responder = FindOcspResponder(response, issuer, Array.Empty<BcX509Certificate>());
        if (responder is null) return false;
        try {
            return response.Verify(responder.GetPublicKey());
        } catch (Exception exception) when (IsValidationException(exception)) {
            return false;
        }
    }

    private static DateTime? ReadGeneralizedTime(Asn1OctetString? extension) {
        if (extension is null) return null;
        try {
            return Asn1GeneralizedTime.GetInstance(Asn1Object.FromByteArray(extension.GetOctets())).ToDateTime().ToUniversalTime();
        } catch (Exception exception) when (IsValidationException(exception)) {
            return null;
        }
    }

    /// <summary>
    /// Only complete, direct CRLs are evaluated: delta CRLs and partitioned or reason-limited scopes
    /// (RFC 5280 sections 5.2.4 and 5.2.5) cannot prove that a certificate is absent.
    /// </summary>
    private static string? CheckCrlScope(X509Crl crl, BcX509Certificate certificate) {
        if (crl.GetExtensionValue(X509Extensions.DeltaCrlIndicator) is not null) return "is a delta CRL, which is not supported";
        // RFC 5280 sections 5.2 and 5.3: a CRL with a critical CRL or entry extension that cannot be processed
        // must not be used to determine the status of any certificate.
        foreach (string oid in crl.GetCriticalExtensionOids() ?? new HashSet<string>()) {
            if (!SupportedCriticalCrlExtensions.Contains(oid)) return "has the unsupported critical extension " + oid;
        }
        foreach (X509CrlEntry entry in crl.GetRevokedCertificates() ?? new HashSet<X509CrlEntry>()) {
            foreach (string oid in entry.GetCriticalExtensionOids() ?? new HashSet<string>()) {
                if (!SupportedCriticalCrlEntryExtensions.Contains(oid)) return "has an entry with the unsupported critical extension " + oid;
            }
        }
        Asn1OctetString? value = crl.GetExtensionValue(X509Extensions.IssuingDistributionPoint);
        if (value is null) return null;
        IssuingDistributionPoint scope;
        try {
            scope = IssuingDistributionPoint.GetInstance(Asn1Object.FromByteArray(value.GetOctets()));
        } catch (Exception exception) when (IsValidationException(exception)) {
            return "has an unreadable issuing distribution point";
        }
        if (scope.IsIndirectCrl) return "is an indirect CRL, which is not supported";
        if (scope.OnlySomeReasons is not null) return "covers only some revocation reasons";
        if (scope.OnlyContainsAttributeCerts) return "covers only attribute certificates";
        bool isAuthority = certificate.GetBasicConstraints() >= 0;
        if (scope.OnlyContainsCACerts && !isAuthority) return "covers only CA certificates";
        if (scope.OnlyContainsUserCerts && isAuthority) return "covers only end-entity certificates";
        if (scope.DistributionPoint?.Type == DistributionPointName.FullName) {
            HashSet<string> crlNames = UriNames(scope.DistributionPoint.Name);
            Asn1OctetString? points = certificate.GetExtensionValue(X509Extensions.CrlDistributionPoints);
            if (points is not null) {
                var certificateNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (DistributionPoint point in CrlDistPoint.GetInstance(Asn1Object.FromByteArray(points.GetOctets())).GetDistributionPoints()) {
                    if (point.DistributionPointName?.Type == DistributionPointName.FullName) {
                        certificateNames.UnionWith(UriNames(point.DistributionPointName.Name));
                    }
                }
                if (certificateNames.Count > 0 && crlNames.Count > 0 && !certificateNames.Overlaps(crlNames)) {
                    return "belongs to a different distribution point than the certificate names";
                }
            }
        }
        return null;
    }

    private static HashSet<string> UriNames(Asn1Encodable names) {
        var result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (GeneralName name in GeneralNames.GetInstance(names).GetNames()) {
            if (name.TagNo == GeneralName.UniformResourceIdentifier) result.Add(DerIA5String.GetInstance(name.Name).GetString());
        }
        return result;
    }

    private static Candidate? EvaluateCrl(
        EvaluationContext context,
        RevocationToken token,
        BcX509Certificate certificate,
        BcX509Certificate issuer,
        DateTime validationTime,
        List<string> reasons) {
        X509Crl crl = token.Crl!;
        if (!crl.IssuerDN.Equivalent(issuer.SubjectDN, inOrder: true)) return null;
        string? scopeProblem = CheckCrlScope(crl, certificate);
        if (scopeProblem is not null) {
            reasons.Add("the CRL (" + Describe(token) + ") " + scopeProblem);
            return null;
        }
        if (!AllowsUsage(issuer, Org.BouncyCastle.Asn1.X509.KeyUsage.CrlSign)) {
            reasons.Add("the CRL issuer key usage does not permit CRL signing (" + Describe(token) + ")");
            return null;
        }
        try {
            crl.Verify(issuer.GetPublicKey());
        } catch (Exception exception) when (IsValidationException(exception)) {
            reasons.Add("the CRL signature (" + Describe(token) + ") does not verify");
            return null;
        }
        DateTime thisUpdate = crl.ThisUpdate.ToUniversalTime();
        DateTime? nextUpdate = crl.NextUpdate?.ToUniversalTime();
        if (!IsSignerUsable(issuer, thisUpdate, token.ProofOfExistence, context.Now, out string? signerProblem)) {
            reasons.Add("the CRL issuer certificate " + signerProblem + " (" + Describe(token) + ")");
            return null;
        }

        X509CrlEntry? entry = crl.GetRevokedCertificate(certificate.SerialNumber);
        if (entry is not null) {
            DateTime revocationTime = entry.RevocationDate.ToUniversalTime();
            if (revocationTime <= validationTime) {
                return new Candidate(token) { Revoked = true, RevocationTime = revocationTime };
            }
            if (!IsFresh(thisUpdate, nextUpdate, validationTime, context.Now)) {
                reasons.Add(StaleReason("CRL", token, thisUpdate, nextUpdate));
                return null;
            }
            return new Candidate(token) { RevocationTime = revocationTime };
        }
        if (!IsFresh(thisUpdate, nextUpdate, validationTime, context.Now)) {
            reasons.Add(StaleReason("CRL", token, thisUpdate, nextUpdate));
            return null;
        }
        // RFC 5280 section 3.3: CAs may drop expired certificates, so absence after expiry proves nothing
        // unless the CRL declares it keeps entries for certificates that expired since ExpiredCertsOnCRL.
        if (thisUpdate > certificate.NotAfter.ToUniversalTime() &&
            !(ReadGeneralizedTime(crl.GetExtensionValue(ExpiredCertificatesOnCrlOid)) is DateTime kept &&
              kept <= certificate.NotAfter.ToUniversalTime())) {
            reasons.Add("the CRL (" + Describe(token) + ") was issued after the certificate expired and does not state that it keeps expired certificates");
            return null;
        }
        return new Candidate(token);
    }

    /// <summary>
    /// Revocation data proves the status at the validation time when it was issued at or after that time,
    /// or when its validity window (thisUpdate to nextUpdate) contains it.
    /// </summary>
    private static bool IsFresh(DateTime thisUpdate, DateTime? nextUpdate, DateTime validationTime, DateTime now) {
        if (thisUpdate > now + ClockSkew) return false;
        if (thisUpdate >= validationTime - ClockSkew) return true;
        return nextUpdate.HasValue && validationTime <= nextUpdate.Value;
    }

    /// <summary>
    /// A token signer must have been valid when it signed, and must still be valid now unless a trusted
    /// time-stamp proves the token existed before the signer certificate expired.
    /// </summary>
    private static bool IsSignerUsable(
        BcX509Certificate signer,
        DateTime signedAt,
        DateTime? proofOfExistence,
        DateTime now,
        out string? problem) {
        DateTime notBefore = signer.NotBefore.ToUniversalTime();
        DateTime notAfter = signer.NotAfter.ToUniversalTime();
        if (signedAt < notBefore - ClockSkew || signedAt > notAfter + ClockSkew) {
            problem = "was not valid when the data was produced";
            return false;
        }
        DateTime existedBy = proofOfExistence ?? now;
        if (existedBy > notAfter) {
            problem = proofOfExistence.HasValue
                ? "expired before the earliest time-stamp that covers the data"
                : "has expired and no time-stamp proves the data existed before it expired";
            return false;
        }
        problem = null;
        return true;
    }

    private static BcX509Certificate? FindOcspResponder(BasicOcspResp response, BcX509Certificate issuer, IEnumerable<BcX509Certificate> pool) {
        if (MatchesResponderId(response.ResponderId, issuer)) return issuer;
        IEnumerable<BcX509Certificate> candidates = response.GetCerts() ?? Array.Empty<BcX509Certificate>();
        foreach (BcX509Certificate candidate in candidates.Concat(pool)) {
            if (!MatchesResponderId(response.ResponderId, candidate)) continue;
            // RFC 6960 section 4.2.2.2: a delegated responder is issued directly by the CA and carries id-kp-OCSPSigning.
            if (!candidate.IssuerDN.Equivalent(issuer.SubjectDN, inOrder: true) || !IsSignedBy(candidate, issuer)) continue;
            if (!PermitsSigning(candidate)) continue;
            IList<DerObjectIdentifier>? usages = candidate.GetExtendedKeyUsage();
            if (usages is null || !usages.Any(static usage => usage.Id == OcspSigningExtendedKeyUsageOid)) continue;
            return candidate;
        }
        return null;
    }

    private static bool MatchesResponderId(RespID responderId, BcX509Certificate certificate) {
        try {
            return responderId.Equals(new RespID(certificate.SubjectDN)) ||
                responderId.Equals(new RespID(certificate.GetPublicKey()));
        } catch (Exception exception) when (IsValidationException(exception)) {
            return false;
        }
    }

    /// <summary>RFC 5280 section 4.2.1.3: a key used to sign OCSP responses needs digitalSignature or nonRepudiation.</summary>
    private static bool PermitsSigning(BcX509Certificate certificate) {
        bool[]? keyUsage = certificate.GetKeyUsage();
        return keyUsage is null || (keyUsage.Length > 0 && keyUsage[0]) || (keyUsage.Length > 1 && keyUsage[1]);
    }

    private static bool AllowsUsage(BcX509Certificate certificate, int usage) {
        bool[]? keyUsage = certificate.GetKeyUsage();
        if (keyUsage is null) return true;
        int bit = usage switch {
            Org.BouncyCastle.Asn1.X509.KeyUsage.CrlSign => 6,
            Org.BouncyCastle.Asn1.X509.KeyUsage.KeyCertSign => 5,
            _ => -1
        };
        return bit < 0 || (bit < keyUsage.Length && keyUsage[bit]);
    }

    private static List<PathLink> BuildPath(CertificatePool pool, BcX509Certificate leaf, out BcX509Certificate? unresolved) {
        var path = new List<PathLink>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        BcX509Certificate current = leaf;
        unresolved = null;
        while (path.Count < MaxPathLength && seen.Add(Identity(current))) {
            if (IsSelfSigned(current)) {
                path.Add(new PathLink(current, null));
                return path;
            }
            BcX509Certificate? issuer = pool.All.FirstOrDefault(candidate =>
                candidate.SubjectDN.Equivalent(current.IssuerDN, inOrder: true) &&
                !string.Equals(Identity(candidate), Identity(current), StringComparison.Ordinal) &&
                IsSignedBy(current, candidate));
            if (issuer is null) {
                path.Add(new PathLink(current, null));
                unresolved = current;
                return path;
            }
            path.Add(new PathLink(current, issuer));
            current = issuer;
        }
        unresolved = current;
        return path;
    }

    private static bool IsSelfSigned(BcX509Certificate certificate) =>
        certificate.IssuerDN.Equivalent(certificate.SubjectDN, inOrder: true) && IsSignedBy(certificate, certificate);

    private static bool IsSignedBy(BcX509Certificate certificate, BcX509Certificate issuer) {
        try {
            certificate.Verify(issuer.GetPublicKey());
            return true;
        } catch (Exception exception) when (IsValidationException(exception)) {
            return false;
        }
    }

    private static void AddToken(
        List<RevocationToken> tokens,
        byte[] encoded,
        bool isOcsp,
        OfficeRevocationTokenSource source,
        DateTime? proofOfExistence,
        List<SecurityFinding> findings) {
        if (tokens.Count >= MaxTokens) {
            if (tokens.Count == MaxTokens) {
                findings.Add(Finding(SecurityFindingSeverity.Warning, "RevocationTokenLimitExceeded",
                    "Only the first " + MaxTokens + " OCSP responses and CRLs were considered."));
                tokens.Add(new RevocationToken(source, proofOfExistence));
            }
            return;
        }
        var token = new RevocationToken(source, proofOfExistence);
        try {
            if (isOcsp) {
                var response = new OcspResp(encoded);
                if (response.Status != OcspRespStatus.Successful) {
                    findings.Add(Finding(SecurityFindingSeverity.Info, "RevocationOcspUnsuccessful",
                        "An OCSP response (" + Describe(token) + ") reports status " + response.Status + " and carries no certificate status."));
                    return;
                }
                token.Ocsp = response.GetResponseObject() as BasicOcspResp;
                if (token.Ocsp is null) {
                    findings.Add(Finding(SecurityFindingSeverity.Warning, "RevocationTokenMalformed",
                        "An OCSP response (" + Describe(token) + ") is not a basic OCSP response."));
                    return;
                }
            } else {
                token.Crl = new X509CrlParser().ReadCrl(encoded);
                if (token.Crl is null) throw new InvalidDataException("No CRL was decoded.");
            }
            tokens.Add(token);
        } catch (Exception exception) when (IsValidationException(exception)) {
            findings.Add(Finding(SecurityFindingSeverity.Warning, "RevocationTokenMalformed",
                "A " + (isOcsp ? "OCSP response" : "CRL") + " (" + Describe(token) + ") could not be decoded: " + exception.Message));
        }
    }

    private static string Describe(RevocationToken token) => token.Source switch {
        OfficeRevocationTokenSource.Document => token.ProofOfExistence.HasValue
            ? "document validation data, time-stamped " + Format(token.ProofOfExistence)
            : "document validation data",
        OfficeRevocationTokenSource.Online => "fetched online",
        _ => "embedded in the signature"
    };

    private static string StaleReason(string kind, RevocationToken token, DateTime thisUpdate, DateTime? nextUpdate) =>
        "the " + kind + " (" + Describe(token) + ") was issued at " + Format(thisUpdate) +
        (nextUpdate.HasValue ? " and is valid until " + Format(nextUpdate) : " without a next update") +
        ", which does not cover the validation time";

    private static string Format(DateTime? value) =>
        value.HasValue ? value.Value.ToUniversalTime().ToString("u", System.Globalization.CultureInfo.InvariantCulture) : "an unknown time";

    private static string Format(DateTimeOffset? value) => Format(value?.UtcDateTime);

    private static string Identity(BcX509Certificate certificate) => Convert.ToBase64String(Sha256(certificate.GetEncoded()));

    private static byte[] Sha256(byte[] value) {
#if NET8_0_OR_GREATER
        return SHA256.HashData(value);
#else
        using SHA256 sha256 = SHA256.Create();
        return sha256.ComputeHash(value);
#endif
    }

    private static SecurityFinding Finding(SecurityFindingSeverity severity, string code, string message) =>
        new SecurityFinding(severity, code, message);

    private static bool IsValidationException(Exception exception) =>
        exception is not OutOfMemoryException &&
        exception is not StackOverflowException &&
        exception is not AccessViolationException &&
        exception is not OperationCanceledException;

    private static IReadOnlyList<BcX509Certificate> ToBouncyCastle(X509Certificate2Collection? certificates) {
        if (certificates is null || certificates.Count == 0) return Array.Empty<BcX509Certificate>();
        var parser = new X509CertificateParser();
        var result = new List<BcX509Certificate>(certificates.Count);
        foreach (X509Certificate2 certificate in certificates) {
            try {
                BcX509Certificate? converted = parser.ReadCertificate(certificate.RawData);
                if (converted is not null) result.Add(converted);
            } catch (Exception exception) when (IsValidationException(exception)) {
                // An unreadable supplied certificate cannot identify a signer.
            }
        }
        return result;
    }

    private enum SigningCertificateBinding {
        Absent,
        Valid,
        Mismatch
    }

    private sealed class EvaluationContext {
        internal EvaluationContext(
            CertificatePool pool,
            List<RevocationToken> tokens,
            DateTime now,
            OfficeRevocationRequest request,
            List<SecurityFinding> findings) {
            Pool = pool;
            Tokens = tokens;
            Now = now;
            Request = request;
            Findings = findings;
        }

        internal CertificatePool Pool { get; }
        internal List<RevocationToken> Tokens { get; }
        internal DateTime Now { get; }
        internal OfficeRevocationRequest Request { get; }
        internal List<SecurityFinding> Findings { get; }
        internal int CheckedCertificates { get; set; }
        internal int SettledByDocument { get; set; }
        internal int SettledByDocumentWithProof { get; set; }
    }

    private sealed class RevocationToken {
        internal RevocationToken(OfficeRevocationTokenSource source, DateTime? proofOfExistence) {
            Source = source;
            ProofOfExistence = proofOfExistence;
        }

        internal OfficeRevocationTokenSource Source { get; }
        internal DateTime? ProofOfExistence { get; }
        internal BasicOcspResp? Ocsp { get; set; }
        internal X509Crl? Crl { get; set; }
    }

    private sealed class Candidate {
        internal Candidate(RevocationToken token) => Token = token;
        internal RevocationToken Token { get; }
        internal bool Revoked { get; set; }
        internal DateTime? RevocationTime { get; set; }
    }

    private sealed class Decision {
        internal SecurityValidationStatus Status { get; set; } = SecurityValidationStatus.Indeterminate;
        internal DateTime? RevocationTime { get; set; }
        internal OfficeRevocationTokenSource Source { get; set; }
        internal string SourceText { get; set; } = string.Empty;
        internal bool HasProofOfExistence { get; set; }
        internal List<string> Reasons { get; } = new List<string>();
    }

    private readonly struct PathLink {
        internal PathLink(BcX509Certificate certificate, BcX509Certificate? issuer) {
            Certificate = certificate;
            Issuer = issuer;
        }

        internal BcX509Certificate Certificate { get; }
        internal BcX509Certificate? Issuer { get; }
    }

    private sealed class StatusAccumulator {
        private bool _any;
        private bool _invalid;
        private bool _indeterminate;

        internal void Add(SecurityValidationStatus status) {
            if (status == SecurityValidationStatus.NotPerformed) return;
            _any = true;
            _invalid |= status == SecurityValidationStatus.Invalid;
            _indeterminate |= status == SecurityValidationStatus.Indeterminate;
        }

        internal SecurityValidationStatus Result =>
            !_any ? SecurityValidationStatus.NotPerformed
            : _invalid ? SecurityValidationStatus.Invalid
            : _indeterminate ? SecurityValidationStatus.Indeterminate
            : SecurityValidationStatus.Valid;

        internal static SecurityValidationStatus Combine(StatusAccumulator first, StatusAccumulator second) {
            var combined = new StatusAccumulator();
            combined.Add(first.Result);
            combined.Add(second.Result);
            return combined.Result;
        }
    }

    private sealed class CertificatePool {
        private readonly Dictionary<string, BcX509Certificate> _certificates = new(StringComparer.Ordinal);
        private readonly X509CertificateParser _parser = new();

        internal IEnumerable<BcX509Certificate> All => _certificates.Values;

        internal void AddRange(IEnumerable<BcX509Certificate> certificates) {
            foreach (BcX509Certificate certificate in certificates) {
                string identity = Identity(certificate);
                if (!_certificates.ContainsKey(identity)) _certificates[identity] = certificate;
            }
        }

        internal void TryAdd(byte[] encoded) {
            try {
                BcX509Certificate? certificate = _parser.ReadCertificate(encoded);
                if (certificate is not null) AddRange(new[] { certificate });
            } catch (Exception exception) when (IsValidationException(exception)) {
                // An unreadable certificate cannot be part of a verified path; path building reports the gap.
            }
        }
    }

    private sealed class EncodedSet {
        private readonly HashSet<string> _seen = new(StringComparer.Ordinal);
        internal List<byte[]> Values { get; } = new List<byte[]>();

        internal void Add(byte[] value) {
            if (_seen.Add(Convert.ToBase64String(Sha256(value)))) Values.Add(value);
        }
    }

    private readonly struct EmbeddedValue {
        internal EmbeddedValue(byte[] encoded, bool isOcsp) {
            Encoded = encoded;
            IsOcsp = isOcsp;
        }

        internal byte[] Encoded { get; }
        internal bool IsOcsp { get; }
    }

    private sealed class RevocationSubject {
        internal RevocationSubject(string role, BcX509Certificate? certificate, DateTime? time, bool isTimestampAuthority) {
            Role = role;
            Certificate = certificate;
            Time = time;
            IsTimestampAuthority = isTimestampAuthority;
        }

        internal string Role { get; }
        internal BcX509Certificate? Certificate { get; }
        internal DateTime? Time { get; }
        internal bool IsTimestampAuthority { get; }
    }

    private sealed class ParsedCms {
        internal List<BcX509Certificate> Certificates { get; } = new List<BcX509Certificate>();
        internal List<EmbeddedValue> EmbeddedTokens { get; } = new List<EmbeddedValue>();
        internal List<RevocationSubject> Subjects { get; } = new List<RevocationSubject>();
        internal bool HasSigningCertificateAttribute { get; private set; }

        /// <summary>True when an ESS signing-certificate attribute names a different certificate than the signer's.</summary>
        internal bool SigningCertificateMismatch { get; private set; }

        internal static ParsedCms Parse(byte[] encoded, bool isTimestampToken, IEnumerable<BcX509Certificate>? suppliedCertificates = null) {
            var parsed = new ParsedCms();
            BcX509Certificate[] supplied = suppliedCertificates?.ToArray() ?? Array.Empty<BcX509Certificate>();
            var signedData = new CmsSignedData(encoded);
            if (isTimestampToken) {
                var token = new TimeStampToken(signedData);
                parsed.AddContainer(signedData);
                parsed.Subjects.Add(new RevocationSubject("Time-stamp authority",
                    FindSigner(token.GetCertificates().EnumerateMatches(null).Concat(supplied), token.SignerID),
                    token.TimeStampInfo.GenTime.ToUniversalTime(), isTimestampAuthority: true));
                return parsed;
            }

            parsed.AddContainer(signedData);
            List<SignerInformation> signers = signedData.GetSignerInfos().GetSigners().ToList();
            if (signers.Count == 0) throw new InvalidDataException("The CMS container has no signer.");
            // Every signer is checked, so one revoked signer cannot hide behind another's good status.
            parsed.HasSigningCertificateAttribute = true;
            for (int signerIndex = 0; signerIndex < signers.Count; signerIndex++) {
                SignerInformation signer = signers[signerIndex];
                string role = signers.Count > 1 ? "Signer " + (signerIndex + 1) : "Signer";
                // Signer certificates may come from caller-supplied certificates, as in CMS verification.
                BcX509Certificate? signerCertificate = FindSigner(parsed.Certificates.Concat(supplied), signer.SignerID);
                parsed.Subjects.Add(new RevocationSubject(role, signerCertificate, null, isTimestampAuthority: false));

                Org.BouncyCastle.Asn1.Cms.AttributeTable? signedAttributes = signer.SignedAttributes;
                SigningCertificateBinding binding = CheckSigningCertificate(signedAttributes, signerCertificate);
                parsed.HasSigningCertificateAttribute &= binding == SigningCertificateBinding.Valid;
                parsed.SigningCertificateMismatch |= binding == SigningCertificateBinding.Mismatch;
                Org.BouncyCastle.Asn1.Cms.Attribute? archival = signedAttributes?[new DerObjectIdentifier(AdobeRevocationInfoArchivalOid)];
                if (archival is not null) parsed.AddRevocationInfoArchival(archival);

                Org.BouncyCastle.Asn1.Cms.Attribute? timestamps = signer.UnsignedAttributes?[new DerObjectIdentifier(SignatureTimestampTokenOid)];
                if (timestamps is null) continue;
                int index = 0;
                foreach (Asn1Encodable value in timestamps.AttrValues) {
                    index++;
                    try {
                        var tokenData = new CmsSignedData(Org.BouncyCastle.Asn1.Cms.ContentInfo.GetInstance(value));
                        var token = new TimeStampToken(tokenData);
                        parsed.AddContainer(tokenData);
                        parsed.Subjects.Add(new RevocationSubject(
                            role + " time-stamp authority" + (timestamps.AttrValues.Count > 1 ? " " + index : string.Empty),
                            FindSigner(token.GetCertificates().EnumerateMatches(null).Concat(supplied), token.SignerID),
                            token.TimeStampInfo.GenTime.ToUniversalTime(), isTimestampAuthority: true));
                    } catch (Exception exception) when (IsValidationException(exception)) {
                        // Time-stamp validation reports malformed tokens; there is no authority path to check.
                    }
                }
            }
            return parsed;
        }

        private static BcX509Certificate? FindSigner(IEnumerable<BcX509Certificate> certificates, SignerID signerId) =>
            certificates.FirstOrDefault(signerId.Match);

        /// <summary>
        /// RFC 5035: the first ESSCertID(v2) identifies the signer certificate by hash and, when present, by
        /// issuer and serial number. A different certificate means the attribute binds someone else.
        /// </summary>
        private static SigningCertificateBinding CheckSigningCertificate(
            Org.BouncyCastle.Asn1.Cms.AttributeTable? attributes,
            BcX509Certificate? signer) {
            if (attributes is null) return SigningCertificateBinding.Absent;
            Org.BouncyCastle.Asn1.Cms.Attribute? v2 = attributes[new DerObjectIdentifier(SigningCertificateV2Oid)];
            Org.BouncyCastle.Asn1.Cms.Attribute? v1 = attributes[new DerObjectIdentifier(SigningCertificateOid)];
            if (v2 is null && v1 is null) return SigningCertificateBinding.Absent;
            if (signer is null) return SigningCertificateBinding.Absent;
            try {
                byte[] encoded = signer.GetEncoded();
                byte[] expected;
                byte[] actual;
                Org.BouncyCastle.Asn1.X509.IssuerSerial? issuerSerial;
                if (v2 is not null) {
                    Org.BouncyCastle.Asn1.Ess.EssCertIDv2 id = Org.BouncyCastle.Asn1.Ess.SigningCertificateV2
                        .GetInstance(v2.AttrValues[0]).GetCerts()[0];
                    string algorithm = id.HashAlgorithm?.Algorithm.Id ?? "2.16.840.1.101.3.4.2.1";
                    expected = id.GetCertHash();
                    actual = Org.BouncyCastle.Security.DigestUtilities.CalculateDigest(algorithm, encoded);
                    issuerSerial = id.IssuerSerial;
                } else {
                    Org.BouncyCastle.Asn1.Ess.EssCertID id = Org.BouncyCastle.Asn1.Ess.SigningCertificate
                        .GetInstance(v1!.AttrValues[0]).GetCerts()[0];
                    expected = id.GetCertHash();
                    actual = Org.BouncyCastle.Security.DigestUtilities.CalculateDigest("1.3.14.3.2.26", encoded);
                    issuerSerial = id.IssuerSerial;
                }
                if (!Org.BouncyCastle.Utilities.Arrays.FixedTimeEquals(expected, actual)) return SigningCertificateBinding.Mismatch;
                if (issuerSerial is not null && !issuerSerial.Serial.Value.Equals(signer.SerialNumber)) {
                    return SigningCertificateBinding.Mismatch;
                }
                return SigningCertificateBinding.Valid;
            } catch (Exception exception) when (IsValidationException(exception)) {
                return SigningCertificateBinding.Mismatch;
            }
        }

        private void AddContainer(CmsSignedData signedData) {
            Certificates.AddRange(signedData.GetCertificates().EnumerateMatches(null));
            Asn1Set? revocationInfo = SignedData.GetInstance(signedData.ContentInfo.Content).CRLs;
            if (revocationInfo is null) return;
            foreach (Asn1Encodable item in revocationInfo) {
                if (item is Asn1TaggedObject tagged) {
                    // RFC 5940 OtherRevocationInfoFormat carrying an OCSPResponse.
                    if (tagged.TagNo != 1) continue;
                    var other = OtherRevocationInfoFormat.GetInstance(tagged, false);
                    if (other.InfoFormat.Id == OcspResponseRevocationInfoFormatOid) {
                        EmbeddedTokens.Add(new EmbeddedValue(other.Info.ToAsn1Object().GetEncoded(Asn1Encodable.Der), isOcsp: true));
                    }
                    continue;
                }
                EmbeddedTokens.Add(new EmbeddedValue(item.ToAsn1Object().GetEncoded(Asn1Encodable.Der), isOcsp: false));
            }
        }

        private void AddRevocationInfoArchival(Org.BouncyCastle.Asn1.Cms.Attribute attribute) {
            // Adobe RevocationInfoArchival ::= SEQUENCE { crl [0] EXPLICIT SEQUENCE OF CRLs OPTIONAL,
            //   ocsp [1] EXPLICIT SEQUENCE OF OCSPResponse OPTIONAL, otherRevInfo [2] ... OPTIONAL }
            foreach (Asn1Encodable value in attribute.AttrValues) {
                if (value is not Asn1Sequence archival) continue;
                foreach (Asn1Encodable element in archival) {
                    if (element is not Asn1TaggedObject tagged || tagged.TagNo > 1) continue;
                    if (tagged.GetExplicitBaseObject() is not Asn1Sequence values) continue;
                    foreach (Asn1Encodable item in values) {
                        EmbeddedTokens.Add(new EmbeddedValue(item.ToAsn1Object().GetEncoded(Asn1Encodable.Der), tagged.TagNo == 1));
                    }
                }
            }
        }
    }
}
