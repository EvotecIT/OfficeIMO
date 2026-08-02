using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security;

internal enum CertificateUsagePurpose {
    CmsSigner,
    DocumentSigner,
    TimestampAuthority
}

internal static class CertificateChainValidator {
    internal static CertificateValidationResult Validate(
        X509Certificate2? certificate,
        IEnumerable<X509Certificate2> embeddedCertificates,
        CertificateValidationOptions options,
        IList<SecurityFinding> findings,
        string role,
        CertificateUsagePurpose purpose,
        int? signerIndex = null) {
        if (certificate == null) {
            return Empty(SecurityValidationStatus.Indeterminate);
        }
        bool usageAccepted = ValidateCertificateUsage(certificate, purpose, findings, role, signerIndex);
        if (!options.ValidateChain) {
            return Empty(usageAccepted
                ? SecurityValidationStatus.NotPerformed
                : SecurityValidationStatus.Invalid);
        }

        IReadOnlyList<X509Certificate2> embedded = embeddedCertificates as IReadOnlyList<X509Certificate2>
            ?? embeddedCertificates.ToArray();

        using var chain = new X509Chain();
        chain.ChainPolicy.RevocationMode = options.RevocationMode;
        chain.ChainPolicy.RevocationFlag = options.RevocationFlag;
        chain.ChainPolicy.VerificationFlags = options.VerificationFlags;
        if (!TryApplyCertificateDownloadPolicy(chain.ChainPolicy, options.DisableCertificateDownloads)) {
            IEnumerable<X509Certificate2> offlineCandidates = embedded
                .Concat(options.ExtraCertificates.Cast<X509Certificate2>());
            if (!HasCompleteOfflinePath(certificate, offlineCandidates)) {
                findings.Add(new SecurityFinding(
                    SecurityFindingSeverity.Warning,
                    "CertificateDownloadPolicyUnavailable",
                    role + " certificate chain was not built because this runtime cannot enforce the requested no-download policy and the supplied certificates do not contain a complete verified issuer path.",
                    signerIndex));
                return Empty(usageAccepted
                    ? SecurityValidationStatus.Indeterminate
                    : SecurityValidationStatus.Invalid);
            }
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Info,
                "CertificateDownloadPolicyOfflineFallback",
                role + " certificate chain uses the complete cryptographically verified issuer path supplied by the caller or signed object because this runtime cannot set DisableCertificateDownloads.",
                signerIndex));
        }
        chain.ChainPolicy.UrlRetrievalTimeout = options.UrlRetrievalTimeout;
        if (options.VerificationTime.HasValue) {
            chain.ChainPolicy.VerificationTime = options.VerificationTime.Value;
        }

        foreach (X509Certificate2 candidate in embedded) {
            if (!string.Equals(candidate.Thumbprint, certificate.Thumbprint, StringComparison.OrdinalIgnoreCase)) {
                chain.ChainPolicy.ExtraStore.Add(candidate);
            }
        }
        chain.ChainPolicy.ExtraStore.AddRange(options.ExtraCertificates);

        bool platformResult;
        try {
            platformResult = chain.Build(certificate);
        } catch (Exception exception) when (exception is CryptographicException or ArgumentException) {
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Warning,
                "CertificateChainFailed",
                role + " certificate chain could not be built: " + exception.Message,
                signerIndex));
            return Empty(usageAccepted
                ? SecurityValidationStatus.Indeterminate
                : SecurityValidationStatus.Invalid);
        }

        bool chainAccepted = options.ChainEvaluator?.Invoke(certificate, chain) ?? platformResult;
        bool accepted = chainAccepted && usageAccepted;
        string[] statuses = chain.ChainStatus
            .Select(static status => string.IsNullOrWhiteSpace(status.StatusInformation)
                ? status.Status.ToString()
                : status.Status + ": " + status.StatusInformation.Trim())
            .ToArray();
        if (!accepted) {
            string statusText = statuses.Length == 0 ? "no platform chain status" : string.Join(", ", statuses);
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Warning,
                "CertificateChainUntrusted",
                role + " certificate chain was not accepted: " + statusText + ".",
                signerIndex));
        }

        return new CertificateValidationResult(
            accepted ? SecurityValidationStatus.Valid : SecurityValidationStatus.Invalid,
            ClassifyRevocation(chain, options.RevocationMode),
            statuses);
    }

    private static bool ValidateCertificateUsage(
        X509Certificate2 certificate,
        CertificateUsagePurpose purpose,
        IList<SecurityFinding> findings,
        string role,
        int? signerIndex) {
        X509KeyUsageExtension? keyUsage = certificate.Extensions
            .OfType<X509KeyUsageExtension>()
            .FirstOrDefault();
        if (keyUsage != null &&
            (keyUsage.KeyUsages & (X509KeyUsageFlags.DigitalSignature | X509KeyUsageFlags.NonRepudiation)) == 0) {
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Error,
                "CertificateKeyUsageInvalid",
                role + " certificate key usage does not permit digital signatures.",
                signerIndex));
            return false;
        }

        X509EnhancedKeyUsageExtension? enhancedKeyUsage = certificate.Extensions
            .OfType<X509EnhancedKeyUsageExtension>()
            .FirstOrDefault();
        if (enhancedKeyUsage == null) {
            if (purpose == CertificateUsagePurpose.TimestampAuthority) {
                findings.Add(new SecurityFinding(
                    SecurityFindingSeverity.Error,
                    "CertificateEnhancedKeyUsageInvalid",
                    role + " certificate does not declare the timestamping enhanced key usage.",
                    signerIndex));
                return false;
            }
            return true;
        }

        if (purpose == CertificateUsagePurpose.TimestampAuthority &&
            (!enhancedKeyUsage.Critical || enhancedKeyUsage.EnhancedKeyUsages.Count != 1)) {
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Error,
                "CertificateEnhancedKeyUsageInvalid",
                role + " certificate must declare only the critical timestamping enhanced key usage.",
                signerIndex));
            return false;
        }

        bool permitted = enhancedKeyUsage.EnhancedKeyUsages
            .Cast<Oid>()
            .Any(oid => IsPermittedEnhancedKeyUsage(oid.Value, purpose));
        if (!permitted) {
            findings.Add(new SecurityFinding(
                SecurityFindingSeverity.Error,
                "CertificateEnhancedKeyUsageInvalid",
                role + " certificate enhanced key usage is not valid for " +
                    (purpose == CertificateUsagePurpose.TimestampAuthority
                        ? "timestamping."
                        : purpose == CertificateUsagePurpose.DocumentSigner
                            ? "document signing."
                            : "CMS signing."),
                signerIndex));
        }
        return permitted;
    }

    private static bool IsPermittedEnhancedKeyUsage(string? oid, CertificateUsagePurpose purpose) {
        if (purpose == CertificateUsagePurpose.TimestampAuthority) {
            return string.Equals(oid, "1.3.6.1.5.5.7.3.8", StringComparison.Ordinal);
        }

        if (purpose == CertificateUsagePurpose.DocumentSigner) {
            return oid is "2.5.29.37.0" or
                "1.3.6.1.5.5.7.3.3" or
                "1.3.6.1.4.1.311.10.3.12";
        }

        return oid is "2.5.29.37.0" or
            "1.3.6.1.5.5.7.3.3" or
            "1.3.6.1.5.5.7.3.4" or
            "1.3.6.1.5.5.7.3.36" or
            "1.3.6.1.4.1.311.10.3.12";
    }

    private static CertificateValidationResult Empty(SecurityValidationStatus chainStatus) =>
        new CertificateValidationResult(
            chainStatus,
            SecurityValidationStatus.NotPerformed,
            Array.Empty<string>());

    private static bool TryApplyCertificateDownloadPolicy(
        X509ChainPolicy chainPolicy,
        bool disableCertificateDownloads) {
#if NETSTANDARD2_0 || NETFRAMEWORK
        System.Reflection.PropertyInfo? property = typeof(X509ChainPolicy).GetProperty("DisableCertificateDownloads");
        if (property == null || !property.CanWrite) return !disableCertificateDownloads;
        try {
            property.SetValue(chainPolicy, disableCertificateDownloads, null);
            return true;
        } catch (Exception exception) when (exception is ArgumentException or System.Reflection.TargetInvocationException) {
            return !disableCertificateDownloads;
        }
#else
        chainPolicy.DisableCertificateDownloads = disableCertificateDownloads;
        return true;
#endif
    }

    internal static bool HasCompleteOfflinePath(
        X509Certificate2 certificate,
        IEnumerable<X509Certificate2> issuerCandidates) {
        var certificates = issuerCandidates
            .Prepend(certificate)
            .GroupBy(static candidate => candidate.Thumbprint ?? Convert.ToBase64String(candidate.GetCertHash()),
                StringComparer.OrdinalIgnoreCase)
            .Select(static group => group.First())
            .ToArray();
        return HasCompleteOfflinePath(
            certificate,
            certificates,
            new HashSet<string>(StringComparer.OrdinalIgnoreCase));
    }

    private static bool HasCompleteOfflinePath(
        X509Certificate2 certificate,
        IReadOnlyList<X509Certificate2> candidates,
        HashSet<string> path) {
        string identity = certificate.Thumbprint ?? Convert.ToBase64String(certificate.GetCertHash());
        if (!path.Add(identity)) return false;
        try {
            if (certificate.IssuerName.RawData.SequenceEqual(certificate.SubjectName.RawData)) {
                return VerifyCertificateSignature(certificate, certificate);
            }
            foreach (X509Certificate2 issuer in candidates) {
                string issuerIdentity = issuer.Thumbprint ?? Convert.ToBase64String(issuer.GetCertHash());
                if (string.Equals(identity, issuerIdentity, StringComparison.OrdinalIgnoreCase) ||
                    !certificate.IssuerName.RawData.SequenceEqual(issuer.SubjectName.RawData) ||
                    !VerifyCertificateSignature(certificate, issuer)) {
                    continue;
                }
                if (HasCompleteOfflinePath(issuer, candidates, path)) return true;
            }
            return false;
        } finally {
            path.Remove(identity);
        }
    }

    private static bool VerifyCertificateSignature(
        X509Certificate2 certificate,
        X509Certificate2 issuer) {
        try {
            Org.BouncyCastle.X509.X509Certificate candidate =
                Org.BouncyCastle.Security.DotNetUtilities.FromX509Certificate(certificate);
            Org.BouncyCastle.X509.X509Certificate issuerCertificate =
                Org.BouncyCastle.Security.DotNetUtilities.FromX509Certificate(issuer);
            candidate.Verify(issuerCertificate.GetPublicKey());
            return true;
        } catch (Org.BouncyCastle.Security.GeneralSecurityException) {
            return false;
        }
    }

    private static SecurityValidationStatus ClassifyRevocation(X509Chain chain, X509RevocationMode revocationMode) {
        if (revocationMode == X509RevocationMode.NoCheck) return SecurityValidationStatus.NotPerformed;
        bool indeterminate = false;
        foreach (X509ChainStatus status in chain.ChainStatus) {
            if ((status.Status & X509ChainStatusFlags.Revoked) != 0) return SecurityValidationStatus.Invalid;
            if ((status.Status & (X509ChainStatusFlags.RevocationStatusUnknown |
                                  X509ChainStatusFlags.OfflineRevocation)) != 0) {
                indeterminate = true;
            }
        }
        return indeterminate ? SecurityValidationStatus.Indeterminate : SecurityValidationStatus.Valid;
    }
}
