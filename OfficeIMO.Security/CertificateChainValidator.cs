using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security;

internal enum CertificateUsagePurpose {
    CmsSigner,
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
        if (!options.ValidateChain) {
            return Empty(SecurityValidationStatus.NotPerformed);
        }

        using var chain = new X509Chain();
        chain.ChainPolicy.RevocationMode = options.RevocationMode;
        chain.ChainPolicy.RevocationFlag = options.RevocationFlag;
        chain.ChainPolicy.VerificationFlags = options.VerificationFlags;
        chain.ChainPolicy.UrlRetrievalTimeout = options.UrlRetrievalTimeout;
        if (options.VerificationTime.HasValue) {
            chain.ChainPolicy.VerificationTime = options.VerificationTime.Value;
        }

        foreach (X509Certificate2 candidate in embeddedCertificates) {
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
            return Empty(SecurityValidationStatus.Indeterminate);
        }

        bool chainAccepted = options.ChainEvaluator?.Invoke(certificate, chain) ?? platformResult;
        bool usageAccepted = ValidateCertificateUsage(certificate, purpose, findings, role, signerIndex);
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
                    (purpose == CertificateUsagePurpose.TimestampAuthority ? "timestamping." : "document or CMS signing."),
                signerIndex));
        }
        return permitted;
    }

    private static bool IsPermittedEnhancedKeyUsage(string? oid, CertificateUsagePurpose purpose) {
        if (purpose == CertificateUsagePurpose.TimestampAuthority) {
            return string.Equals(oid, "1.3.6.1.5.5.7.3.8", StringComparison.Ordinal);
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
