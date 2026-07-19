using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;

namespace OfficeIMO.Security;

internal static class CertificateChainValidator {
    internal static CertificateValidationResult Validate(
        X509Certificate2? certificate,
        IEnumerable<X509Certificate2> embeddedCertificates,
        CertificateValidationOptions options,
        IList<SecurityFinding> findings,
        string role,
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

        bool accepted = options.ChainEvaluator?.Invoke(certificate, chain) ?? platformResult;
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
