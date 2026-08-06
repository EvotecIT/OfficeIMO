using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Security;

public static partial class OfficeVbaSignatureService {
    /// <summary>Creates legacy, agile, and V3 VBA signatures with managed MS-OVBA canonicalization and provider-backed CMS.</summary>
    public static OfficeVbaSigningResult TrySign(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficeVbaSigningOptions? options = null,
        IEnumerable<X509Certificate2>? certificateChain = null) {
        if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
        if (signingCertificate == null) throw new ArgumentNullException(nameof(signingCertificate));
        options ??= new OfficeVbaSigningOptions();
        ValidateOptions(options);
        ValidateSigningOptions(options);
        string fullPath = NormalizePath(filePath);
        var findings = new List<OfficeVbaSignatureFinding>();
        OfficeVbaSignatureInfo source = Inspect(fullPath, options);
        if (!source.IsMacroEnabledFormat || !source.HasMacroProject || string.IsNullOrWhiteSpace(source.MacroProjectUri)
            || source.Findings.Any(finding => finding.State == OfficePackageSignatureValidationState.Failed)) {
            findings.AddRange(source.Findings);
            findings.Add(Finding("VbaSigningPreflightFailed", OfficePackageSignatureValidationState.Failed,
                "VBA signing requires a valid macro-enabled package with one bounded vbaProject.bin."));
            return SigningResult(fullPath, true, false, null, findings);
        }
        if (!signingCertificate.HasPrivateKey) {
            findings.Add(Finding("VbaSigningPrivateKeyMissing", OfficePackageSignatureValidationState.Failed,
                "The signing certificate has no accessible private key."));
            return SigningResult(fullPath, true, false, null, findings);
        }
        OfficePackageSignatureInfo packageSignatures = OfficePackageSignatureService.Inspect(fullPath,
            new OfficePackageSignatureInspectionOptions { VerifyDigests = false });
        if (packageSignatures.HasSignatures && !options.AllowPackageSignatureInvalidation) {
            findings.Add(Finding("ExistingPackageSignatureInvalidationBlocked", OfficePackageSignatureValidationState.Failed,
                "VBA signing would invalidate existing OPC package signatures."));
            return SigningResult(fullPath, true, false, null, findings);
        }
        if (!TryReadVbaProject(fullPath, source.MacroProjectUri!, options,
                out byte[] projectBytes, out string readDetail)) {
            findings.Add(Finding("VbaProjectReadFailed", OfficePackageSignatureValidationState.Failed, readDetail));
            return SigningResult(fullPath, true, false, null, findings);
        }
        if (!OfficeVbaProjectCanonicalizer.TryCreate(projectBytes, options.MaxMacroProjectBytes,
                out OfficeVbaProjectCanonicalizer.Result? canonical, out string canonicalDetail)
            || canonical == null) {
            findings.Add(Finding("VbaManagedCanonicalizationFailed", OfficePackageSignatureValidationState.Failed,
                canonicalDetail));
            return SigningResult(fullPath, true, false, null, findings);
        }

        var profileParts = new Dictionary<OfficeVbaSignatureProfile, byte[]>();
        try {
            foreach (OfficeVbaSignatureProfile profile in new[] {
                         OfficeVbaSignatureProfile.Legacy,
                         OfficeVbaSignatureProfile.Agile,
                         OfficeVbaSignatureProfile.V3 }) {
                byte[] contentHash = profile switch {
                    OfficeVbaSignatureProfile.Legacy => canonical.ComputeLegacyHash(),
                    OfficeVbaSignatureProfile.Agile => canonical.ComputeAgileHash(),
                    OfficeVbaSignatureProfile.V3 => canonical.ComputeV3Hash(),
                    _ => throw new ArgumentOutOfRangeException(nameof(profile))
                };
                byte[] signedContent = OfficeVbaSignatureEncoding.CreateSignedContent(profile, contentHash);
                CmsSigningOptions cmsOptions = CopyCmsSigningOptions(options.CmsSigning);
                cmsOptions.ContentTypeOid = OfficeVbaSignatureEncoding.AuthenticodeContentTypeOid;
                byte[] cms = securityProvider.SignCmsEncapsulated(
                    signedContent, signingCertificate, cmsOptions, certificateChain);
                profileParts[profile] = OfficeVbaSignatureEncoding.CreateDigSigInfoSerialized(
                    cms, signingCertificate.RawData);
            }
        } catch (Exception exception) when (exception is CryptographicException or InvalidOperationException
            or NotSupportedException or ArgumentException or OverflowException) {
            findings.Add(Finding("VbaCmsCreationFailed", OfficePackageSignatureValidationState.Failed,
                "The security provider could not create all VBA CMS profiles. " + exception.Message));
            return SigningResult(fullPath, true, false, null, findings);
        }

        string stagingPath = string.Empty;
        try {
            stagingPath = OfficeFileCommit.CreateStagingPath(fullPath);
            OfficePackageFileSnapshot.CopyBounded(fullPath, stagingPath, options.Package.MaxPackageBytes);
            string sourceHash = OfficePackageFileSnapshot.ComputeSha256(stagingPath, options.Package.MaxPackageBytes);
            OfficeVbaPackageSignatureWriter.Write(stagingPath, source.MacroProjectUri!, profileParts);

            OfficeVbaSignatureValidationResult validation = Validate(stagingPath, securityProvider, options);
            bool hasAllProfiles = validation.SignatureInfo.Signatures
                .Select(signature => signature.Profile).Distinct().Count() == 3;
            if (!hasAllProfiles || !validation.IsValidUnderPolicy) {
                findings.AddRange(validation.Findings);
                findings.Add(Finding("VbaSignatureReadbackFailed", OfficePackageSignatureValidationState.Failed,
                    "The managed VBA signatures did not satisfy profile, content-binding, CMS, and trust policy."));
                return SigningResult(fullPath, true, false, validation, findings);
            }
            string validatedHash = OfficePackageFileSnapshot.ComputeSha256(stagingPath, options.Package.MaxPackageBytes);
            if (!OfficeFileCommit.TryCommitTemporaryFileAtomicallyIfDestinationUnchanged(
                    stagingPath, fullPath,
                    displaced => string.Equals(sourceHash,
                        OfficePackageFileSnapshot.ComputeSha256(displaced, options.Package.MaxPackageBytes), StringComparison.Ordinal),
                    installed => string.Equals(validatedHash,
                        OfficePackageFileSnapshot.ComputeSha256(installed, options.Package.MaxPackageBytes), StringComparison.Ordinal))) {
                stagingPath = string.Empty;
                findings.Add(Finding("SourcePackageChangedDuringSigning", OfficePackageSignatureValidationState.Failed,
                    "The package changed while VBA signatures were staged; the current source was preserved."));
                return SigningResult(fullPath, true, false, validation, findings);
            }
            stagingPath = string.Empty;
            findings.Add(Finding("VbaSignaturesCommitted", OfficePackageSignatureValidationState.Passed,
                "Managed legacy, agile, and V3 VBA signatures were validated and atomically committed."));
            return SigningResult(fullPath, true, true, validation, findings);
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or
            InvalidDataException or ArgumentException or OverflowException or CryptographicException) {
            findings.Add(Finding("VbaSigningFailed", OfficePackageSignatureValidationState.Failed,
                "Managed VBA signing failed before atomic commit. " + exception.Message));
            return SigningResult(fullPath, true, false, null, findings);
        } finally {
            if (!string.IsNullOrWhiteSpace(stagingPath)) OfficeFileCommit.DeleteIfExists(stagingPath);
        }
    }

    /// <summary>Creates portable managed VBA profiles and throws unless validated atomic signing completes.</summary>
    public static OfficeVbaSigningResult Sign(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficeVbaSigningOptions? options = null,
        IEnumerable<X509Certificate2>? certificateChain = null) {
        OfficeVbaSigningResult result = TrySign(
            filePath, securityProvider, signingCertificate, options, certificateChain);
        if (!result.Succeeded) {
            throw new InvalidOperationException(string.Join(" ", result.Findings.Select(finding => finding.Message)));
        }
        return result;
    }

    private static void ValidateSigningOptions(OfficeVbaSigningOptions options) {
        if (options.CmsSigning.MaxContentBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(options), "CMS content byte limit must be positive.");
        }
        if (options.CmsSigning.DigestAlgorithm != HashAlgorithmName.SHA256) {
            throw new NotSupportedException("Managed VBA signature creation currently requires SHA-256 CMS signatures.");
        }
    }

    private static CmsSigningOptions CopyCmsSigningOptions(CmsSigningOptions source) => new CmsSigningOptions {
        DigestAlgorithm = source.DigestAlgorithm,
        ContentTypeOid = source.ContentTypeOid,
        IncludeSigningTime = source.IncludeSigningTime,
        SigningTime = source.SigningTime,
        IncludeCertificateChain = source.IncludeCertificateChain,
        MaxContentBytes = source.MaxContentBytes
    };

    private static OfficeVbaSigningResult SigningResult(string path, bool supported, bool succeeded,
        OfficeVbaSignatureValidationResult? validation, IReadOnlyList<OfficeVbaSignatureFinding> findings) =>
        new(path, supported, succeeded, validation, findings.ToArray());
}
