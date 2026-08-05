using System.Security.Cryptography.X509Certificates;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Core.Internal;
using OfficeIMO.Security;

namespace OfficeIMO.Word {
    internal static class WordMacroProjectSignatureService {
        internal static WordMacroProjectSignatureValidationResult Validate(
            string filePath,
            IOfficeSecurityProvider securityProvider,
            WordMacroProjectSignatureValidationOptions? options = null,
            WordMacroProjectSigningDependencies? dependencies = null) {
            if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
            options ??= new WordMacroProjectSignatureValidationOptions();
            dependencies ??= WordMacroProjectSigningDependencies.Default;
            ValidateValidationOptions(options);
            using var inspectionBudget = new WordMacroProjectSignatureInspector.InspectionBudget(
                options.Inspection,
                securityProvider);

            string fullPath = NormalizePath(filePath);
            if (!File.Exists(fullPath)) {
                return ValidateSnapshot(fullPath, fullPath, options, dependencies, inspectionBudget);
            }
            string snapshotDirectory = Path.Combine(
                Path.GetTempPath(),
                "OfficeIMO-Word-MacroValidation-" + Guid.NewGuid().ToString("N"));
            string snapshotPath = Path.Combine(snapshotDirectory, "snapshot" + Path.GetExtension(fullPath));
            try {
                try {
                    Directory.CreateDirectory(snapshotDirectory);
                    WordPackageSnapshot.CopyBounded(fullPath, snapshotPath, options.Inspection.PackageSecurity.MaxPackageBytes);
                } catch (Exception exception) when (exception is IOException || exception is UnauthorizedAccessException || exception is InvalidDataException) {
                    WordMacroProjectSignatureInfo info = WordMacroProjectSignatureInspector.Inspect(
                        fullPath, options.Inspection, inspectionBudget);
                    var findings = CollectInspectionFindings(info);
                    if (!findings.Any(finding => finding.State == WordSignatureValidationState.Failed)) {
                        findings.Add(Finding("MacroValidationSnapshotFailed", WordSignatureValidationState.Failed,
                            "A bounded immutable validation snapshot could not be created. " + exception.Message));
                    }
                    return Result(info, dependencies.Platform.IsWindows,
                        WordSignatureValidationState.NotChecked, findings, options);
                }
                return ValidateSnapshot(fullPath, snapshotPath, options, dependencies, inspectionBudget);
            } finally {
                OfficeFileCommit.DeleteIfExists(snapshotPath);
                try {
                    if (Directory.Exists(snapshotDirectory)) Directory.Delete(snapshotDirectory);
                } catch (IOException) {
                    // The bounded snapshot file was already removed; cleanup races are non-fatal.
                } catch (UnauthorizedAccessException) {
                    // A transient scanner handle must not replace the structured validation result.
                }
            }
        }

        private static WordMacroProjectSignatureValidationResult ValidateSnapshot(
            string originalPath,
            string snapshotPath,
            WordMacroProjectSignatureValidationOptions options,
            WordMacroProjectSigningDependencies dependencies,
            WordMacroProjectSignatureInspector.InspectionBudget inspectionBudget) {
            WordMacroProjectSignatureInfo snapshotInfo =
                WordMacroProjectSignatureInspector.Inspect(
                    snapshotPath, options.Inspection, inspectionBudget);
            WordMacroProjectSignatureInfo info = WithFilePath(snapshotInfo, originalPath);

            var findings = CollectInspectionFindings(info);
            if (!CanValidate(info, findings)) {
                return Result(info, dependencies.Platform.IsWindows,
                    WordSignatureValidationState.NotChecked, findings, options);
            }
            if (!dependencies.Platform.IsWindows) {
                findings.Add(Finding("MacroSigningPlatformUnsupported", WordSignatureValidationState.Unsupported,
                    "Native VBA content-binding validation requires Windows and Microsoft's registered Office SIP."));
                return Result(info, false, WordSignatureValidationState.Unsupported, findings, options);
            }
            if (!dependencies.Platform.TryGetSubjectInterfacePackage(snapshotPath, out _, out string sipDetail)) {
                findings.Add(Finding("MicrosoftOfficeSipUnavailable", WordSignatureValidationState.Unsupported, sipDetail));
                return Result(info, false, WordSignatureValidationState.Unsupported, findings, options);
            }
            if (options.RequireV3Signature && !info.HasV3Signature) {
                findings.Add(Finding("MacroV3SignatureRequired", WordSignatureValidationState.Failed,
                    "The current V3 VBA signature profile is required by policy."));
                return Result(info, true, WordSignatureValidationState.NotChecked, findings, options);
            }

            WordMacroProjectSignaturePartInfo? selected = SelectHighestProfile(info);
            byte[]? expectedDigest = selected?.SignedContentDigest;
            if (selected == null || string.IsNullOrWhiteSpace(selected.SignedContentDigestAlgorithmOid) ||
                expectedDigest == null) {
                findings.Add(Finding("MacroSignedContentDigestMissing", WordSignatureValidationState.Failed,
                    "The selected VBA signature does not contain a decodable Authenticode subject digest.",
                    selected?.Profile));
                return Result(info, true, WordSignatureValidationState.NotChecked, findings, options);
            }
            WordMacroProjectContentBindingResult binding = dependencies.Platform.ValidateContentBinding(
                snapshotPath, selected.SignedContentDigestAlgorithmOid!, expectedDigest);
            if (!binding.IsSupported) {
                findings.Add(Finding("MacroContentBindingUnsupported", WordSignatureValidationState.Unsupported,
                    binding.Detail, selected.Profile));
                return Result(info, false, WordSignatureValidationState.Unsupported, findings, options);
            }
            if (!binding.IsValid) {
                findings.Add(Finding("MacroContentBindingInvalid", WordSignatureValidationState.Failed,
                    binding.Detail, selected.Profile));
                return Result(info, true, WordSignatureValidationState.Failed, findings, options);
            }
            findings.Add(Finding("MacroContentBindingValid", WordSignatureValidationState.Passed,
                binding.Detail, selected.Profile));
            return Result(info, true, WordSignatureValidationState.Passed, findings, options);
        }

        private static WordMacroProjectSignatureInfo WithFilePath(
            WordMacroProjectSignatureInfo source,
            string filePath) =>
            new WordMacroProjectSignatureInfo(
                filePath,
                source.IsMacroEnabledFormat,
                source.HasMacroProject,
                source.MacroProjectPartUri,
                source.MacroProjectLength,
                source.MacroProjectSha256,
                source.Signatures,
                source.Findings);

        internal static WordMacroProjectSigningResult TrySign(
            string filePath,
            IOfficeSecurityProvider securityProvider,
            string certificateThumbprint,
            WordMacroProjectSigningOptions? options = null,
            WordMacroProjectSigningDependencies? dependencies = null) {
            if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
            options ??= new WordMacroProjectSigningOptions();
            dependencies ??= WordMacroProjectSigningDependencies.Default;
            ValidateSigningOptions(options);

            string fullPath = NormalizePath(filePath);
            using var inspectionBudget = new WordMacroProjectSignatureInspector.InspectionBudget(
                options.Inspection,
                securityProvider);
            WordMacroProjectSignatureInfo sourceInfo = WordMacroProjectSignatureInspector.Inspect(
                fullPath, options.Inspection, inspectionBudget, validateCmsOverride: false);
            var findings = new List<WordMacroProjectSignatureFinding>();
            findings.AddRange(sourceInfo.Findings);
            if (!sourceInfo.IsMacroEnabledFormat || !sourceInfo.HasMacroProject ||
                string.IsNullOrWhiteSpace(sourceInfo.MacroProjectSha256)) {
                findings.Add(Finding("MacroProjectSigningPreflightFailed", WordSignatureValidationState.Failed,
                    "VBA signing requires a readable saved DOCM or DOTM package containing a bounded VBA project."));
                return SigningResult(fullPath, dependencies.Platform.IsWindows, false, false, null, findings);
            }
            if (!TryNormalizeThumbprint(certificateThumbprint, out string thumbprint, out string thumbprintDetail)) {
                findings.Add(Finding("CertificateThumbprintInvalid", WordSignatureValidationState.Failed, thumbprintDetail));
                return SigningResult(fullPath, dependencies.Platform.IsWindows, false, false, null, findings);
            }
            if (options.TimestampAuthorityUrl != null && !IsHttpTimestampUrl(options.TimestampAuthorityUrl)) {
                findings.Add(Finding("TimestampAuthorityUrlInvalid", WordSignatureValidationState.Failed,
                    "The RFC 3161 timestamp-authority URL must be an absolute HTTP or HTTPS URL."));
                return SigningResult(fullPath, dependencies.Platform.IsWindows, false, false, null, findings);
            }
            if (!dependencies.Platform.IsWindows) {
                findings.Add(Finding("MacroSigningPlatformUnsupported", WordSignatureValidationState.Unsupported,
                    "VBA signature creation requires Windows, SignTool, and Microsoft's registered Office SIP."));
                return SigningResult(fullPath, false, false, false, null, findings);
            }
            if (!TryResolveSignTool(options.SignToolPath, out string signTool, out string signToolDetail)) {
                findings.Add(Finding("SignToolUnavailable", WordSignatureValidationState.Unsupported, signToolDetail));
                return SigningResult(fullPath, false, false, false, null, findings);
            }
            if (!TryResolveClearSignatureTool(options.OfficeSipsDirectory, out string clearTool, out string clearDetail)) {
                findings.Add(Finding("OfficeSignatureClearToolUnavailable", WordSignatureValidationState.Unsupported, clearDetail));
                return SigningResult(fullPath, false, false, false, null, findings);
            }
            if (!dependencies.Platform.TryGetSubjectInterfacePackage(fullPath, out _, out string sipDetail)) {
                findings.Add(Finding("MicrosoftOfficeSipUnavailable", WordSignatureValidationState.Unsupported, sipDetail));
                return SigningResult(fullPath, false, false, false, null, findings);
            }

            string stagingPath = OfficeFileCommit.CreateStagingPath(fullPath);
            string validationSnapshotPath = string.Empty;
            try {
                WordPackageSnapshot.CopyBounded(
                    fullPath,
                    stagingPath,
                    options.Inspection.PackageSecurity.MaxPackageBytes);
                WordMacroProjectSignatureInfo stagingPreflight;
                using (var stagingPreflightBudget = new WordMacroProjectSignatureInspector.InspectionBudget(
                    options.Inspection,
                    securityProvider)) {
                    stagingPreflight = WordMacroProjectSignatureInspector.Inspect(
                        stagingPath,
                        options.Inspection,
                        stagingPreflightBudget,
                        validateCmsOverride: false);
                }
                foreach (WordMacroProjectSignatureFinding finding in stagingPreflight.Findings) {
                    if (!findings.Any(existing =>
                        existing.Code == finding.Code &&
                        existing.Profile == finding.Profile &&
                        existing.Message == finding.Message)) {
                        findings.Add(finding);
                    }
                }
                if (stagingPreflight.Findings.Any(finding => finding.State == WordSignatureValidationState.Failed) ||
                    !stagingPreflight.HasMacroProject ||
                    !string.Equals(sourceInfo.MacroProjectSha256, stagingPreflight.MacroProjectSha256,
                        StringComparison.OrdinalIgnoreCase)) {
                    findings.Add(Finding("MacroSigningStagingPreflightFailed", WordSignatureValidationState.Failed,
                        "The exact bounded staging snapshot changed or failed package-security inspection before native signing tools were invoked."));
                    return SigningResult(fullPath, true, false, false, null, findings);
                }
                string sourcePackageHash = WordPackageSnapshot.ComputeSha256(
                    stagingPath,
                    options.Inspection.PackageSecurity.MaxPackageBytes);
                WordSignatureInfo existingPackageSignatures = InspectPackageSignatures(stagingPath);
                if (existingPackageSignatures.HasSignatures &&
                    options.ExistingPackageSignaturePolicy != WordSignedDocumentSavePolicy.AllowSignatureInvalidation) {
                    findings.Add(Finding("ExistingPackageSignatureInvalidationBlocked", WordSignatureValidationState.Failed,
                        "VBA signing was blocked because changing VBA signature relationships would invalidate existing OPC package signatures. " +
                        "Set ExistingPackageSignaturePolicy to AllowSignatureInvalidation only when the package will be re-signed afterward."));
                    return SigningResult(fullPath, dependencies.Platform.IsWindows, false, false, null, findings);
                }
                if (existingPackageSignatures.HasSignatures) {
                    findings.Add(Finding("ExistingPackageSignaturesWillBeInvalidated", WordSignatureValidationState.NotChecked,
                        "Existing OPC package signatures are permitted to become invalid; sign the VBA project first and the OPC package last."));
                }
                WordMacroProjectToolResult clear = dependencies.Runner.Run(
                    new WordMacroProjectToolInvocation(clearTool, new[] { stagingPath }),
                    options.ToolTimeout,
                    options.MaxToolOutputCharacters);
                if (!clear.Succeeded) {
                    findings.Add(ToolFailure("MacroSignatureClearFailed", "offclearsig", clear));
                    return SigningResult(fullPath, ToolWasAvailable(clear), false, false, null, findings);
                }

                IReadOnlyList<string> signArguments = BuildSignArguments(stagingPath, thumbprint, options);
                for (int profileIndex = 0; profileIndex < 3; profileIndex++) {
                    WordMacroProjectSignatureProfile profile = (WordMacroProjectSignatureProfile)(profileIndex + 1);
                    WordMacroProjectToolResult sign = dependencies.Runner.Run(
                        new WordMacroProjectToolInvocation(signTool, signArguments),
                        options.ToolTimeout,
                        options.MaxToolOutputCharacters);
                    if (!sign.Succeeded) {
                        findings.Add(ToolFailure("Macro" + profile + "SigningFailed",
                            profile + " VBA signing", sign, profile));
                        return SigningResult(fullPath, ToolWasAvailable(sign), false, false, null, findings);
                    }
                    WordMacroProjectSignatureInfo profileReadback =
                        WordMacroProjectSignatureInspector.Inspect(
                            stagingPath, options.Inspection, inspectionBudget, profile, validateCmsOverride: false);
                    WordMacroProjectSignaturePartInfo? createdProfile = profileReadback.Signatures
                        .FirstOrDefault(signature => signature.Profile == profile);
                    if (createdProfile == null) {
                        WordMacroProjectSignatureInfo diagnosticReadback =
                            WordMacroProjectSignatureInspector.Inspect(
                                stagingPath, options.Inspection, inspectionBudget);
                        foreach (WordMacroProjectSignatureFinding finding in diagnosticReadback.Findings
                            .Concat(diagnosticReadback.Signatures.SelectMany(signature => signature.Findings))) {
                            if (!findings.Any(existing =>
                                existing.Code == finding.Code &&
                                existing.Profile == finding.Profile &&
                                existing.Message == finding.Message)) {
                                findings.Add(finding);
                            }
                        }
                        findings.Add(Finding("Macro" + profile + "ReadbackFailed",
                            WordSignatureValidationState.Failed,
                            "The " + profile + " VBA signature was not present after creation.",
                            profile));
                        return SigningResult(fullPath, true, false, false, null, findings);
                    }
                    findings.Add(Finding("Macro" + profile + "SignatureCreated",
                        WordSignatureValidationState.Passed,
                        "The " + profile + " VBA signature was created and discovered before the next profile.", profile));
                }

                validationSnapshotPath = OfficeFileCommit.CreateStagingPath(fullPath);
                WordPackageSnapshot.CopyBounded(
                    stagingPath,
                    validationSnapshotPath,
                    options.Inspection.PackageSecurity.MaxPackageBytes);
                string validatedPackageHash = WordPackageSnapshot.ComputeSha256(
                    validationSnapshotPath,
                    options.Inspection.PackageSecurity.MaxPackageBytes);
                WordMacroProjectSignatureValidationResult stagedValidation = ValidateSnapshot(
                    fullPath, validationSnapshotPath, options, dependencies, inspectionBudget);
                if (!string.Equals(
                    validatedPackageHash,
                    WordPackageSnapshot.ComputeSha256(
                        validationSnapshotPath,
                        options.Inspection.PackageSecurity.MaxPackageBytes),
                    StringComparison.Ordinal)) {
                    findings.Add(Finding("MacroValidatedSnapshotChanged", WordSignatureValidationState.Failed,
                        "The bounded VBA validation snapshot changed during policy validation; the original file was not replaced."));
                    return SigningResult(fullPath, true, false, false, stagedValidation, findings);
                }
                findings.AddRange(stagedValidation.Findings.Where(finding =>
                    !findings.Any(existing => existing.Code == finding.Code && existing.Profile == finding.Profile)));
                WordMacroProjectSignatureInfo stagedInfo = stagedValidation.SignatureInfo;
                bool preserved = string.Equals(sourceInfo.MacroProjectSha256, stagedInfo.MacroProjectSha256,
                    StringComparison.OrdinalIgnoreCase);
                if (!preserved) {
                    findings.Add(Finding("MacroProjectChangedDuringSigning", WordSignatureValidationState.Failed,
                        "The encoded VBA project changed while signatures were staged; the original file was not replaced."));
                    return SigningResult(fullPath, true, false, false, stagedValidation, findings);
                }
                if (!HasAllProfiles(stagedInfo)) {
                    findings.Add(Finding("MacroSignatureProfilesIncomplete", WordSignatureValidationState.Failed,
                        "Signing did not produce the required legacy, agile, and V3 VBA signature parts."));
                    return SigningResult(fullPath, true, false, true, stagedValidation, findings);
                }
                WordMacroProjectSignatureFinding? rejectedProfileFinding = stagedValidation.Findings
                    .FirstOrDefault(finding =>
                        finding.State == WordSignatureValidationState.Failed &&
                        finding.Profile.HasValue);
                if (rejectedProfileFinding != null) {
                    findings.Add(Finding("MacroSignatureProfilePolicyFailed", WordSignatureValidationState.Failed,
                        "The completed " + rejectedProfileFinding.Profile!.Value +
                        " VBA signature profile had failed inspection or policy evidence.",
                        rejectedProfileFinding.Profile));
                    return SigningResult(fullPath, true, false, true, stagedValidation, findings);
                }
                WordMacroProjectSignaturePartInfo? rejectedProfile = stagedInfo.Signatures
                    .OrderBy(signature => signature.Profile)
                    .FirstOrDefault(signature => !ProfileAccepted(signature, options));
                if (rejectedProfile != null) {
                    findings.Add(Finding("MacroSignatureProfilePolicyFailed", WordSignatureValidationState.Failed,
                        "The completed " + rejectedProfile.Profile +
                        " VBA signature did not satisfy CMS, certificate, revocation, and timestamp policy.",
                        rejectedProfile.Profile));
                    return SigningResult(fullPath, true, false, true, stagedValidation, findings);
                }
                if (!stagedValidation.IsValidUnderPolicy) {
                    findings.Add(Finding("MacroSignatureReadbackFailed", WordSignatureValidationState.Failed,
                        "The completed V3 VBA signature did not satisfy content-binding and certificate policy."));
                    return SigningResult(fullPath, true, false, true, stagedValidation, findings);
                }

                if (!OfficeFileCommit.TryCommitTemporaryFileAtomicallyIfDestinationUnchanged(
                    stagingPath,
                    fullPath,
                    displacedPath => string.Equals(
                        sourcePackageHash,
                        WordPackageSnapshot.ComputeSha256(
                            displacedPath,
                            options.Inspection.PackageSecurity.MaxPackageBytes),
                        StringComparison.Ordinal),
                    installedPath => string.Equals(
                        validatedPackageHash,
                        WordPackageSnapshot.ComputeSha256(
                            installedPath,
                            options.Inspection.PackageSecurity.MaxPackageBytes),
                        StringComparison.Ordinal))) {
                    stagingPath = string.Empty;
                    findings.Add(Finding("SourcePackageChangedDuringSigning", WordSignatureValidationState.Failed,
                        "The source document changed while native signing tools were running; the newer source was preserved and the staged signature was discarded."));
                    return SigningResult(fullPath, true, false, true, stagedValidation, findings);
                }
                stagingPath = string.Empty;
                WordMacroProjectSignatureInfo committedInfo = WithFilePath(stagedInfo, fullPath);
                var committedValidation = new WordMacroProjectSignatureValidationResult(
                    committedInfo, stagedValidation.IsSupported, stagedValidation.ContentBindingStatus,
                    stagedValidation.Findings) {
                    IsValidUnderPolicy = stagedValidation.IsValidUnderPolicy
                };
                findings.Add(Finding("MacroProjectSignatureCommitted", WordSignatureValidationState.Passed,
                    "The signed package was atomically committed after preservation and policy validation."));
                return SigningResult(fullPath, true, true, true, committedValidation, findings);
            } catch (Exception exception) when (IsSigningException(exception)) {
                findings.Add(Finding("MacroProjectSigningFailed", WordSignatureValidationState.Failed,
                    "VBA signing failed before atomic commit. " + exception.Message));
                return SigningResult(fullPath, true, false, false, null, findings);
            } finally {
                OfficeFileCommit.DeleteIfExists(stagingPath);
                OfficeFileCommit.DeleteIfExists(validationSnapshotPath);
            }
        }

        private static WordMacroProjectSignatureValidationResult Result(
            WordMacroProjectSignatureInfo info,
            bool supported,
            WordSignatureValidationState contentBindingStatus,
            IReadOnlyList<WordMacroProjectSignatureFinding> findings,
            WordMacroProjectSignatureValidationOptions options) {
            var result = new WordMacroProjectSignatureValidationResult(info, supported,
                contentBindingStatus, findings.ToArray());
            WordMacroProjectSignaturePartInfo? selected = SelectHighestProfile(info);
            bool cmsAccepted = !options.Inspection.ValidateCms ||
                selected != null && ProfileAccepted(selected, options);
            result.IsValidUnderPolicy = supported &&
                contentBindingStatus == WordSignatureValidationState.Passed &&
                info.HasMacroProject && info.HasSignatures &&
                (!options.RequireV3Signature || info.HasV3Signature) &&
                cmsAccepted &&
                !findings.Any(finding =>
                    finding.State == WordSignatureValidationState.Failed &&
                    (!finding.Profile.HasValue || selected == null || finding.Profile == selected.Profile));
            return result;
        }

        private static bool ProfileAccepted(
            WordMacroProjectSignaturePartInfo signature,
            WordMacroProjectSignatureValidationOptions options) =>
            signature.CmsParsed &&
            signature.CryptographicStatus == WordSignatureValidationState.Passed &&
            signature.CertificateChainStatus == WordSignatureValidationState.Passed &&
            RevocationAccepted(signature, options) &&
            TimestampAccepted(signature, options) &&
            !signature.Findings.Any(finding => finding.State == WordSignatureValidationState.Failed);

        private static bool RevocationAccepted(
            WordMacroProjectSignaturePartInfo signature,
            WordMacroProjectSignatureValidationOptions options) {
            X509RevocationMode mode = options.Inspection.CmsVerification.CertificateValidation.RevocationMode;
            return mode == X509RevocationMode.NoCheck
                ? signature.RevocationStatus != WordSignatureValidationState.Failed
                : signature.RevocationStatus == WordSignatureValidationState.Passed;
        }

        private static bool TimestampAccepted(
            WordMacroProjectSignaturePartInfo signature,
            WordMacroProjectSignatureValidationOptions options) {
            if (options.RequireTimestamp) return signature.TimestampStatus == WordSignatureValidationState.Passed;
            return signature.TimestampStatus != WordSignatureValidationState.Failed &&
                signature.TimestampStatus != WordSignatureValidationState.Unsupported;
        }

        private static WordMacroProjectSignaturePartInfo? SelectHighestProfile(WordMacroProjectSignatureInfo info) =>
            info.Signatures.OrderByDescending(signature => signature.Profile).FirstOrDefault();

        private static bool CanValidate(
            WordMacroProjectSignatureInfo info,
            ICollection<WordMacroProjectSignatureFinding> findings) {
            if (!info.IsMacroEnabledFormat || !info.HasMacroProject) return false;
            if (!info.HasSignatures) {
                findings.Add(Finding("MacroSignatureNotPresent", WordSignatureValidationState.NotPresent,
                    "The VBA project does not contain a signature part."));
                return false;
            }
            return true;
        }

        private static List<WordMacroProjectSignatureFinding> CollectInspectionFindings(
            WordMacroProjectSignatureInfo info) {
            var findings = new List<WordMacroProjectSignatureFinding>(info.Findings);
            findings.AddRange(info.Signatures.SelectMany(signature => signature.Findings));
            return findings;
        }

        private static bool HasAllProfiles(WordMacroProjectSignatureInfo info) =>
            info.Signatures.Any(signature => signature.Profile == WordMacroProjectSignatureProfile.Legacy) &&
            info.Signatures.Any(signature => signature.Profile == WordMacroProjectSignatureProfile.Agile) &&
            info.Signatures.Any(signature => signature.Profile == WordMacroProjectSignatureProfile.V3);

        private static IReadOnlyList<string> BuildSignArguments(
            string filePath,
            string thumbprint,
            WordMacroProjectSigningOptions options) {
            var arguments = new List<string> {
                "sign", "/sha1", thumbprint, "/s", GetWindowsStoreName(options.StoreName)
            };
            if (options.StoreLocation == StoreLocation.LocalMachine) arguments.Add("/sm");
            arguments.Add("/fd");
            arguments.Add("SHA256");
            if (options.TimestampAuthorityUrl != null) {
                arguments.Add("/tr");
                arguments.Add(options.TimestampAuthorityUrl.AbsoluteUri);
                arguments.Add("/td");
                arguments.Add("SHA256");
            }
            arguments.Add(filePath);
            return arguments;
        }

        private static string GetWindowsStoreName(StoreName storeName) =>
            storeName == StoreName.CertificateAuthority ? "CA" : storeName.ToString();

        private static void ValidateValidationOptions(WordMacroProjectSignatureValidationOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            WordMacroProjectSignatureInspector.ValidateOptions(options.Inspection);
            if (!options.Inspection.ValidateCms) {
                throw new ArgumentException(
                    "VBA content-binding validation requires CMS inspection so the signed Office SIP digest is authenticated.",
                    nameof(options));
            }
            if (options.RequireTimestamp && !options.Inspection.CmsVerification.ValidateTimestamps) {
                throw new ArgumentException(
                    "RequireTimestamp cannot be combined with disabled timestamp inspection.", nameof(options));
            }
        }

        private static void ValidateSigningOptions(WordMacroProjectSigningOptions options) {
            ValidateValidationOptions(options);
            if (options.ToolTimeout <= TimeSpan.Zero || options.ToolTimeout.TotalMilliseconds > int.MaxValue) {
                throw new ArgumentOutOfRangeException(nameof(options.ToolTimeout));
            }
            if (options.MaxToolOutputCharacters <= 0 || options.MaxToolOutputCharacters > 4 * 1024 * 1024) {
                throw new ArgumentOutOfRangeException(nameof(options.MaxToolOutputCharacters));
            }
            if (!Enum.IsDefined(typeof(WordSignedDocumentSavePolicy), options.ExistingPackageSignaturePolicy)) {
                throw new ArgumentOutOfRangeException(nameof(options.ExistingPackageSignaturePolicy));
            }
            if (options.RequireTimestamp && options.TimestampAuthorityUrl == null) {
                throw new ArgumentException(
                    "TimestampAuthorityUrl is required when newly created VBA signatures must contain a timestamp.",
                    nameof(options));
            }
        }

        private static WordSignatureInfo InspectPackageSignatures(string filePath) {
            using WordprocessingDocument document = WordprocessingDocument.Open(filePath, false);
            DigitalSignatureOriginPart? origin = document.DigitalSignatureOriginPart;
            bool hasApplicationMetadata = document.ExtendedFilePropertiesPart?.Properties?.DigitalSignature != null;
            return WordSignatureInspector.Inspect(document, origin, hasApplicationMetadata,
                packageBytes: null, verifyDigests: false);
        }

        private static bool TryResolveSignTool(string? configuredPath, out string path, out string detail) {
            if (string.IsNullOrWhiteSpace(configuredPath)) {
                path = "signtool.exe";
                detail = "SignTool.exe will be resolved through PATH.";
                return true;
            }
            path = Path.GetFullPath(configuredPath);
            if (File.Exists(path)) {
                detail = "SignTool.exe was resolved from the configured path.";
                return true;
            }
            detail = "The configured SignTool.exe does not exist: " + path;
            return false;
        }

        private static bool TryResolveClearSignatureTool(
            string? officeSipsDirectory,
            out string path,
            out string detail) {
            path = string.Empty;
            if (string.IsNullOrWhiteSpace(officeSipsDirectory)) {
                detail = "OfficeSipsDirectory must identify Microsoft's Office SIP tools containing offclearsig.exe.";
                return false;
            }
            path = Path.Combine(Path.GetFullPath(officeSipsDirectory), "offclearsig.exe");
            if (File.Exists(path)) {
                detail = "Microsoft's offclearsig.exe was resolved from OfficeSipsDirectory.";
                return true;
            }
            detail = "Microsoft's offclearsig.exe does not exist at " + path;
            return false;
        }

        private static bool TryNormalizeThumbprint(string? value, out string thumbprint, out string detail) {
            thumbprint = string.Empty;
            if (string.IsNullOrWhiteSpace(value)) {
                detail = "A certificate SHA-1 thumbprint is required.";
                return false;
            }
            var characters = new List<char>(value!.Length);
            foreach (char character in value) {
                if (Uri.IsHexDigit(character)) characters.Add(char.ToUpperInvariant(character));
                else if (char.IsWhiteSpace(character) || character == ':' || character == '-') continue;
                else {
                    detail = "The certificate thumbprint contains invalid character '" + character + "'.";
                    return false;
                }
            }
            if (characters.Count != 40) {
                detail = "SignTool certificate selection requires a 40-character SHA-1 thumbprint.";
                return false;
            }
            thumbprint = new string(characters.ToArray());
            detail = string.Empty;
            return true;
        }

        private static bool IsHttpTimestampUrl(Uri value) => value.IsAbsoluteUri &&
            (string.Equals(value.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase) ||
             string.Equals(value.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase));

        private static WordMacroProjectSignatureFinding ToolFailure(
            string code,
            string operation,
            WordMacroProjectToolResult result,
            WordMacroProjectSignatureProfile? profile = null) {
            string outcome = result.TimedOut
                ? operation + " exceeded its configured timeout."
                : operation + " failed" + (result.ExitCode.HasValue ? " with exit code " + result.ExitCode.Value : " to start") + ".";
            string output = SummarizeOutput(result.Output);
            if (!string.IsNullOrWhiteSpace(output)) outcome += " " + output;
            return Finding(code, result.TimedOut
                ? WordSignatureValidationState.Unsupported
                : WordSignatureValidationState.Failed, outcome, profile);
        }

        private static string SummarizeOutput(string value) {
            if (string.IsNullOrWhiteSpace(value)) return string.Empty;
            string flattened = string.Join(" ", value.Split(new[] { '\r', '\n' },
                StringSplitOptions.RemoveEmptyEntries).Select(line => line.Trim()));
            return flattened.Length <= 2048 ? flattened : flattened.Substring(0, 2048) + "…";
        }

        private static bool ToolWasAvailable(WordMacroProjectToolResult result) =>
            result.ExitCode.HasValue || result.TimedOut;

        private static WordMacroProjectSigningResult SigningResult(
            string filePath,
            bool supported,
            bool succeeded,
            bool preserved,
            WordMacroProjectSignatureValidationResult? validation,
            IReadOnlyList<WordMacroProjectSignatureFinding> findings) =>
            new WordMacroProjectSigningResult(filePath, supported, succeeded, preserved, validation, findings.ToArray());

        private static WordMacroProjectSignatureFinding Finding(
            string code,
            WordSignatureValidationState state,
            string message,
            WordMacroProjectSignatureProfile? profile = null) =>
            new WordMacroProjectSignatureFinding(code, state, message, profile);

        private static string NormalizePath(string filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A Word package path is required.", nameof(filePath));
            return Path.GetFullPath(filePath);
        }

        private static bool IsSigningException(Exception exception) =>
            exception is IOException || exception is UnauthorizedAccessException ||
            exception is InvalidDataException || exception is NotSupportedException ||
            exception is ArgumentException || exception is OverflowException;
    }
}
