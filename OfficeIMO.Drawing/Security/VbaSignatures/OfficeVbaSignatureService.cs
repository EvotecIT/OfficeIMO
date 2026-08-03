using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Security;

/// <summary>Shared bounded VBA signature inspection, validation, and Windows Office SIP signing.</summary>
public static partial class OfficeVbaSignatureService {
    private const string VbaProjectContentType = "application/vnd.ms-office.vbaProject";
    private const string LegacyRelationship = "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature";
    private const string AgileRelationship = "http://schemas.microsoft.com/office/2014/relationships/vbaProjectSignatureAgile";
    private const string AgileCompatibilityRelationship = "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignatureAgile";
    private const string V3Relationship = "http://schemas.microsoft.com/office/2020/07/relationships/vbaProjectSignatureV3";

    /// <summary>Inspects VBA signature profile carriers without performing cryptographic validation.</summary>
    public static OfficeVbaSignatureInfo Inspect(
        string filePath,
        OfficeVbaSignatureInspectionOptions? options = null) =>
        InspectCore(filePath, null, options ?? new OfficeVbaSignatureInspectionOptions());

    /// <summary>Inspects and cryptographically validates VBA signature profile carriers.</summary>
    public static OfficeVbaSignatureInfo Inspect(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        OfficeVbaSignatureInspectionOptions? options = null) {
        if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
        return InspectCore(filePath, securityProvider, options ?? new OfficeVbaSignatureInspectionOptions());
    }

    /// <summary>Validates the highest-precedence VBA signature against CMS policy and the registered Office SIP.</summary>
    public static OfficeVbaSignatureValidationResult Validate(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        OfficeVbaSignatureInspectionOptions? options = null) {
        if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
        OfficeVbaSignatureInfo info = Inspect(filePath, securityProvider, options);
        var findings = new List<OfficeVbaSignatureFinding>(info.Findings);
        findings.AddRange(info.Signatures.SelectMany(signature => signature.Findings));
        OfficeVbaSignaturePartInfo? selected = info.Signatures
            .OrderByDescending(signature => signature.Profile)
            .FirstOrDefault();
        if (selected == null) {
            findings.Add(Finding("VbaSignatureNotPresent", OfficePackageSignatureValidationState.NotPresent,
                "The VBA project does not contain a signature profile."));
            return new OfficeVbaSignatureValidationResult(info, false,
                OfficePackageSignatureValidationState.NotPresent, findings);
        }
        if (string.IsNullOrWhiteSpace(selected.SubjectDigestAlgorithmOid) || selected.SubjectDigest == null) {
            findings.Add(Finding("VbaSubjectDigestMissing", OfficePackageSignatureValidationState.Failed,
                "The highest VBA signature profile does not contain a bounded Authenticode Office SIP digest.", selected.Profile));
            return new OfficeVbaSignatureValidationResult(info, false,
                OfficePackageSignatureValidationState.Failed, findings);
        }

        OfficeVbaContentBindingResult binding = OfficeVbaWindowsSip.ValidateContentBinding(
            info.FilePath, selected.SubjectDigestAlgorithmOid!, selected.SubjectDigest);
        findings.Add(Finding(binding.IsValid ? "VbaContentBindingValid" : "VbaContentBindingInvalid",
            binding.IsSupported
                ? binding.IsValid ? OfficePackageSignatureValidationState.Passed : OfficePackageSignatureValidationState.Failed
                : OfficePackageSignatureValidationState.Unsupported,
            binding.Detail, selected.Profile));
        return new OfficeVbaSignatureValidationResult(info, binding.IsSupported,
            binding.IsSupported
                ? binding.IsValid ? OfficePackageSignatureValidationState.Passed : OfficePackageSignatureValidationState.Failed
                : OfficePackageSignatureValidationState.Unsupported,
            findings);
    }

    private static OfficeVbaSignatureInfo InspectCore(
        string filePath,
        IOfficeSecurityProvider? securityProvider,
        OfficeVbaSignatureInspectionOptions options) {
        ValidateOptions(options);
        string fullPath = NormalizePath(filePath);
        bool macroEnabled = IsMacroEnabledPath(fullPath);
        var findings = new List<OfficeVbaSignatureFinding>();
        if (!macroEnabled) findings.Add(Finding("MacroEnabledFormatRequired",
            OfficePackageSignatureValidationState.Failed,
            "VBA signature operations require DOCM, DOTM, XLSM, XLTM, XLAM, XLSB, PPTM, POTM, PPSM, or PPAM."));
        var file = new FileInfo(fullPath);
        if (!file.Exists) {
            findings.Add(Finding("FileNotFound", OfficePackageSignatureValidationState.Failed,
                "The Office package does not exist: " + fullPath));
            return Empty(fullPath, macroEnabled, findings);
        }
        if (file.Length > options.Package.MaxPackageBytes) {
            findings.Add(Finding("PackageByteLimitExceeded", OfficePackageSignatureValidationState.Failed,
                "The Office package exceeds the configured byte limit."));
            return Empty(fullPath, macroEnabled, findings);
        }

        try {
            byte[] packageBytes = File.ReadAllBytes(fullPath);
            using var archive = new OfficePackageSignatureArchive(
                packageBytes, options.Package.MaxPackageParts, options.Package.MaxPartBytes);
            string[] vbaParts = archive.PartUris.Where(uri =>
                    archive.TryGetContentType(uri, out string contentType) &&
                    string.Equals(contentType, VbaProjectContentType, StringComparison.OrdinalIgnoreCase) ||
                    uri.EndsWith("/vbaProject.bin", StringComparison.OrdinalIgnoreCase))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            if (vbaParts.Length == 0) {
                findings.Add(Finding("MacroProjectNotPresent", OfficePackageSignatureValidationState.NotPresent,
                    "The Office package does not contain vbaProject.bin."));
                return Empty(fullPath, macroEnabled, findings);
            }
            if (vbaParts.Length > 1) {
                findings.Add(Finding("MultipleMacroProjects", OfficePackageSignatureValidationState.Failed,
                    "The Office package declares more than one VBA project part."));
                return Empty(fullPath, macroEnabled, findings);
            }

            string vbaUri = vbaParts[0];
            byte[] vbaBytes = archive.ReadPart(vbaUri, options.MaxMacroProjectBytes);
            string hash;
            using (SHA256 algorithm = SHA256.Create()) {
                hash = BitConverter.ToString(algorithm.ComputeHash(vbaBytes)).Replace("-", string.Empty);
            }
            IReadOnlyList<OfficeVbaSignaturePartInfo> signatures = InspectRelationships(
                archive, vbaUri, securityProvider, options, findings);
            return new OfficeVbaSignatureInfo(fullPath, macroEnabled, true, vbaUri,
                vbaBytes.LongLength, hash, signatures, findings);
        } catch (Exception exception) when (exception is IOException or InvalidDataException or XmlException or CryptographicException) {
            findings.Add(Finding("VbaSignatureInspectionFailed", OfficePackageSignatureValidationState.Failed,
                "The VBA signature package structure could not be inspected. " + exception.Message));
            return Empty(fullPath, macroEnabled, findings);
        }
    }

    private static IReadOnlyList<OfficeVbaSignaturePartInfo> InspectRelationships(
        OfficePackageSignatureArchive archive,
        string vbaUri,
        IOfficeSecurityProvider? provider,
        OfficeVbaSignatureInspectionOptions options,
        ICollection<OfficeVbaSignatureFinding> findings) {
        string relationshipsUri = GetRelationshipPartUri(vbaUri);
        if (!archive.ContainsPart(relationshipsUri)) return Array.Empty<OfficeVbaSignaturePartInfo>();
        XDocument relationships = LoadXml(archive.ReadPart(relationshipsUri, options.MaxSignatureBytes));
        XElement[] declarations = relationships.Descendants()
            .Where(element => element.Name.LocalName == "Relationship")
            .Take(options.MaxRelationships + 1)
            .ToArray();
        if (declarations.Length > options.MaxRelationships) {
            throw new InvalidDataException("The VBA project has more relationships than the configured limit.");
        }

        var signatures = new List<OfficeVbaSignaturePartInfo>();
        var profiles = new HashSet<OfficeVbaSignatureProfile>();
        long totalBytes = 0;
        ICmsVerificationSession? session = provider != null && options.ValidateCms
            ? provider.CreateCmsVerificationSession(options.CmsVerification)
            : null;
        try {
            foreach (XElement declaration in declarations) {
                string? relationshipType = (string?)declaration.Attribute("Type");
                if (!TryGetProfile(relationshipType, out OfficeVbaSignatureProfile profile)) continue;
                if (string.Equals((string?)declaration.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) {
                    findings.Add(Finding("VbaSignatureExternalTarget", OfficePackageSignatureValidationState.Failed,
                        "A VBA signature relationship targets an external resource.", profile));
                    continue;
                }
                if (!profiles.Add(profile)) {
                    findings.Add(Finding("DuplicateVbaSignatureProfile", OfficePackageSignatureValidationState.Failed,
                        "More than one " + profile + " VBA signature relationship is present.", profile));
                    continue;
                }
                string? target = (string?)declaration.Attribute("Target");
                string? partUri = ResolvePartUri(vbaUri, target);
                if (partUri == null || !archive.ContainsPart(partUri)) {
                    findings.Add(Finding("VbaSignaturePartMissing", OfficePackageSignatureValidationState.Failed,
                        "The " + profile + " VBA signature target is missing.", profile));
                    continue;
                }
                byte[] encoded = archive.ReadPart(partUri, options.MaxSignatureBytes);
                totalBytes = checked(totalBytes + encoded.LongLength);
                if (totalBytes > options.MaxTotalSignatureBytes) {
                    throw new InvalidDataException("The aggregate VBA signature bytes exceed the configured limit.");
                }
                archive.TryGetContentType(partUri, out string contentType);
                signatures.Add(InspectPart(profile, partUri, relationshipType!, contentType ?? string.Empty,
                    encoded, session, options));
            }
        } finally {
            session?.Dispose();
        }
        return signatures.OrderBy(signature => signature.Profile).ToArray();
    }

    private static OfficeVbaSignaturePartInfo InspectPart(
        OfficeVbaSignatureProfile profile,
        string uri,
        string relationshipType,
        string contentType,
        byte[] encoded,
        ICmsVerificationSession? session,
        OfficeVbaSignatureInspectionOptions options) {
        var findings = new List<OfficeVbaSignatureFinding>();
        string expectedContentType = GetContentType(profile);
        if (!string.Equals(contentType, expectedContentType, StringComparison.OrdinalIgnoreCase)) {
            findings.Add(Finding("VbaSignatureContentTypeUnexpected", OfficePackageSignatureValidationState.Failed,
                "The " + profile + " signature content type is '" + contentType + "' instead of '" + expectedContentType + "'.", profile));
        }
        if (!TryExtractCms(encoded, options.CmsVerification.MaxEncodedBytes, out byte[] cmsBytes, out string detail)) {
            findings.Add(Finding("VbaSignatureContainerMalformed", OfficePackageSignatureValidationState.Failed, detail, profile));
            return Part(profile, uri, relationshipType, contentType, encoded.LongLength, false,
                OfficePackageSignatureValidationState.Failed, OfficePackageSignatureValidationState.NotChecked,
                OfficePackageSignatureValidationState.NotChecked, OfficePackageSignatureValidationState.NotChecked,
                null, null, null, null, findings);
        }
        if (session == null) {
            return Part(profile, uri, relationshipType, contentType, encoded.LongLength, true,
                OfficePackageSignatureValidationState.NotChecked, OfficePackageSignatureValidationState.NotChecked,
                OfficePackageSignatureValidationState.NotChecked, OfficePackageSignatureValidationState.NotChecked,
                null, null, null, null, findings);
        }

        CmsVerificationResult cms = session.Verify(cmsBytes, CertificateValidationPurpose.DocumentSigning);
        findings.AddRange(cms.Findings.Select(finding => Finding(finding.Code,
            finding.Severity == SecurityFindingSeverity.Error
                ? OfficePackageSignatureValidationState.Failed
                : OfficePackageSignatureValidationState.NotChecked,
            finding.Message, profile)));
        CmsSignerVerificationResult? signer = cms.Signers.Count == 1 ? cms.Signers[0] : null;
        if (signer == null) findings.Add(Finding("VbaSignatureSignerCountInvalid",
            OfficePackageSignatureValidationState.Failed,
            "A VBA signature profile must contain exactly one CMS signer.", profile));
        return Part(profile, uri, relationshipType, contentType, encoded.LongLength, cms.Parsed,
            signer != null && cms.IsCryptographicallyValid
                ? OfficePackageSignatureValidationState.Passed : OfficePackageSignatureValidationState.Failed,
            signer == null ? OfficePackageSignatureValidationState.NotPresent : Map(signer.CertificateValidation.ChainStatus),
            signer == null ? OfficePackageSignatureValidationState.NotPresent : Map(signer.CertificateValidation.RevocationStatus),
            signer == null ? OfficePackageSignatureValidationState.NotPresent : Map(signer.TimestampStatus),
            signer?.Subject, signer?.Thumbprint,
            cms.AuthenticodeIndirectData?.DigestAlgorithmOid,
            cms.AuthenticodeIndirectData?.Digest,
            findings.Concat(cms.Signers.SelectMany(item => item.Findings).Select(finding => Finding(finding.Code,
                finding.Severity == SecurityFindingSeverity.Error
                    ? OfficePackageSignatureValidationState.Failed
                    : OfficePackageSignatureValidationState.NotChecked,
                finding.Message, profile))).ToArray());
    }

    private static OfficeVbaSignaturePartInfo Part(OfficeVbaSignatureProfile profile, string uri,
        string relationshipType, string contentType, long length, bool cmsParsed,
        OfficePackageSignatureValidationState crypto, OfficePackageSignatureValidationState chain,
        OfficePackageSignatureValidationState revocation, OfficePackageSignatureValidationState timestamp,
        string? subject, string? thumbprint, string? digestAlgorithm, byte[]? digest,
        IReadOnlyList<OfficeVbaSignatureFinding> findings) =>
        new(profile, uri, relationshipType, contentType, length, cmsParsed, crypto, chain,
            revocation, timestamp, subject, thumbprint, digestAlgorithm, digest, findings);

    private static bool TryExtractCms(byte[] encoded, long maxCmsBytes, out byte[] cms, out string detail) {
        const int headerLength = 36;
        cms = Array.Empty<byte>();
        detail = string.Empty;
        if (encoded.Length < headerLength) {
            detail = "The VBA signature part is shorter than DigSigInfoSerialized.";
            return false;
        }
        uint signatureLength = ReadUInt32(encoded, 0);
        uint serializedOffset = ReadUInt32(encoded, 4);
        if (serializedOffset < 44) {
            detail = "The VBA signature CMS offset is outside the bounded signature part.";
            return false;
        }
        uint offset = serializedOffset - 8;
        if (signatureLength == 0 || signatureLength > maxCmsBytes || offset > encoded.Length ||
            signatureLength > encoded.Length - offset) {
            detail = "The VBA signature CMS offset or length is outside the bounded signature part.";
            return false;
        }
        cms = new byte[signatureLength];
        Buffer.BlockCopy(encoded, checked((int)offset), cms, 0, checked((int)signatureLength));
        return true;
    }

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        (uint)(bytes[offset] | bytes[offset + 1] << 8 | bytes[offset + 2] << 16 | bytes[offset + 3] << 24);

    private static string GetRelationshipPartUri(string partUri) {
        string normalized = OfficePackageSignatureArchive.NormalizePartUri(partUri);
        int slash = normalized.LastIndexOf('/');
        string directory = slash <= 0 ? string.Empty : normalized.Substring(0, slash);
        string name = normalized.Substring(slash + 1);
        return directory + "/_rels/" + name + ".rels";
    }

    private static string? ResolvePartUri(string sourcePartUri, string? target) {
        if (string.IsNullOrWhiteSpace(target)) return null;
        if (target!.StartsWith("/", StringComparison.Ordinal)) return OfficePackageSignatureArchive.NormalizePartUri(target);
        string source = OfficePackageSignatureArchive.NormalizePartUri(sourcePartUri);
        int slash = source.LastIndexOf('/');
        string combined = (slash <= 0 ? "/" : source.Substring(0, slash + 1)) + target;
        var segments = new List<string>();
        foreach (string segment in combined.Split('/')) {
            if (segment.Length == 0 || segment == ".") continue;
            if (segment == "..") {
                if (segments.Count == 0) return null;
                segments.RemoveAt(segments.Count - 1);
            } else segments.Add(segment);
        }
        return "/" + string.Join("/", segments);
    }

    private static bool TryGetProfile(string? relationshipType, out OfficeVbaSignatureProfile profile) {
        if (string.Equals(relationshipType, LegacyRelationship, StringComparison.Ordinal)) {
            profile = OfficeVbaSignatureProfile.Legacy;
            return true;
        }
        if (string.Equals(relationshipType, AgileRelationship, StringComparison.Ordinal) ||
            string.Equals(relationshipType, AgileCompatibilityRelationship, StringComparison.Ordinal)) {
            profile = OfficeVbaSignatureProfile.Agile;
            return true;
        }
        if (string.Equals(relationshipType, V3Relationship, StringComparison.Ordinal)) {
            profile = OfficeVbaSignatureProfile.V3;
            return true;
        }
        profile = default;
        return false;
    }

    private static string GetContentType(OfficeVbaSignatureProfile profile) => profile switch {
        OfficeVbaSignatureProfile.Legacy => "application/vnd.ms-office.vbaProjectSignature",
        OfficeVbaSignatureProfile.Agile => "application/vnd.ms-office.vbaProjectSignatureAgile",
        OfficeVbaSignatureProfile.V3 => "application/vnd.ms-office.vbaProjectSignatureV3",
        _ => throw new ArgumentOutOfRangeException(nameof(profile))
    };

    private static bool IsMacroEnabledPath(string path) {
        string extension = Path.GetExtension(path);
        return new[] { ".docm", ".dotm", ".xlsm", ".xltm", ".xlam", ".xlsb", ".pptm", ".potm", ".ppsm", ".ppam" }
            .Contains(extension, StringComparer.OrdinalIgnoreCase);
    }

    private static void ValidateOptions(OfficeVbaSignatureInspectionOptions options) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.Package.Validate();
        if (options.MaxMacroProjectBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxMacroProjectBytes));
        if (options.MaxSignatureBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxSignatureBytes));
        if (options.MaxTotalSignatureBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxTotalSignatureBytes));
        if (options.MaxRelationships <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxRelationships));
    }

    private static XDocument LoadXml(byte[] bytes) {
        using var stream = new MemoryStream(bytes, writable: false);
        using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = Math.Max(1, bytes.LongLength)
        });
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    private static string NormalizePath(string filePath) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("An Office package path is required.", nameof(filePath));
        return Path.GetFullPath(filePath);
    }

    private static OfficeVbaSignatureInfo Empty(string filePath, bool macroEnabled,
        IReadOnlyList<OfficeVbaSignatureFinding> findings) =>
        new(filePath, macroEnabled, false, null, null, null,
            Array.Empty<OfficeVbaSignaturePartInfo>(), findings);

    private static OfficeVbaSignatureFinding Finding(string code,
        OfficePackageSignatureValidationState state, string message,
        OfficeVbaSignatureProfile? profile = null) => new(code, state, message, profile);

    private static OfficePackageSignatureValidationState Map(SecurityValidationStatus status) => status switch {
        SecurityValidationStatus.Valid => OfficePackageSignatureValidationState.Passed,
        SecurityValidationStatus.Invalid => OfficePackageSignatureValidationState.Failed,
        SecurityValidationStatus.NotPerformed => OfficePackageSignatureValidationState.NotChecked,
        _ => OfficePackageSignatureValidationState.Unsupported
    };
}
