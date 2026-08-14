using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Security;

/// <summary>Dependency-light structural inspection and optional-provider validation for OPC package signatures.</summary>
public static class OfficePackageSignatureService {
    private const string OriginRelationshipType =
        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
    private const string SignatureRelationshipType =
        "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature";
    private const string SignatureContentType =
        "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";
    private const string ExtendedPropertiesRelationshipType =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
    private const string ExtendedPropertiesContentType =
        "application/vnd.openxmlformats-officedocument.extended-properties+xml";
    private static readonly XNamespace DigitalSignatureNamespace = XmlDigitalSignatureAlgorithms.Namespace;

    /// <summary>Creates an OPC XML signature in a saved package and returns structured failure evidence.</summary>
    public static OfficePackageSigningResult TrySign(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficePackageSigningOptions? options = null) {
        if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
        options ??= new OfficePackageSigningOptions();
        options.ValidateBeforeCommit ??= (stagingPath, _, _) => {
            var validationOptions = new OfficePackageSignatureValidationOptions {
                ValidateCertificateTrust = false
            };
            validationOptions.Inspection.MaxPackageBytes = options.MaxPackageBytes;
            validationOptions.Inspection.MaxPackageParts = options.MaxPackageParts;
            validationOptions.Inspection.MaxPartBytes = options.MaxPartBytes;
            validationOptions.Inspection.MaxTotalDigestBytes = options.MaxTotalDigestBytes;
            validationOptions.Inspection.MaxSignatureBytes = options.MaxSignatureBytes;
            validationOptions.Inspection.MaxSignedReferences = options.MaxSignedReferences;
            validationOptions.Inspection.MaxCertificates = options.MaxCertificates;
            validationOptions.Inspection.MaxCertificateBytes = options.MaxCertificateBytes;
            validationOptions.Inspection.MaxTotalCertificateBytes = options.MaxTotalCertificateBytes;
            OfficePackageSignatureValidationReport report = Validate(stagingPath, securityProvider, validationOptions);
            return report.IsCryptographicallyValid
                ? null
                : "The created OPC package signature failed bounded cryptographic validation readback.";
        };
        return OfficePackageSignatureWriter.Sign(filePath, signingCertificate, securityProvider, options);
    }

    /// <summary>Creates an OPC XML signature and throws if creation or validation readback fails.</summary>
    public static OfficePackageSigningResult Sign(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficePackageSigningOptions? options = null) {
        OfficePackageSigningResult result = TrySign(filePath, securityProvider, signingCertificate, options);
        if (!result.Succeeded) {
            throw new InvalidOperationException(string.Join(" ", result.Details));
        }
        return result;
    }

    /// <summary>Inspects a saved OPC package without loading a host-specific document model.</summary>
    public static OfficePackageSignatureInfo Inspect(
        string filePath,
        OfficePackageSignatureInspectionOptions? options = null) {
        byte[] packageBytes = ReadPackage(filePath, options ?? new OfficePackageSignatureInspectionOptions());
        return Inspect(packageBytes, options);
    }

    /// <summary>Inspects encoded OPC package bytes.</summary>
    public static OfficePackageSignatureInfo Inspect(
        byte[] packageBytes,
        OfficePackageSignatureInspectionOptions? options = null) {
        if (packageBytes == null) throw new ArgumentNullException(nameof(packageBytes));
        options ??= new OfficePackageSignatureInspectionOptions();
        options.Validate();
        if (packageBytes.LongLength > options.MaxPackageBytes) {
            throw new InvalidDataException("The OPC package exceeds the configured byte limit.");
        }

        using var archive = new OfficePackageSignatureArchive(
            packageBytes, options.MaxPackageParts, options.MaxPartBytes);
        return Inspect(archive, options);
    }

    /// <summary>Validates OPC reference digests, XML signature math, signer trust, and revocation through a caller provider.</summary>
    public static OfficePackageSignatureValidationReport Validate(
        string filePath,
        IOfficeSecurityProvider securityProvider,
        OfficePackageSignatureValidationOptions? options = null) {
        if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
        options ??= new OfficePackageSignatureValidationOptions();
        options.Inspection.VerifyDigests = true;
        byte[] packageBytes = ReadPackage(filePath, options.Inspection);
        return Validate(packageBytes, securityProvider, options);
    }

    /// <summary>Validates encoded OPC package bytes through a caller provider.</summary>
    public static OfficePackageSignatureValidationReport Validate(
        byte[] packageBytes,
        IOfficeSecurityProvider securityProvider,
        OfficePackageSignatureValidationOptions? options = null) {
        if (packageBytes == null) throw new ArgumentNullException(nameof(packageBytes));
        if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
        options ??= new OfficePackageSignatureValidationOptions();
        options.Inspection.VerifyDigests = true;
        options.Inspection.Validate();
        if (packageBytes.LongLength > options.Inspection.MaxPackageBytes) {
            throw new InvalidDataException("The OPC package exceeds the configured byte limit.");
        }

        using var archive = new OfficePackageSignatureArchive(
            packageBytes,
            options.Inspection.MaxPackageParts,
            options.Inspection.MaxPartBytes,
            securityProvider);
        OfficePackageSignatureInfo info = Inspect(archive, options.Inspection);
        var results = new List<OfficePackageSignaturePartValidationResult>(info.SignatureParts.Count);
        var findings = new List<string>(info.Findings);

        foreach (OfficePackageSignaturePartInfo part in info.SignatureParts) {
            OfficePackageSignaturePartValidationResult result = ValidatePart(
                archive, part, securityProvider, options);
            results.Add(result);
            findings.AddRange(result.Findings.Select(finding => finding.Message));
        }

        if (!info.HasSignatures) findings.Add("No OPC package signatures were found.");
        return new OfficePackageSignatureValidationReport(
            info,
            results,
            findings.Distinct(StringComparer.Ordinal).ToArray());
    }

    private static OfficePackageSignatureInfo Inspect(
        OfficePackageSignatureArchive archive,
        OfficePackageSignatureInspectionOptions options) {
        var findings = new List<string>();
        OriginDiscovery origins = ReadOrigins(archive, options, findings);
        string? originUri = origins.FirstOriginUri;
        bool hasOrigin = origins.ExistingPartCount > 0;
        bool hasApplicationMetadata = ReadApplicationSignatureMetadata(archive, options, findings);
        var contentTypeSignatureUris = new HashSet<string>(archive.PartUris
            .Where(uri => archive.TryGetContentType(uri, out string contentType) &&
                string.Equals(contentType, SignatureContentType, StringComparison.OrdinalIgnoreCase)),
            StringComparer.OrdinalIgnoreCase);
        HashSet<string> relationshipSignatureUris = ReadSignatureRelationships(
            archive, origins, options, findings);
        string[] signatureUris = contentTypeSignatureUris
            .Concat(relationshipSignatureUris)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(uri => uri, StringComparer.OrdinalIgnoreCase)
            .ToArray();

        bool signatureDiscoveryComplete = signatureUris.Length <= options.MaxSignatureParts;
        if (!signatureDiscoveryComplete) {
            findings.Add("The package contains more XML signature parts than the configured limit.");
            signatureUris = signatureUris.Take(options.MaxSignatureParts).ToArray();
        }
        if (originUri != null && !hasOrigin) findings.Add("The signature-origin relationship targets a missing part.");
        if (hasOrigin && signatureUris.Length == 0) findings.Add("A signature-origin part exists without XML signature parts.");
        foreach (string orphan in contentTypeSignatureUris.Where(uri => !relationshipSignatureUris.Contains(uri))) {
            findings.Add("XML signature part " + orphan + " is not reachable from the unique package signature origin.");
        }
        foreach (string related in relationshipSignatureUris.Where(uri => !contentTypeSignatureUris.Contains(uri))) {
            findings.Add("Signature relationship target " + related + " does not declare the OPC XML-signature content type.");
        }

        var parts = new List<OfficePackageSignaturePartInfo>(signatureUris.Length);
        long totalDigestBytes = 0;
        foreach (string signatureUri in signatureUris) {
            bool reachable = origins.RelationshipCount == 1 && origins.ExistingPartCount == 1 &&
                relationshipSignatureUris.Contains(signatureUri) && contentTypeSignatureUris.Contains(signatureUri);
            parts.Add(InspectPart(archive, signatureUri, reachable, options, ref totalDigestBytes));
        }
        return new OfficePackageSignatureInfo(origins.RelationshipCount, origins.ExistingPartCount,
            originUri, hasApplicationMetadata, signatureDiscoveryComplete, parts, findings);
    }

    private static OfficePackageSignaturePartInfo InspectPart(
        OfficePackageSignatureArchive archive,
        string signatureUri,
        bool isReachableFromOrigin,
        OfficePackageSignatureInspectionOptions options,
        ref long totalDigestBytes) {
        long length = 0;
        try {
            if (archive.TryGetPartLength(signatureUri, out length) &&
                length > options.MaxTotalDigestBytes - totalDigestBytes) {
                throw new InvalidDataException("OPC package signature inspection exceeds the configured aggregate limit.");
            }
            byte[] bytes = archive.ReadPart(signatureUri, options.MaxSignatureBytes);
            ReserveInspectionBytes(ref totalDigestBytes, bytes.LongLength, options.MaxTotalDigestBytes);
            XDocument document = LoadXml(bytes);
            XElement? signature = document.Root;
            if (signature == null || signature.Name != DigitalSignatureNamespace + "Signature") {
                throw new InvalidDataException("The XML signature part does not have a ds:Signature root element.");
            }

            string? signatureMethod = (string?)signature
                .Descendants(DigitalSignatureNamespace + "SignatureMethod")
                .FirstOrDefault()?.Attribute("Algorithm");
            OfficeXmlSignatureBinding.AuthenticatedContent authenticated = OfficeXmlSignatureBinding.Resolve(
                signature, DigitalSignatureNamespace + "Manifest", options.MaxSignedReferences,
                requirePayload: false);
            XElement[] manifestReferences = authenticated.SignedInfoReferences
                .Where(reference => OfficePackageSignatureArchive.NormalizeReferencePartUri(
                    (string?)reference.Attribute("URI")) != null)
                .Concat(authenticated.Payloads.SelectMany(manifest =>
                    manifest.Elements(DigitalSignatureNamespace + "Reference")))
                .Distinct()
                .ToArray();

            var references = new List<OfficePackageSignatureReferenceInfo>(manifestReferences.Length);
            foreach (XElement reference in manifestReferences) {
                string? uri = ((string?)reference.Attribute("URI"))?.Trim();
                string? target = OfficePackageSignatureArchive.NormalizeReferencePartUri(uri);
                bool? exists = target == null ? null : archive.ContainsPart(target);
                string? digestMethod = (string?)reference
                    .Element(DigitalSignatureNamespace + "DigestMethod")?.Attribute("Algorithm");
                string? digestValue = reference.Element(DigitalSignatureNamespace + "DigestValue")?.Value.Trim();
                string[] transforms = reference
                    .Element(DigitalSignatureNamespace + "Transforms")?
                    .Elements(DigitalSignatureNamespace + "Transform")
                    .Select(transform => ((string?)transform.Attribute("Algorithm"))?.Trim())
                    .Where(algorithm => !string.IsNullOrWhiteSpace(algorithm))
                    .Select(algorithm => algorithm!)
                    .ToArray() ?? Array.Empty<string>();
                OfficePackageDigestResult digest = options.VerifyDigests
                    ? archive.VerifyReference(
                        reference,
                        options.MaxPartBytes,
                        options.MaxTotalDigestBytes - totalDigestBytes)
                    : OfficePackageDigestResult.NotChecked("Digest verification was not requested.");
                ReserveInspectionBytes(ref totalDigestBytes, digest.DigestWorkBytes, options.MaxTotalDigestBytes);
                references.Add(new OfficePackageSignatureReferenceInfo(
                    uri, digestMethod, digestValue, target, exists, transforms, digest.Status, digest.Detail));
            }

            OfficePackageSignatureTimestampInfo[] timestamps = ReadTimestamps(signature);
            var subjects = signature.Descendants(DigitalSignatureNamespace + "X509SubjectName")
                .Select(element => element.Value.Trim())
                .Where(value => value.Length > 0)
                .Distinct(StringComparer.Ordinal)
                .ToList();
            var certificateBudget = new OfficePackageCertificateByteBudget(options.MaxTotalCertificateBytes);
            var certificates = new List<byte[]>(ReadCertificates(signature, options, certificateBudget));
            certificates.AddRange(ReadRelatedCertificates(
                archive, signatureUri, options, certificateBudget));
            if (certificates.Count > options.MaxCertificates) {
                throw new InvalidDataException("The XML signature contains more certificates than the configured limit.");
            }
            foreach (byte[] encoded in certificates) {
                try {
                    using X509Certificate2 certificate = OfficePackageCertificateLoader.Load(encoded);
                    if (!string.IsNullOrWhiteSpace(certificate.Subject)) subjects.Add(certificate.Subject);
                } catch (System.Security.Cryptography.CryptographicException exception) {
                    throw new InvalidDataException("A signature certificate could not be parsed.", exception);
                }
            }
            return new OfficePackageSignaturePartInfo(
                signatureUri, length, isReachableFromOrigin, signatureMethod, references, timestamps,
                subjects.Distinct(StringComparer.OrdinalIgnoreCase).ToArray(), certificates.ToArray(), null);
        } catch (Exception exception) when (exception is IOException or InvalidDataException or XmlException or FormatException or OverflowException) {
            return new OfficePackageSignaturePartInfo(
                signatureUri, length, isReachableFromOrigin, null, Array.Empty<OfficePackageSignatureReferenceInfo>(),
                Array.Empty<OfficePackageSignatureTimestampInfo>(), Array.Empty<string>(), Array.Empty<byte[]>(), exception.Message);
        }
    }

    private static void ReserveInspectionBytes(ref long totalBytes, long bytes, long maximumBytes) {
        if (bytes < 0 || totalBytes > maximumBytes - bytes) {
            throw new InvalidDataException("OPC package signature inspection exceeds the configured aggregate limit.");
        }
        totalBytes += bytes;
    }

    private static OfficePackageSignaturePartValidationResult ValidatePart(
        OfficePackageSignatureArchive archive,
        OfficePackageSignaturePartInfo part,
        IOfficeSecurityProvider provider,
        OfficePackageSignatureValidationOptions options) {
        var findings = new List<SecurityFinding>();
        if (part.HasParseError) {
            findings.Add(new SecurityFinding(SecurityFindingSeverity.Error,
                "OpcSignatureMalformed", part.ParseError ?? "The XML signature part is malformed."));
            return Result(part, OfficePackageSignatureValidationState.Failed,
                OfficePackageSignatureValidationState.NotChecked,
                OfficePackageSignatureValidationState.NotChecked, options.ValidateCertificateTrust,
                IsRevocationRequired(options), findings);
        }

        var certificates = new List<X509Certificate2>();
        try {
            foreach (byte[] encoded in part.CertificateBytes) certificates.Add(OfficePackageCertificateLoader.Load(encoded));
            if (certificates.Count == 0) {
                findings.Add(new SecurityFinding(SecurityFindingSeverity.Error,
                    "OpcSignerCertificateMissing", "The signature does not embed a signer certificate."));
                return Result(part, OfficePackageSignatureValidationState.Failed,
                    OfficePackageSignatureValidationState.NotPresent,
                    OfficePackageSignatureValidationState.NotPresent, options.ValidateCertificateTrust,
                    IsRevocationRequired(options), findings);
            }

            byte[] signatureXml = archive.ReadPart(part.Uri, options.Inspection.MaxSignatureBytes);
            var request = new XmlDigitalSignatureVerificationRequest(signatureXml, certificates) {
                MaxSignatureBytes = options.Inspection.MaxSignatureBytes,
                MaxReferences = options.Inspection.MaxSignedReferences,
                MaxTotalDigestWorkBytes = options.Inspection.MaxTotalDigestBytes
            };
            XmlDigitalSignatureVerificationResult xml = provider.VerifyXmlSignature(request);
            findings.AddRange(xml.Findings);
            OfficePackageSignatureValidationState crypto = Map(xml.Status);
            if (xml.MatchingCertificates.Count == 0) {
                return Result(part, crypto,
                    OfficePackageSignatureValidationState.NotPresent,
                    OfficePackageSignatureValidationState.NotPresent, options.ValidateCertificateTrust,
                    IsRevocationRequired(options), findings);
            }

            X509Certificate2 signer = xml.MatchingCertificates[0];
            if (!options.ValidateCertificateTrust) {
                return Result(part, crypto,
                    OfficePackageSignatureValidationState.NotChecked,
                    OfficePackageSignatureValidationState.NotChecked, false, false, findings);
            }
            CertificateTrustValidationResult trust = provider.ValidateCertificate(
                signer, ExcludeSigner(certificates, signer), options.CertificateValidation,
                CertificateValidationPurpose.DocumentSigning);
            findings.AddRange(trust.Findings);
            return Result(part, crypto,
                Map(trust.Validation.ChainStatus),
                Map(trust.Validation.RevocationStatus), true, IsRevocationRequired(options), findings);
        } catch (Exception exception) when (exception is InvalidDataException or IOException or System.Security.Cryptography.CryptographicException or NotSupportedException) {
            findings.Add(new SecurityFinding(SecurityFindingSeverity.Error,
                "OpcSignatureValidationFailed", exception.Message));
            return Result(part, OfficePackageSignatureValidationState.Failed,
                OfficePackageSignatureValidationState.NotChecked,
                OfficePackageSignatureValidationState.NotChecked, options.ValidateCertificateTrust,
                IsRevocationRequired(options), findings);
        } finally {
            foreach (X509Certificate2 certificate in certificates) certificate.Dispose();
        }
    }

    private static OfficePackageSignaturePartValidationResult Result(
        OfficePackageSignaturePartInfo part,
        OfficePackageSignatureValidationState cryptographicStatus,
        OfficePackageSignatureValidationState certificateStatus,
        OfficePackageSignatureValidationState revocationStatus,
        bool certificateTrustRequired,
        bool revocationRequired,
        IReadOnlyList<SecurityFinding> findings) =>
        new(part, cryptographicStatus, certificateStatus, revocationStatus,
            certificateTrustRequired, revocationRequired, findings.ToArray());

    private static bool IsRevocationRequired(OfficePackageSignatureValidationOptions options) =>
        options.ValidateCertificateTrust && options.CertificateValidation.RevocationMode != X509RevocationMode.NoCheck;

    private static OfficePackageSignatureTimestampInfo[] ReadTimestamps(XElement signature) {
        XNamespace packageSignature = "http://schemas.openxmlformats.org/package/2006/digital-signature";
        var timestamps = new List<OfficePackageSignatureTimestampInfo>();
        foreach (XElement element in signature.Descendants()) {
            if (element.Name == packageSignature + "SignatureTime") {
                timestamps.Add(new OfficePackageSignatureTimestampInfo(
                    "SignatureTime",
                    element.Element(packageSignature + "Value")?.Value.Trim(),
                    element.Element(packageSignature + "Format")?.Value.Trim()));
                continue;
            }
            if (element.Name.LocalName == "SigningTime" &&
                element.Name.NamespaceName.StartsWith("http://uri.etsi.org/01903/", StringComparison.Ordinal)) {
                timestamps.Add(new OfficePackageSignatureTimestampInfo("SigningTime", element.Value.Trim(), null));
                continue;
            }
            if (element.Name.LocalName == "EncapsulatedTimeStamp" &&
                element.Name.NamespaceName.StartsWith("http://uri.etsi.org/01903/", StringComparison.Ordinal)) {
                timestamps.Add(new OfficePackageSignatureTimestampInfo("EncapsulatedTimeStamp", null, null));
            }
        }
        return timestamps.ToArray();
    }

    private static byte[][] ReadCertificates(
        XElement signature,
        OfficePackageSignatureInspectionOptions options,
        OfficePackageCertificateByteBudget certificateBudget) {
        var certificates = new List<byte[]>();
        foreach (XElement element in signature.Descendants(DigitalSignatureNamespace + "X509Certificate")) {
            if (certificates.Count >= options.MaxCertificates) {
                throw new InvalidDataException("The XML signature contains more certificates than the configured limit.");
            }
            string value = element.Value.Trim();
            if (OfficePackageBase64.ExceedsDecodedByteLimit(value, options.MaxCertificateBytes)) {
                throw new InvalidDataException("An embedded signer certificate exceeds the configured byte limit.");
            }
            byte[] encoded = Convert.FromBase64String(value);
            if (encoded.LongLength > options.MaxCertificateBytes) {
                throw new InvalidDataException("An embedded signer certificate exceeds the configured byte limit.");
            }
            certificateBudget.Reserve(encoded.LongLength);
            certificates.Add(encoded);
        }
        return certificates.ToArray();
    }

    private static IReadOnlyList<byte[]> ReadRelatedCertificates(
        OfficePackageSignatureArchive archive,
        string signatureUri,
        OfficePackageSignatureInspectionOptions options,
        OfficePackageCertificateByteBudget certificateBudget) {
        const string certificateRelationshipType =
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate";
        string relationshipsUri = GetRelationshipPartUri(signatureUri);
        if (!archive.ContainsPart(relationshipsUri)) return Array.Empty<byte[]>();
        XNamespace relationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        XDocument relationships = LoadXml(archive.ReadPart(relationshipsUri, options.MaxSignatureBytes));
        XElement? root = relationships.Root;
        if (root?.Name != relationshipsNamespace + "Relationships") {
            throw new InvalidDataException("The XML signature relationships part has an invalid root element.");
        }
        XElement[] declarations = root.Elements(relationshipsNamespace + "Relationship")
            .Where(element => string.Equals((string?)element.Attribute("Type"),
                certificateRelationshipType, StringComparison.Ordinal))
            .Take(options.MaxCertificates + 1)
            .ToArray();
        if (declarations.Length > options.MaxCertificates) {
            throw new InvalidDataException("The XML signature has more related certificates than the configured limit.");
        }
        var certificates = new List<byte[]>(declarations.Length);
        foreach (XElement relationship in declarations) {
            if (string.Equals((string?)relationship.Attribute("TargetMode"), "External",
                    StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException("An external signature-certificate relationship was rejected.");
            }
            string? target = (string?)relationship.Attribute("Target");
            if (string.IsNullOrWhiteSpace(target)) {
                throw new InvalidDataException("A signature-certificate relationship has no target.");
            }
            string certificateUri = ResolveRelationshipTarget(signatureUri, target!);
            byte[] encoded = archive.ReadPart(certificateUri, options.MaxCertificateBytes);
            certificateBudget.Reserve(encoded.LongLength);
            certificates.Add(encoded);
        }
        return certificates;
    }

    private static IReadOnlyList<X509Certificate2> ExcludeSigner(
        IReadOnlyList<X509Certificate2> certificates,
        X509Certificate2 signer) => certificates
        .Where(certificate => !certificate.RawData.SequenceEqual(signer.RawData))
        .ToArray();

    private static HashSet<string> ReadSignatureRelationships(
        OfficePackageSignatureArchive archive,
        OriginDiscovery origins,
        OfficePackageSignatureInspectionOptions options,
        ICollection<string> findings) {
        var signatureUris = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (origins.RelationshipCount != 1 || origins.ExistingPartUris.Count != 1) {
            if (origins.RelationshipCount > 0) {
                findings.Add("A valid signature chain requires exactly one internal root signature-origin relationship and one existing origin part.");
            }
            return signatureUris;
        }

        string originUri = origins.ExistingPartUris[0];
        string relationshipsUri = GetRelationshipPartUri(originUri);
        if (!archive.ContainsPart(relationshipsUri)) {
            findings.Add("The signature-origin part does not have a relationships part.");
            return signatureUris;
        }
        try {
            XNamespace relationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
            XDocument relationships = LoadXml(archive.ReadPart(relationshipsUri, options.MaxSignatureBytes));
            XElement? root = relationships.Root;
            if (root?.Name != relationshipsNamespace + "Relationships") {
                throw new InvalidDataException("The signature-origin relationships part has an invalid root element.");
            }
            XElement[] declarations = root.Elements(relationshipsNamespace + "Relationship")
                .Where(element => string.Equals((string?)element.Attribute("Type"),
                    SignatureRelationshipType, StringComparison.Ordinal))
                .Take(options.MaxSignatureParts + 1)
                .ToArray();
            if (declarations.Length > options.MaxSignatureParts) {
                throw new InvalidDataException("The signature origin exceeds the configured signature-relationship limit.");
            }
            foreach (XElement relationship in declarations) {
                if (string.Equals((string?)relationship.Attribute("TargetMode"), "External",
                        StringComparison.OrdinalIgnoreCase)) {
                    findings.Add("An external signature relationship was rejected.");
                    continue;
                }
                string? target = (string?)relationship.Attribute("Target");
                if (string.IsNullOrWhiteSpace(target)) {
                    findings.Add("A signature relationship has no target.");
                    continue;
                }
                string resolved = ResolveRelationshipTarget(originUri, target!);
                if (!archive.ContainsPart(resolved)) {
                    findings.Add("A signature relationship targets missing part " + resolved + ".");
                    continue;
                }
                if (!signatureUris.Add(resolved)) {
                    findings.Add("The signature origin contains duplicate relationships to " + resolved + ".");
                }
            }
        } catch (Exception exception) when (exception is IOException or InvalidDataException or XmlException or UriFormatException) {
            findings.Add("The signature-origin relationships part could not be parsed: " + exception.Message);
            signatureUris.Clear();
        }
        return signatureUris;
    }

    private static string GetRelationshipPartUri(string sourcePartUri) {
        string normalized = OfficePackageSignatureArchive.NormalizePartUri(sourcePartUri);
        int slash = normalized.LastIndexOf('/');
        string directory = slash <= 0 ? "/" : normalized.Substring(0, slash + 1);
        string fileName = normalized.Substring(slash + 1);
        return directory + "_rels/" + fileName + ".rels";
    }

    private static string ResolveRelationshipTarget(string sourcePartUri, string target) {
        if (target.IndexOf('#') >= 0 || target.IndexOf('?') >= 0) {
            throw new InvalidDataException("A package signature relationship target cannot contain a query or fragment.");
        }
        var source = new Uri("http://officeimo.package" +
            OfficePackageSignatureArchive.NormalizePartUri(sourcePartUri), UriKind.Absolute);
        var resolved = new Uri(source, target);
        if (!string.Equals(resolved.Scheme, source.Scheme, StringComparison.Ordinal) ||
            !string.Equals(resolved.Host, source.Host, StringComparison.Ordinal)) {
            throw new InvalidDataException("A package signature relationship target leaves the package namespace.");
        }
        return OfficePackageSignatureArchive.NormalizePartUri(Uri.UnescapeDataString(resolved.AbsolutePath));
    }

    private static OriginDiscovery ReadOrigins(
        OfficePackageSignatureArchive archive,
        OfficePackageSignatureInspectionOptions options,
        ICollection<string> findings) {
        const string rootRelationships = "/_rels/.rels";
        if (!archive.ContainsPart(rootRelationships)) return new OriginDiscovery(0, Array.Empty<string>(), null);
        try {
            XDocument relationships = LoadXml(archive.ReadPart(rootRelationships, options.MaxSignatureBytes));
            XNamespace relationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
            XElement? root = relationships.Root;
            if (root == null || root.Name != relationshipsNamespace + "Relationships") {
                throw new InvalidDataException("The package root relationships part has an invalid root element.");
            }
            XElement[] declarations = root.Elements(relationshipsNamespace + "Relationship")
                .Where(element => string.Equals((string?)element.Attribute("Type"),
                    OriginRelationshipType, StringComparison.Ordinal))
                .Take(options.MaxPackageParts + 1)
                .ToArray();
            if (declarations.Length > options.MaxPackageParts) {
                throw new InvalidDataException("The root relationship part exceeds the configured relationship limit.");
            }
            if (declarations.Length > 1) findings.Add("The package declares more than one signature-origin relationship.");
            var partUris = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (XElement relationship in declarations) {
                if (string.Equals((string?)relationship.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) {
                    findings.Add("A signature-origin relationship is external and was rejected.");
                    continue;
                }
                string? target = (string?)relationship.Attribute("Target");
                if (!string.IsNullOrWhiteSpace(target)) {
                    try {
                        partUris.Add(ResolveRelationshipTarget("/", target!));
                    } catch (Exception exception) when (exception is InvalidDataException or UriFormatException) {
                        findings.Add("A signature-origin relationship target is invalid: " + exception.Message);
                    }
                }
            }
            string? first = partUris.OrderBy(uri => uri, StringComparer.OrdinalIgnoreCase).FirstOrDefault();
            string[] existing = partUris.Where(archive.ContainsPart)
                .OrderBy(uri => uri, StringComparer.OrdinalIgnoreCase).ToArray();
            return new OriginDiscovery(declarations.Length, existing, first);
        } catch (Exception exception) when (exception is IOException or InvalidDataException or XmlException or UriFormatException) {
            findings.Add("The root relationship part could not be parsed: " + exception.Message);
            return new OriginDiscovery(0, Array.Empty<string>(), null);
        }
    }

    private readonly struct OriginDiscovery {
        internal OriginDiscovery(int relationshipCount, IReadOnlyList<string> existingPartUris, string? firstOriginUri) {
            RelationshipCount = relationshipCount;
            ExistingPartUris = existingPartUris;
            FirstOriginUri = firstOriginUri;
        }
        internal int RelationshipCount { get; }
        internal IReadOnlyList<string> ExistingPartUris { get; }
        internal int ExistingPartCount => ExistingPartUris.Count;
        internal string? FirstOriginUri { get; }
    }

    private static bool ReadApplicationSignatureMetadata(
        OfficePackageSignatureArchive archive,
        OfficePackageSignatureInspectionOptions options,
        ICollection<string> findings) {
        const string rootRelationships = "/_rels/.rels";
        if (!archive.ContainsPart(rootRelationships)) return false;
        try {
            XDocument relationships = LoadXml(archive.ReadPart(rootRelationships, options.MaxSignatureBytes));
            XNamespace relationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
            XElement? root = relationships.Root;
            if (root == null || root.Name != relationshipsNamespace + "Relationships") {
                throw new InvalidDataException("The package root relationships part has an invalid root element.");
            }
            XElement[] declarations = root.Elements(relationshipsNamespace + "Relationship")
                .Where(element => string.Equals(
                    (string?)element.Attribute("Type"),
                    ExtendedPropertiesRelationshipType,
                    StringComparison.Ordinal))
                .Take(options.MaxPackageParts + 1)
                .ToArray();
            if (declarations.Length > options.MaxPackageParts) {
                throw new InvalidDataException("The root relationship part exceeds the configured relationship limit.");
            }
            XNamespace extendedProperties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
            foreach (XElement declaration in declarations) {
                if (string.Equals((string?)declaration.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                string? target = (string?)declaration.Attribute("Target");
                if (string.IsNullOrWhiteSpace(target)) continue;
                string partUri = ResolveRelationshipTarget("/", target!);
                if (!archive.ContainsPart(partUri)) continue;
                if (!archive.TryGetContentType(partUri, out string contentType) ||
                    !string.Equals(contentType, ExtendedPropertiesContentType, StringComparison.OrdinalIgnoreCase)) {
                    findings.Add("An extended-properties relationship target has an unexpected OPC content type.");
                    continue;
                }
                XDocument properties = LoadXml(archive.ReadPart(partUri, options.MaxSignatureBytes));
                if (properties.Root?.Name == extendedProperties + "Properties" &&
                    properties.Root.Elements(extendedProperties + "DigSig").Any()) {
                    return true;
                }
            }
            return false;
        } catch (Exception exception) when (exception is IOException or InvalidDataException or XmlException or UriFormatException) {
            findings.Add("Extended application properties could not be parsed: " + exception.Message);
            return false;
        }
    }

    private static XDocument LoadXml(byte[] bytes) {
        using var stream = new MemoryStream(bytes, writable: false);
        var settings = new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = Math.Max(1, bytes.LongLength)
        };
        using XmlReader reader = XmlReader.Create(stream, settings);
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    private static byte[] ReadPackage(string filePath, OfficePackageSignatureInspectionOptions options) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A package path is required.", nameof(filePath));
        options.Validate();
        string fullPath = Path.GetFullPath(filePath);
        var info = new FileInfo(fullPath);
        if (!info.Exists) throw new FileNotFoundException("The package file does not exist.", fullPath);
        if (info.Length > options.MaxPackageBytes) throw new InvalidDataException("The OPC package exceeds the configured byte limit.");
        return File.ReadAllBytes(fullPath);
    }

    private static OfficePackageSignatureValidationState Map(SecurityValidationStatus status) => status switch {
        SecurityValidationStatus.Valid => OfficePackageSignatureValidationState.Passed,
        SecurityValidationStatus.Invalid => OfficePackageSignatureValidationState.Failed,
        SecurityValidationStatus.NotPerformed => OfficePackageSignatureValidationState.NotChecked,
        _ => OfficePackageSignatureValidationState.Unsupported
    };
}
