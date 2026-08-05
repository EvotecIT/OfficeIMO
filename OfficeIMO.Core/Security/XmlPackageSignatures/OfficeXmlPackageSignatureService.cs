using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Security;

/// <summary>Provider-backed, bounded XML package manifests for ODF and EPUB signature carriers.</summary>
public static class OfficeXmlPackageSignatureService {
    private static readonly XNamespace Ds = XmlDigitalSignatureAlgorithms.Namespace;
    private static readonly XNamespace PackageManifestNamespace = "urn:officeimo:security:package-manifest:1";
    private static readonly XNamespace OdfSignatureNamespace =
        "urn:oasis:names:tc:opendocument:xmlns:digitalsignature:1.0";

    /// <summary>Creates a signature, validates the staged package, and commits it atomically.</summary>
    public static OfficeXmlPackageSigningResult Sign(
        string filePath,
        OfficeXmlPackageSignatureFormat format,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficeXmlPackageSignatureOptions? options = null) {
        OfficeXmlPackageSigningResult result = TrySign(filePath, format, securityProvider, signingCertificate, options);
        if (!result.Succeeded) throw new InvalidOperationException(string.Join(" ", result.Findings));
        return result;
    }

    /// <summary>Attempts atomic signature creation and returns structured failure evidence.</summary>
    public static OfficeXmlPackageSigningResult TrySign(
        string filePath,
        OfficeXmlPackageSignatureFormat format,
        IOfficeSecurityProvider securityProvider,
        X509Certificate2 signingCertificate,
        OfficeXmlPackageSignatureOptions? options = null) {
        if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
        if (signingCertificate == null) throw new ArgumentNullException(nameof(signingCertificate));
        options ??= new OfficeXmlPackageSignatureOptions();
        ValidateOptions(options);
        string fullPath = NormalizePath(filePath);
        var findings = new List<string>();
        if (!File.Exists(fullPath)) return Failed(fullPath, "The package file does not exist.");
        if (!signingCertificate.HasPrivateKey) return Failed(fullPath, "The signing certificate does not have a private key.");
        if (new FileInfo(fullPath).Length > options.MaxPackageBytes) return Failed(fullPath, "The package exceeds the configured byte limit.");

        string stagingPath = string.Empty;
        try {
            stagingPath = OfficeFileCommit.CreateStagingPath(fullPath);
            OfficePackageFileSnapshot.CopyBounded(fullPath, stagingPath, options.MaxPackageBytes);
            string sourceHash = OfficePackageFileSnapshot.ComputeSha256(stagingPath, options.MaxPackageBytes);
            string carrierPath = CarrierPath(format);
            XDocument carrier;
            using (var archive = ZipFile.OpenRead(stagingPath)) {
                ValidateArchive(archive, options);
                carrier = ReadOrCreateCarrier(archive, format, options);
                XElement manifest = CreateManifest(archive, carrierPath, format, options);
                string signatureId = "OfficeIMO" + Guid.NewGuid().ToString("N");
                string objectId = "PackageManifest" + Guid.NewGuid().ToString("N");
                byte[] wrapper = Encoding.UTF8.GetBytes(new XElement("Root", manifest).ToString(SaveOptions.DisableFormatting));
                var request = new XmlDigitalSignatureCreationRequest(
                    wrapper, signingCertificate, signatureId, objectId,
                    "urn:officeimo:security:package-manifest:1",
                    XmlDigitalSignatureAlgorithms.CanonicalXml,
                    XmlDigitalSignatureAlgorithms.RsaSha256,
                    XmlDigitalSignatureAlgorithms.Sha256) {
                    AdditionalCertificates = options.AdditionalCertificates,
                    MaxObjectBytes = options.MaxSignatureBytes,
                    MaxOutputBytes = options.MaxSignatureBytes
                };
                byte[] signatureXml = securityProvider.CreateXmlSignature(request);
                XDocument signatureDocument = LoadXml(signatureXml, options.MaxSignatureBytes);
                carrier.Root!.Add(signatureDocument.Root!);
            }
            using (var archive = ZipFile.Open(stagingPath, ZipArchiveMode.Update)) {
                WriteCarrier(archive, carrierPath, carrier, options);
            }

            var readbackOptions = CloneOptions(options);
            readbackOptions.ValidateCertificateTrust = false;
            OfficeXmlPackageSignatureValidationReport validation = Validate(
                stagingPath, format, securityProvider, readbackOptions);
            if (validation.Signatures.LastOrDefault()?.IsValidUnderPolicy != true) {
                findings.Add("The created XML package signature failed bounded validation readback.");
                findings.AddRange(validation.Findings);
                return new OfficeXmlPackageSigningResult(fullPath, false, validation.Signatures.Count, validation, findings);
            }
            string validatedHash = OfficePackageFileSnapshot.ComputeSha256(stagingPath, options.MaxPackageBytes);
            if (!OfficeFileCommit.TryCommitTemporaryFileAtomicallyIfDestinationUnchanged(
                stagingPath, fullPath,
                displaced => string.Equals(sourceHash,
                    OfficePackageFileSnapshot.ComputeSha256(displaced, options.MaxPackageBytes), StringComparison.Ordinal),
                installed => string.Equals(validatedHash,
                    OfficePackageFileSnapshot.ComputeSha256(installed, options.MaxPackageBytes), StringComparison.Ordinal))) {
                stagingPath = string.Empty;
                return Failed(fullPath, "The package changed while its XML signature was staged; the current source was preserved.");
            }
            stagingPath = string.Empty;
            findings.Add("The XML package signature was validated and atomically committed.");
            return new OfficeXmlPackageSigningResult(fullPath, true, validation.Signatures.Count, validation, findings);
        } catch (Exception exception) when (exception is IOException or InvalidDataException or XmlException or
            CryptographicException or UnauthorizedAccessException or ArgumentException or NotSupportedException) {
            return Failed(fullPath, "XML package signing failed before atomic commit. " + exception.Message);
        } finally {
            if (!string.IsNullOrWhiteSpace(stagingPath)) OfficeFileCommit.DeleteIfExists(stagingPath);
        }
    }

    /// <summary>Validates an ODF or EPUB XML signature carrier and its signed entry manifest.</summary>
    public static OfficeXmlPackageSignatureValidationReport Validate(
        string filePath,
        OfficeXmlPackageSignatureFormat format,
        IOfficeSecurityProvider securityProvider,
        OfficeXmlPackageSignatureOptions? options = null) {
        if (securityProvider == null) throw new ArgumentNullException(nameof(securityProvider));
        options ??= new OfficeXmlPackageSignatureOptions();
        ValidateOptions(options);
        string fullPath = NormalizePath(filePath);
        string carrierPath = CarrierPath(format);
        var findings = new List<string>();
        if (!File.Exists(fullPath)) throw new FileNotFoundException("The package file does not exist.", fullPath);
        if (new FileInfo(fullPath).Length > options.MaxPackageBytes) throw new InvalidDataException("The package exceeds the configured byte limit.");

        try {
            using var archive = ZipFile.OpenRead(fullPath);
            ValidateArchive(archive, options);
            ZipArchiveEntry? carrierEntry = FindEntry(archive, carrierPath);
            if (carrierEntry == null) {
                return new OfficeXmlPackageSignatureValidationReport(fullPath, carrierPath, false, true,
                    Array.Empty<OfficeXmlPackageSignatureResult>(), new[] { "The package does not contain a signature carrier." });
            }
            XDocument carrier = LoadXml(ReadEntry(carrierEntry, options.MaxSignatureBytes), options.MaxSignatureBytes);
            ValidateCarrierStructure(carrier, format);
            XElement[] signatures = carrier.Root!.Elements(Ds + "Signature").Take(options.MaxSignatures + 1).ToArray();
            if (signatures.Length > options.MaxSignatures) throw new InvalidDataException("The signature carrier exceeds the configured signature-count limit.");
            var results = signatures.Select(signature => ValidateSignature(
                archive, carrierPath, format, signature, securityProvider, options)).ToArray();
            if (results.Length == 0) findings.Add("The signature carrier does not contain XML signatures.");
            return new OfficeXmlPackageSignatureValidationReport(fullPath, carrierPath, true, true, results, findings);
        } catch (Exception exception) when (exception is IOException or InvalidDataException or XmlException or CryptographicException) {
            findings.Add(exception.Message);
            return new OfficeXmlPackageSignatureValidationReport(fullPath, carrierPath, true, false,
                Array.Empty<OfficeXmlPackageSignatureResult>(), findings);
        }
    }

    private static OfficeXmlPackageSignatureResult ValidateSignature(
        ZipArchive archive,
        string carrierPath,
        OfficeXmlPackageSignatureFormat format,
        XElement signature,
        IOfficeSecurityProvider provider,
        OfficeXmlPackageSignatureOptions options) {
        var findings = new List<SecurityFinding>();
        string? signatureId = (string?)signature.Attribute("Id");
        int maxAuthenticatedReferences = options.MaxEntries == int.MaxValue
            ? int.MaxValue
            : options.MaxEntries + 1;
        OfficeXmlSignatureBinding.AuthenticatedContent authenticated = OfficeXmlSignatureBinding.Resolve(
            signature, PackageManifestNamespace + "PackageManifest", maxAuthenticatedReferences);
        if (authenticated.Payloads.Count != 1) {
            throw new InvalidDataException("The OfficeIMO XML signature profile requires exactly one authenticated package manifest.");
        }
        IReadOnlyList<OfficeXmlPackageEntryDigestResult> entries = ValidateManifest(
            archive, carrierPath, format, authenticated.Payloads[0], options);
        X509Certificate2[] certificates;
        var parsedCertificates = new List<X509Certificate2>();
        try {
            var certificateBudget = new OfficePackageCertificateByteBudget(options.MaxTotalCertificateBytes);
            foreach (XElement element in signature.Descendants(Ds + "X509Certificate")) {
                if (parsedCertificates.Count >= options.MaxCertificates) {
                    throw new InvalidDataException("The XML signature exceeds the configured certificate-count limit.");
                }
                string encodedValue = element.Value.Trim();
                if (OfficePackageBase64.ExceedsDecodedByteLimit(encodedValue, options.MaxCertificateBytes)) {
                    throw new InvalidDataException("An XML signature certificate exceeds the configured byte limit.");
                }
                byte[] encodedCertificate = Convert.FromBase64String(encodedValue);
                certificateBudget.Reserve(encodedCertificate.LongLength);
                parsedCertificates.Add(OfficePackageCertificateLoader.Load(encodedCertificate));
            }
            certificates = parsedCertificates.ToArray();
        } catch (Exception exception) when (exception is FormatException or CryptographicException or InvalidDataException) {
            foreach (X509Certificate2 certificate in parsedCertificates) certificate.Dispose();
            findings.Add(new SecurityFinding(SecurityFindingSeverity.Error, "XmlSignerCertificateMalformed", exception.Message));
            return new OfficeXmlPackageSignatureResult(signatureId, OfficePackageSignatureValidationState.Failed,
                OfficePackageSignatureValidationState.NotPresent, OfficePackageSignatureValidationState.NotPresent,
                options.ValidateCertificateTrust, entries, findings);
        }
        try {
            byte[] signatureXml = Encoding.UTF8.GetBytes(signature.ToString(SaveOptions.DisableFormatting));
            XmlDigitalSignatureVerificationResult xml = provider.VerifyXmlSignature(
                new XmlDigitalSignatureVerificationRequest(signatureXml, certificates) {
                    MaxSignatureBytes = options.MaxSignatureBytes,
                    MaxReferences = options.MaxEntries,
                    MaxTotalDigestWorkBytes = options.MaxTotalDigestBytes
                });
            findings.AddRange(xml.Findings);
            OfficePackageSignatureValidationState crypto = Map(xml.Status);
            if (xml.MatchingCertificates.Count == 0) {
                return new OfficeXmlPackageSignatureResult(signatureId, crypto,
                    OfficePackageSignatureValidationState.NotPresent, OfficePackageSignatureValidationState.NotPresent,
                    options.ValidateCertificateTrust, entries, findings);
            }
            if (!options.ValidateCertificateTrust) {
                return new OfficeXmlPackageSignatureResult(signatureId, crypto,
                    OfficePackageSignatureValidationState.NotChecked, OfficePackageSignatureValidationState.NotChecked,
                    false, entries, findings);
            }
            CertificateTrustValidationResult trust = provider.ValidateCertificate(
                xml.MatchingCertificates[0], ExcludeSigner(certificates, xml.MatchingCertificates[0]), options.CertificateValidation,
                CertificateValidationPurpose.DocumentSigning);
            findings.AddRange(trust.Findings);
            return new OfficeXmlPackageSignatureResult(signatureId, crypto,
                Map(trust.Validation.ChainStatus), Map(trust.Validation.RevocationStatus),
                true, entries, findings);
        } finally {
            foreach (X509Certificate2 certificate in certificates) certificate.Dispose();
        }
    }

    private static IReadOnlyList<OfficeXmlPackageEntryDigestResult> ValidateManifest(
        ZipArchive archive, string carrierPath, OfficeXmlPackageSignatureFormat format,
        XElement manifest, OfficeXmlPackageSignatureOptions options) {
        string? declaredFormat = (string?)manifest.Attribute("Format");
        if (!string.Equals(declaredFormat, format.ToString(), StringComparison.Ordinal)) {
            return new[] { new OfficeXmlPackageEntryDigestResult(string.Empty, false,
                OfficePackageSignatureValidationState.Failed, "The signed manifest format does not match the package host.") };
        }
        var results = new List<OfficeXmlPackageEntryDigestResult>();
        var declared = new HashSet<string>(StringComparer.Ordinal);
        long totalBytes = 0;
        XElement[] declaredEntries = manifest.Elements(PackageManifestNamespace + "Entry")
            .Take(options.MaxEntries + 1).ToArray();
        if (declaredEntries.Length > options.MaxEntries) {
            throw new InvalidDataException("The signed manifest exceeds the configured entry-count limit.");
        }
        foreach (XElement item in declaredEntries) {
            string path = NormalizeEntryPath((string?)item.Attribute("Path"));
            if (path.Length == 0 || !declared.Add(path)) {
                results.Add(new OfficeXmlPackageEntryDigestResult(path, false,
                    OfficePackageSignatureValidationState.Failed, "The signed manifest contains an empty or duplicate entry path."));
                continue;
            }
            ZipArchiveEntry? entry = FindEntry(archive, path);
            if (entry == null) {
                results.Add(new OfficeXmlPackageEntryDigestResult(path, false,
                    OfficePackageSignatureValidationState.Failed, "The signed package entry is missing."));
                continue;
            }
            byte[] bytes = ReadEntry(entry, options.MaxEntryBytes);
            totalBytes = checked(totalBytes + bytes.LongLength);
            if (totalBytes > options.MaxTotalDigestBytes) throw new InvalidDataException("Package digest work exceeds the configured aggregate limit.");
            string actual = Sha256(bytes);
            string? expected = ((string?)item.Attribute("DigestValue"))?.Trim();
            bool matches = string.Equals(expected, actual, StringComparison.Ordinal);
            results.Add(new OfficeXmlPackageEntryDigestResult(path, true,
                matches ? OfficePackageSignatureValidationState.Passed : OfficePackageSignatureValidationState.Failed,
                matches ? "Package entry digest matches." : "Package entry digest does not match."));
        }
        foreach (string unsigned in archive.Entries.Where(entry => entry.Name.Length > 0)
            .Select(entry => NormalizeEntryPath(entry.FullName))
            .Where(path => !string.Equals(path, carrierPath, StringComparison.Ordinal) && !declared.Contains(path))) {
            results.Add(new OfficeXmlPackageEntryDigestResult(unsigned, true,
                OfficePackageSignatureValidationState.Failed, "The package contains an unsigned entry."));
        }
        return results;
    }

    private static XElement CreateManifest(ZipArchive archive, string carrierPath,
        OfficeXmlPackageSignatureFormat format, OfficeXmlPackageSignatureOptions options) {
        var manifest = new XElement(PackageManifestNamespace + "PackageManifest",
            new XAttribute("Format", format.ToString()),
            new XAttribute("DigestMethod", XmlDigitalSignatureAlgorithms.Sha256));
        long totalBytes = 0;
        ZipArchiveEntry[] entries = archive.Entries.Where(entry => entry.Name.Length > 0 &&
                !string.Equals(NormalizeEntryPath(entry.FullName), carrierPath, StringComparison.Ordinal))
            .OrderBy(entry => entry.FullName, StringComparer.Ordinal)
            .ToArray();
        if (entries.Length > options.MaxEntries) throw new InvalidDataException("The package exceeds the configured entry-count limit.");
        foreach (ZipArchiveEntry entry in entries) {
            byte[] bytes = ReadEntry(entry, options.MaxEntryBytes);
            totalBytes = checked(totalBytes + bytes.LongLength);
            if (totalBytes > options.MaxTotalDigestBytes) throw new InvalidDataException("Package digest work exceeds the configured aggregate limit.");
            manifest.Add(new XElement(PackageManifestNamespace + "Entry",
                new XAttribute("Path", NormalizeEntryPath(entry.FullName)),
                new XAttribute("DigestValue", Sha256(bytes))));
        }
        return manifest;
    }

    private static XDocument ReadOrCreateCarrier(ZipArchive archive,
        OfficeXmlPackageSignatureFormat format, OfficeXmlPackageSignatureOptions options) {
        ZipArchiveEntry? entry = FindEntry(archive, CarrierPath(format));
        if (entry != null) {
            XDocument existing = LoadXml(ReadEntry(entry, options.MaxSignatureBytes), options.MaxSignatureBytes);
            ValidateCarrierStructure(existing, format);
            return existing;
        }
        return new XDocument(new XElement(ExpectedCarrierRootName(format)));
    }

    private static void ValidateCarrierStructure(XDocument carrier, OfficeXmlPackageSignatureFormat format) {
        XElement? root = carrier.Root;
        if (root == null || root.Name != ExpectedCarrierRootName(format)) {
            throw new InvalidDataException("The XML signature carrier root does not match the selected package format.");
        }
        if (root.Elements().Any(element => element.Name != Ds + "Signature") ||
            root.Descendants(Ds + "Signature").Any(signature => signature.Parent != root)) {
            throw new InvalidDataException("The XML signature carrier contains content outside the bounded OfficeIMO signature profile.");
        }
    }

    private static XName ExpectedCarrierRootName(OfficeXmlPackageSignatureFormat format) =>
        OdfSignatureNamespace + (format == OfficeXmlPackageSignatureFormat.OpenDocument
            ? "document-signatures"
            : format == OfficeXmlPackageSignatureFormat.Epub
                ? "signatures"
                : throw new ArgumentOutOfRangeException(nameof(format)));

    private static IReadOnlyList<X509Certificate2> ExcludeSigner(
        IReadOnlyList<X509Certificate2> certificates,
        X509Certificate2 signer) => certificates
        .Where(certificate => !certificate.RawData.SequenceEqual(signer.RawData))
        .ToArray();

    private static void WriteCarrier(ZipArchive archive, string carrierPath,
        XDocument carrier, OfficeXmlPackageSignatureOptions options) {
        byte[] bytes = Encoding.UTF8.GetBytes(carrier.ToString(SaveOptions.DisableFormatting));
        if (bytes.LongLength > options.MaxSignatureBytes) throw new InvalidDataException("The generated signature carrier exceeds the configured byte limit.");
        FindEntry(archive, carrierPath)?.Delete();
        ZipArchiveEntry entry = archive.CreateEntry(carrierPath, CompressionLevel.Optimal);
        using Stream output = entry.Open();
        output.Write(bytes, 0, bytes.Length);
    }

    private static void ValidateArchive(ZipArchive archive, OfficeXmlPackageSignatureOptions options) {
        if (archive.Entries.Count > options.MaxEntries + 1) throw new InvalidDataException("The package exceeds the configured entry-count limit.");
        var paths = new HashSet<string>(StringComparer.Ordinal);
        foreach (ZipArchiveEntry entry in archive.Entries.Where(entry => entry.Name.Length > 0)) {
            string normalized = NormalizeEntryPath(entry.FullName);
            if (!paths.Add(normalized)) throw new InvalidDataException("The package contains duplicate entry path '" + normalized + "'.");
            if (entry.Length > options.MaxEntryBytes && !normalized.StartsWith("META-INF/", StringComparison.Ordinal)) {
                throw new InvalidDataException("Package entry '" + normalized + "' exceeds the configured byte limit.");
            }
        }
    }

    private static byte[] ReadEntry(ZipArchiveEntry entry, long maxBytes) {
        if (entry.Length > maxBytes) throw new InvalidDataException("Package entry '" + entry.FullName + "' exceeds the configured byte limit.");
        using Stream input = entry.Open();
        using var output = new MemoryStream(entry.Length > int.MaxValue ? 0 : (int)entry.Length);
        var buffer = new byte[81920];
        long total = 0;
        while (true) {
            int read = input.Read(buffer, 0, buffer.Length);
            if (read == 0) break;
            total = checked(total + read);
            if (total > maxBytes) throw new InvalidDataException("Package entry expanded beyond the configured byte limit.");
            output.Write(buffer, 0, read);
        }
        return output.ToArray();
    }

    private static XDocument LoadXml(byte[] bytes, long maxBytes) {
        if (bytes.LongLength > maxBytes) throw new InvalidDataException("XML exceeds the configured byte limit.");
        using var stream = new MemoryStream(bytes, writable: false);
        using XmlReader reader = XmlReader.Create(stream, new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
            MaxCharactersInDocument = Math.Max(1, bytes.LongLength)
        });
        return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
    }

    private static ZipArchiveEntry? FindEntry(ZipArchive archive, string path) =>
        archive.Entries.FirstOrDefault(entry =>
            string.Equals(NormalizeEntryPath(entry.FullName), NormalizeEntryPath(path), StringComparison.Ordinal));

    private static string NormalizeEntryPath(string? path) {
        string value = (path ?? string.Empty).Replace('\\', '/').TrimStart('/');
        if (value.Split('/').Any(segment => segment is "" or "." or "..")) return string.Empty;
        return value;
    }

    private static string CarrierPath(OfficeXmlPackageSignatureFormat format) => format switch {
        OfficeXmlPackageSignatureFormat.OpenDocument => "META-INF/documentsignatures.xml",
        OfficeXmlPackageSignatureFormat.Epub => "META-INF/signatures.xml",
        _ => throw new ArgumentOutOfRangeException(nameof(format))
    };

    private static string Sha256(byte[] bytes) {
        using SHA256 algorithm = SHA256.Create();
        return Convert.ToBase64String(algorithm.ComputeHash(bytes));
    }

    private static string NormalizePath(string filePath) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("A package path is required.", nameof(filePath));
        return Path.GetFullPath(filePath);
    }

    private static void ValidateOptions(OfficeXmlPackageSignatureOptions options) {
        if (options.MaxPackageBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxPackageBytes));
        if (options.MaxEntries <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxEntries));
        if (options.MaxEntryBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxEntryBytes));
        if (options.MaxTotalDigestBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxTotalDigestBytes));
        if (options.MaxSignatureBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxSignatureBytes));
        if (options.MaxSignatures <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxSignatures));
        if (options.MaxCertificates <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxCertificates));
        if (options.MaxCertificateBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxCertificateBytes));
        if (options.MaxTotalCertificateBytes <= 0) throw new ArgumentOutOfRangeException(nameof(options.MaxTotalCertificateBytes));
    }

    private static OfficeXmlPackageSigningResult Failed(string path, string finding) =>
        new(path, false, 0, null, new[] { finding });

    private static OfficeXmlPackageSignatureOptions CloneOptions(OfficeXmlPackageSignatureOptions source) {
        var target = new OfficeXmlPackageSignatureOptions {
            MaxPackageBytes = source.MaxPackageBytes,
            MaxEntries = source.MaxEntries,
            MaxEntryBytes = source.MaxEntryBytes,
            MaxTotalDigestBytes = source.MaxTotalDigestBytes,
            MaxSignatureBytes = source.MaxSignatureBytes,
            MaxSignatures = source.MaxSignatures,
            MaxCertificates = source.MaxCertificates,
            MaxCertificateBytes = source.MaxCertificateBytes,
            MaxTotalCertificateBytes = source.MaxTotalCertificateBytes,
            ValidateCertificateTrust = source.ValidateCertificateTrust,
            AdditionalCertificates = source.AdditionalCertificates
        };
        target.CertificateValidation.ValidateChain = source.CertificateValidation.ValidateChain;
        target.CertificateValidation.RevocationMode = source.CertificateValidation.RevocationMode;
        target.CertificateValidation.RevocationFlag = source.CertificateValidation.RevocationFlag;
        target.CertificateValidation.VerificationFlags = source.CertificateValidation.VerificationFlags;
        target.CertificateValidation.DisableCertificateDownloads = source.CertificateValidation.DisableCertificateDownloads;
        target.CertificateValidation.VerificationTime = source.CertificateValidation.VerificationTime;
        target.CertificateValidation.UrlRetrievalTimeout = source.CertificateValidation.UrlRetrievalTimeout;
        target.CertificateValidation.ChainEvaluator = source.CertificateValidation.ChainEvaluator;
        target.CertificateValidation.ExtraCertificates.AddRange(source.CertificateValidation.ExtraCertificates);
        return target;
    }

    private static OfficePackageSignatureValidationState Map(SecurityValidationStatus status) => status switch {
        SecurityValidationStatus.Valid => OfficePackageSignatureValidationState.Passed,
        SecurityValidationStatus.Invalid => OfficePackageSignatureValidationState.Failed,
        SecurityValidationStatus.NotPerformed => OfficePackageSignatureValidationState.NotChecked,
        _ => OfficePackageSignatureValidationState.Unsupported
    };
}
