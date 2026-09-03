using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.Xlsb.Package;
using OfficeIMO.Provenance;
using OfficeIMO.Security;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    private const string SignatureOriginContentType = "application/vnd.openxmlformats-package.digital-signature-origin";
    private const string SignaturePartContentType = "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml";
    private const string SignatureRelationshipPrefix = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/";
    private const string ExtendedPropertiesContentType = "application/vnd.openxmlformats-officedocument.extended-properties+xml";
    private const string ExtendedPropertiesRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

    /// <summary>Inspects C2PA and IPTC provenance in a saved Open XML workbook and its supported embedded images.</summary>
    public static OfficeProvenanceReport InspectProvenance(string filePath, OfficeProvenanceOptions? options = null) =>
        OfficeProvenancePackageMutation.InspectFile(filePath, options, ValidatePackage);

    /// <summary>Removes selected provenance and atomically writes an Open XML workbook.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        string inputPath,
        string outputPath,
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.RemoveFile(inputPath, outputPath, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);

    /// <summary>Removes selected provenance from encoded Open XML workbook bytes.</summary>
    public static OfficeProvenanceRemovalResult RemoveProvenance(
        byte[] workbookBytes,
        string fileName = "workbook.xlsx",
        OfficeProvenanceRemovalOptions? options = null) =>
        OfficeProvenancePackageMutation.Remove(workbookBytes, fileName, options, StripPackageSignatures, HasPackageSignatures, ValidatePackage);

    private static void ValidatePackage(byte[] data, OfficeProvenanceOptions options) {
        OfficeProvenanceZip.ValidateForOwningPackageMutation(data, options);
        ValidateXlsbDetectionMetadata(data, options);
        if (XlsbPackageDetector.TryFindWorkbookPart(
            data, options.MaxAssetBytes, options.MaxAssetBytes, out _)) {
            ValidateUniqueXlsbPartNames(data);
            return;
        }
        using var stream = new MemoryStream(data, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false);
        if (document.WorkbookPart == null || !IsSupportedWorkbookContentType(document.WorkbookPart.ContentType)) {
            throw new InvalidDataException("The package is not an Excel workbook.");
        }
    }

    private static void ValidateUniqueXlsbPartNames(byte[] data) {
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        var entries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (ZipArchiveEntry entry in archive.Entries) {
            if (string.IsNullOrEmpty(entry.Name)) continue;
            string normalized = NormalizePartName(entry.FullName);
            if (!entries.Add(normalized)) {
                throw new InvalidDataException($"The XLSB package contains duplicate part name '{normalized}'.");
            }
        }
    }

    private static bool IsSupportedWorkbookContentType(string contentType) => new[] {
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml",
        "application/vnd.ms-excel.sheet.macroEnabled.main+xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml",
        "application/vnd.ms-excel.template.macroEnabled.main+xml",
        "application/vnd.ms-excel.addin.macroEnabled.main+xml"
    }.Contains(contentType, StringComparer.OrdinalIgnoreCase);

    private static void ValidateXlsbDetectionMetadata(byte[] data, OfficeProvenanceOptions options) {
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        ValidateXlsbDetectionPart(archive, "_rels/.rels", options);
        ValidateXlsbDetectionPart(archive, "[Content_Types].xml", options);
    }

    private static void ValidateXlsbDetectionPart(ZipArchive archive, string entryName, OfficeProvenanceOptions options) {
        ZipArchiveEntry[] matches = archive.Entries
            .Where(entry => NormalizePartName(entry.FullName).Equals(entryName, StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (matches.Length == 0) return;
        if (matches.Length != 1) throw new InvalidDataException("The workbook package contains duplicate detection metadata parts.");
        byte[] xml = ReadBoundedEntry(matches[0], options.MaxAssetBytes);
        OfficeProvenanceXml.ValidateMaterializedNodeBudget(xml, options, "XLSB package detection metadata");
    }

    private static bool HasPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions options) {
        if (OfficeProvenanceZip.HasNativePackageSignature(data, options)) return true;
        if (XlsbPackageDetector.TryFindWorkbookPart(
            data, options.Limits.MaxAssetBytes, options.Limits.MaxAssetBytes, out _)) {
            return HasXlsbRelationshipOwnedApplicationSignatureMetadata(data, options.Limits);
        }
        using var stream = new MemoryStream(data, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false);
        ExtendedFilePropertiesPart? applicationProperties = document.ExtendedFilePropertiesPart;
        ValidateApplicationProperties(applicationProperties, options.Limits);
        return applicationProperties?.Properties?.DigitalSignature != null;
    }

    private static OfficeProvenanceSignatureStripResult StripPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions options) {
        OfficeProvenanceOptions limits = options.Limits;
        if (XlsbPackageDetector.TryFindWorkbookPart(
            data, limits.MaxAssetBytes, limits.MaxAssetBytes, out _)) return StripXlsbPackageSignatures(data, options);
        using var stream = new OfficeProvenanceBoundedMemoryStream(options.EffectiveMaxIntermediateBytes, data.Length);
        stream.Write(data, 0, data.Length);
        stream.Position = 0;
        bool hadSignatures;
        using (SpreadsheetDocument document = SpreadsheetDocument.Open(stream, true)) {
            DigitalSignatureOriginPart? origin = document.DigitalSignatureOriginPart;
            ExtendedFilePropertiesPart? applicationProperties = document.ExtendedFilePropertiesPart;
            ValidateApplicationProperties(applicationProperties, limits);
            bool hasApplicationMetadata = applicationProperties?.Properties?.DigitalSignature != null;
            hadSignatures = origin != null || hasApplicationMetadata;
            if (origin != null) document.DeletePart(origin);
            if (hasApplicationMetadata) {
                document.ExtendedFilePropertiesPart!.Properties!.DigitalSignature = null;
                document.ExtendedFilePropertiesPart.Properties.Save();
            }
        }
        return new OfficeProvenanceSignatureStripResult(stream.ToArray(), hadSignatures);
    }

    private static OfficeProvenanceSignatureStripResult StripXlsbPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions options) {
        OfficeProvenanceOptions limits = options.Limits;
        var inspectionOptions = new OfficePackageSignatureInspectionOptions {
            MaxPackageBytes = limits.MaxAssetBytes,
            MaxPackageParts = limits.MaxContainerEntries,
            MaxPartBytes = limits.MaxAssetBytes,
            MaxSignatureBytes = limits.MaxAssetBytes,
            MaxTotalDigestBytes = limits.MaxExpandedContainerBytes,
            VerifyDigests = false
        };
        OfficePackageSignatureInfo info = OfficePackageSignatureService.Inspect(data, inspectionOptions);
        if (!info.SignatureDiscoveryComplete) {
            throw new InvalidDataException("The XLSB package signature state could not be determined safely.");
        }
        bool hasNativeSignature = info.OriginRelationshipCount > 0 || info.HasDigitalSignatureOriginPart ||
            info.SignatureParts.Count > 0;
        bool hasApplicationSignatureMetadata = HasXlsbRelationshipOwnedApplicationSignatureMetadata(data, limits);
        if (!hasNativeSignature && !hasApplicationSignatureMetadata) {
            return new OfficeProvenanceSignatureStripResult((byte[])data.Clone(), hadSignatures: false);
        }

        HashSet<string> signatureEntries = DiscoverXlsbSignatureEntries(data, limits, info);
        HashSet<string> applicationMetadataEntries = DiscoverXlsbApplicationMetadataEntries(data, limits);
        OfficeProvenanceSignatureStripResult rewritten = OfficeProvenanceZip.RemoveEntries(
            data,
            entryName => signatureEntries.Contains(NormalizePartName(entryName)),
            limits.MaxExpandedContainerBytes,
            entryName => IsXlsbSignatureMetadataEntry(entryName, applicationMetadataEntries),
            (entryName, content) => RewriteXlsbSignatureMetadata(
                entryName,
                content,
                signatureEntries,
                limits),
            limits.MaxAssetBytes,
            options.EffectiveMaxOutputBytes,
            limits.CancellationToken);
        return new OfficeProvenanceSignatureStripResult(rewritten.Data, hadSignatures: true);
    }

    private static bool HasXlsbRelationshipOwnedApplicationSignatureMetadata(
        byte[] data,
        OfficeProvenanceOptions limits) {
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        foreach (string target in DiscoverXlsbApplicationMetadataEntries(data, limits)) {
            ZipArchiveEntry[] matches = archive.Entries
                .Where(entry => NormalizePartName(entry.FullName).Equals(target, StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (matches.Length != 1) {
                throw new InvalidDataException("The XLSB extended-properties relationship target is ambiguous.");
            }
            byte[] xml = ReadBoundedEntry(matches[0], limits.MaxAssetBytes);
            OfficeProvenanceXml.ValidateMaterializedNodeBudget(xml, limits, "XLSB application metadata");
            using var input = new MemoryStream(xml, writable: false);
            using XmlReader reader = XmlReader.Create(input, OfficeProvenanceXml.CreateReaderSettings(limits));
            XDocument document = XDocument.Load(reader, LoadOptions.None);
            XNamespace properties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
            if (document.Root?.Name == properties + "Properties" &&
                document.Root.Elements(properties + "DigSig").Any()) return true;
        }
        return false;
    }

    private static HashSet<string> DiscoverXlsbSignatureEntries(
        byte[] data,
        OfficeProvenanceOptions limits,
        OfficePackageSignatureInfo info) {
        var entries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var pending = new Queue<string>();
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        PackageContentTypes contentTypes = ReadPackageContentTypes(archive, limits);
        void AddPart(string? value, string expectedContentType, string role) {
            if (string.IsNullOrWhiteSpace(value)) return;
            string normalized = NormalizePartName(value!);
            ZipArchiveEntry[] matches = archive.Entries
                .Where(entry => NormalizePartName(entry.FullName).Equals(normalized, StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (matches.Length != 1 ||
                !string.Equals(contentTypes.GetContentType(normalized), expectedContentType, StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException($"The XLSB {role} target has an unexpected package content type.");
            }
            if (entries.Add(normalized)) pending.Enqueue(normalized);
        }

        AddPart(info.OriginPartUri, SignatureOriginContentType, "signature-origin");
        foreach (OfficePackageSignaturePartInfo part in info.SignatureParts) {
            AddPart(part.Uri, SignaturePartContentType, "signature-part");
        }
        foreach (XElement relationship in ReadRelationships(archive, "_rels/.rels", limits)) {
            if (!TryGetSignatureRelationshipContentType(relationship, requireOrigin: true, out string expectedContentType)) continue;
            AddPart(
                ResolveRelationshipTarget(string.Empty, (string?)relationship.Attribute("Target")),
                expectedContentType,
                "signature-origin");
        }

        int traversed = 0;
        while (pending.Count != 0) {
            if (++traversed > limits.MaxContainerEntries) {
                throw new InvalidDataException("The XLSB signature graph exceeds the configured container-entry limit.");
            }
            string sourcePart = pending.Dequeue();
            string relationshipsPart = GetRelationshipsPartName(sourcePart);
            XElement[] relationships = ReadRelationships(archive, relationshipsPart, limits);
            if (relationships.Length == 0) continue;
            entries.Add(relationshipsPart);
            foreach (XElement relationship in relationships) {
                if (!TryGetSignatureRelationshipContentType(relationship, requireOrigin: false, out string expectedContentType)) continue;
                AddPart(
                    ResolveRelationshipTarget(sourcePart, (string?)relationship.Attribute("Target")),
                    expectedContentType,
                    "signature-part");
            }
        }
        return entries;
    }

    private static PackageContentTypes ReadPackageContentTypes(ZipArchive archive, OfficeProvenanceOptions limits) {
        ZipArchiveEntry[] matches = archive.Entries
            .Where(entry => NormalizePartName(entry.FullName).Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (matches.Length != 1) throw new InvalidDataException("The XLSB package must contain exactly one content-types part.");
        byte[] xml = ReadBoundedEntry(matches[0], limits.MaxAssetBytes);
        OfficeProvenanceXml.ValidateMaterializedNodeBudget(xml, limits, "XLSB content types");
        using var input = new MemoryStream(xml, writable: false);
        using XmlReader reader = XmlReader.Create(input, OfficeProvenanceXml.CreateReaderSettings(limits));
        XDocument document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        XNamespace contentTypesNamespace = "http://schemas.openxmlformats.org/package/2006/content-types";
        XElement root = document.Root ?? throw new InvalidDataException("The XLSB content-types part has no root element.");
        if (root.Name != contentTypesNamespace + "Types") {
            throw new InvalidDataException("The XLSB content-types part has an unexpected root element.");
        }
        var overrides = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var defaults = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (XElement element in root.Elements()) {
            string? contentType = (string?)element.Attribute("ContentType");
            if (string.IsNullOrWhiteSpace(contentType)) continue;
            if (element.Name == contentTypesNamespace + "Override") {
                string partName = NormalizePartName((string?)element.Attribute("PartName") ?? string.Empty);
                if (partName.Length == 0 || overrides.ContainsKey(partName)) {
                    throw new InvalidDataException("The XLSB package contains ambiguous content-type overrides.");
                }
                overrides.Add(partName, contentType!);
            } else if (element.Name == contentTypesNamespace + "Default") {
                string extension = ((string?)element.Attribute("Extension") ?? string.Empty).TrimStart('.');
                if (extension.Length == 0 || defaults.ContainsKey(extension)) {
                    throw new InvalidDataException("The XLSB package contains ambiguous default content types.");
                }
                defaults.Add(extension, contentType!);
            }
        }
        return new PackageContentTypes(overrides, defaults);
    }

    private static XElement[] ReadRelationships(ZipArchive archive, string entryName, OfficeProvenanceOptions limits) {
        ZipArchiveEntry[] matches = archive.Entries
            .Where(entry => NormalizePartName(entry.FullName).Equals(NormalizePartName(entryName), StringComparison.OrdinalIgnoreCase))
            .ToArray();
        if (matches.Length == 0) return Array.Empty<XElement>();
        if (matches.Length != 1) throw new InvalidDataException("The XLSB package contains duplicate relationship parts.");
        byte[] xml = ReadBoundedEntry(matches[0], limits.MaxAssetBytes);
        OfficeProvenanceXml.ValidateMaterializedNodeBudget(xml, limits, "XLSB signature relationships");
        using var input = new MemoryStream(xml, writable: false);
        using XmlReader reader = XmlReader.Create(input, OfficeProvenanceXml.CreateReaderSettings(limits));
        XDocument document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        XNamespace relationshipsNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";
        XElement root = document.Root ?? throw new InvalidDataException("The XLSB relationship part has no root element.");
        if (root.Name != relationshipsNamespace + "Relationships") {
            throw new InvalidDataException("The XLSB relationship part has an unexpected root element.");
        }
        return root.Elements(relationshipsNamespace + "Relationship").ToArray();
    }

    private static bool TryGetSignatureRelationshipContentType(
        XElement relationship,
        bool requireOrigin,
        out string expectedContentType) {
        expectedContentType = string.Empty;
        if (string.Equals((string?)relationship.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) return false;
        string? type = (string?)relationship.Attribute("Type");
        if (requireOrigin) {
            if (!string.Equals(type, SignatureRelationshipPrefix + "origin", StringComparison.Ordinal)) return false;
            expectedContentType = SignatureOriginContentType;
            return true;
        }
        if (!string.Equals(type, SignatureRelationshipPrefix + "signature", StringComparison.Ordinal)) return false;
        expectedContentType = SignaturePartContentType;
        return true;
    }

    private static string? ResolveRelationshipTarget(string sourcePart, string? target) {
        if (target == null || target.Trim().Length == 0 ||
            (!target.StartsWith("/", StringComparison.Ordinal) && Uri.TryCreate(target, UriKind.Absolute, out _))) return null;
        string normalizedSource = NormalizePartName(sourcePart);
        int separator = normalizedSource.LastIndexOf('/');
        string directory = separator < 0 ? string.Empty : normalizedSource.Substring(0, separator + 1);
        var baseUri = new Uri("http://package/" + directory, UriKind.Absolute);
        Uri resolved = new Uri(baseUri, target!.Replace('\\', '/'));
        return !string.Equals(resolved.Host, "package", StringComparison.OrdinalIgnoreCase)
            ? null
            : NormalizePartName(Uri.UnescapeDataString(resolved.AbsolutePath));
    }

    private static string GetRelationshipsPartName(string sourcePart) {
        int separator = sourcePart.LastIndexOf('/');
        string directory = separator < 0 ? string.Empty : sourcePart.Substring(0, separator + 1);
        string fileName = separator < 0 ? sourcePart : sourcePart.Substring(separator + 1);
        return directory + "_rels/" + fileName + ".rels";
    }

    private static HashSet<string> DiscoverXlsbApplicationMetadataEntries(
        byte[] data,
        OfficeProvenanceOptions limits) {
        var entries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        PackageContentTypes contentTypes = ReadPackageContentTypes(archive, limits);
        foreach (XElement relationship in ReadRelationships(archive, "_rels/.rels", limits)) {
            if (!string.Equals((string?)relationship.Attribute("Type"), ExtendedPropertiesRelationshipType, StringComparison.Ordinal) ||
                string.Equals((string?)relationship.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }
            string? target = ResolveRelationshipTarget(string.Empty, (string?)relationship.Attribute("Target"));
            if (target == null) {
                throw new InvalidDataException("The XLSB extended-properties relationship has an invalid target.");
            }
            ZipArchiveEntry[] matches = archive.Entries
                .Where(entry => NormalizePartName(entry.FullName).Equals(target, StringComparison.OrdinalIgnoreCase))
                .ToArray();
            if (matches.Length != 1 ||
                !string.Equals(contentTypes.GetContentType(target), ExtendedPropertiesContentType, StringComparison.OrdinalIgnoreCase)) {
                throw new InvalidDataException("The XLSB extended-properties relationship target is missing or has an unexpected content type.");
            }
            entries.Add(target);
        }
        return entries;
    }

    private static bool IsXlsbSignatureMetadataEntry(
        string entryName,
        HashSet<string> applicationMetadataEntries) {
        string normalized = NormalizePartName(entryName);
        return normalized.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("_rels/.rels", StringComparison.OrdinalIgnoreCase) ||
            applicationMetadataEntries.Contains(normalized);
    }

    private static byte[] RewriteXlsbSignatureMetadata(
        string entryName,
        byte[] content,
        HashSet<string> signatureEntries,
        OfficeProvenanceOptions limits) {
        limits.CancellationToken.ThrowIfCancellationRequested();
        OfficeProvenanceXml.ValidateMaterializedNodeBudget(content, limits, "XLSB signature metadata");
        using var input = new MemoryStream(content, writable: false);
        using XmlReader reader = XmlReader.Create(input, OfficeProvenanceXml.CreateReaderSettings(limits));
        XDocument document = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        string normalized = NormalizePartName(entryName);
        if (normalized.Equals("_rels/.rels", StringComparison.OrdinalIgnoreCase)) {
            foreach (XElement relationship in document.Descendants()
                .Where(element => element.Name.LocalName == "Relationship" &&
                    IsDigitalSignatureOriginRelationship(element)).ToArray()) relationship.Remove();
        } else if (normalized.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase)) {
            foreach (XElement item in document.Descendants().Where(element => element.Name.LocalName == "Override").ToArray()) {
                string partName = NormalizePartName((string?)item.Attribute("PartName") ?? string.Empty);
                if (signatureEntries.Contains(partName)) item.Remove();
            }
        } else {
            XNamespace properties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
            foreach (XElement signature in document.Descendants(properties + "DigSig").ToArray()) signature.Remove();
        }
        using var output = new OfficeProvenanceBoundedMemoryStream(limits.MaxAssetBytes, content.Length);
        document.Save(output, SaveOptions.DisableFormatting);
        return output.ToArray();
    }

    private static byte[] ReadBoundedEntry(ZipArchiveEntry entry, long maximumBytes) {
        if (entry.Length > maximumBytes || entry.Length > int.MaxValue) {
            throw new InvalidDataException("XLSB signature metadata exceeds the configured asset limit.");
        }
        byte[] content = new byte[(int)entry.Length];
        using Stream input = entry.Open();
        int offset = 0;
        while (offset < content.Length) {
            int read = input.Read(content, offset, content.Length - offset);
            if (read <= 0) throw new EndOfStreamException("An XLSB signature metadata entry ended unexpectedly.");
            offset += read;
        }
        return content;
    }

    private static string NormalizePartName(string value) => value.Replace('\\', '/').TrimStart('/');

    private static bool IsDigitalSignatureOriginRelationship(XElement relationship) =>
        string.Equals(
            (string?)relationship.Attribute("Type"),
            "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin",
            StringComparison.Ordinal);

    private static void ValidateApplicationProperties(ExtendedFilePropertiesPart? part, OfficeProvenanceOptions limits) {
        if (part == null) return;
        using Stream input = part.GetStream(FileMode.Open, FileAccess.Read);
        OfficeProvenanceXml.ValidateMaterializedNodeBudget(input, limits, "Open XML application metadata");
    }

    private sealed class PackageContentTypes {
        private readonly IReadOnlyDictionary<string, string> _overrides;
        private readonly IReadOnlyDictionary<string, string> _defaults;

        internal PackageContentTypes(
            IReadOnlyDictionary<string, string> overrides,
            IReadOnlyDictionary<string, string> defaults) {
            _overrides = overrides;
            _defaults = defaults;
        }

        internal string? GetContentType(string partName) {
            string normalized = NormalizePartName(partName);
            if (_overrides.TryGetValue(normalized, out string? contentType)) return contentType;
            int separator = normalized.LastIndexOf('.');
            return separator >= 0 && separator < normalized.Length - 1 &&
                _defaults.TryGetValue(normalized.Substring(separator + 1), out contentType)
                ? contentType
                : null;
        }
    }
}
