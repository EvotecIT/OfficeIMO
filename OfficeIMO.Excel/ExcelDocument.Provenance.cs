using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.Xlsb.Package;
using OfficeIMO.Provenance;
using OfficeIMO.Security;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
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
        if (XlsbPackageDetector.TryFindWorkbookPart(data, out _)) return;
        using var stream = new MemoryStream(data, writable: false);
        using SpreadsheetDocument document = SpreadsheetDocument.Open(stream, false);
        if (document.WorkbookPart == null || document.WorkbookPart.ContentType is not (
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" or
            "application/vnd.ms-excel.sheet.macroEnabled.main+xml" or
            "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml" or
            "application/vnd.ms-excel.template.macroEnabled.main+xml" or
            "application/vnd.ms-excel.addin.macroEnabled.main+xml")) {
            throw new InvalidDataException("The package is not an Excel workbook.");
        }
    }

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

    private static bool HasPackageSignatures(byte[] data, OfficeProvenanceRemovalOptions options) =>
        OfficeProvenanceZip.HasPackageSignature(data, options) ||
        OfficeProvenanceZip.HasApplicationSignatureMetadata(data, options.Limits);

    private static OfficeProvenanceSignatureStripResult StripPackageSignatures(byte[] data, OfficeProvenanceOptions limits) {
        if (XlsbPackageDetector.TryFindWorkbookPart(data, out _)) return StripXlsbPackageSignatures(data, limits);
        using var stream = new MemoryStream(data.Length);
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

    private static OfficeProvenanceSignatureStripResult StripXlsbPackageSignatures(byte[] data, OfficeProvenanceOptions limits) {
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
        if (!info.HasSignatures) return new OfficeProvenanceSignatureStripResult((byte[])data.Clone(), hadSignatures: false);

        HashSet<string> signatureEntries = DiscoverXlsbSignatureEntries(data, limits, info);
        OfficeProvenanceSignatureStripResult rewritten = OfficeProvenanceZip.RemoveEntries(
            data,
            entryName => signatureEntries.Contains(NormalizePartName(entryName)),
            entryName => IsXlsbSignatureMetadataEntry(entryName),
            (entryName, content) => RewriteXlsbSignatureMetadata(entryName, content, signatureEntries, limits),
            limits.MaxAssetBytes);
        return new OfficeProvenanceSignatureStripResult(rewritten.Data, hadSignatures: true);
    }

    private static HashSet<string> DiscoverXlsbSignatureEntries(
        byte[] data,
        OfficeProvenanceOptions limits,
        OfficePackageSignatureInfo info) {
        var entries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var pending = new Queue<string>();
        void AddPart(string? value) {
            if (string.IsNullOrWhiteSpace(value)) return;
            string normalized = NormalizePartName(value!);
            if (entries.Add(normalized)) pending.Enqueue(normalized);
        }

        AddPart(info.OriginPartUri);
        foreach (OfficePackageSignaturePartInfo part in info.SignatureParts) AddPart(part.Uri);
        using var stream = new MemoryStream(data, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: false);
        foreach (XElement relationship in ReadRelationships(archive, "_rels/.rels", limits)) {
            if (!IsInternalDigitalSignatureRelationship(relationship, requireOrigin: true)) continue;
            AddPart(ResolveRelationshipTarget(string.Empty, (string?)relationship.Attribute("Target")));
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
                if (!IsInternalDigitalSignatureRelationship(relationship, requireOrigin: false)) continue;
                AddPart(ResolveRelationshipTarget(sourcePart, (string?)relationship.Attribute("Target")));
            }
        }
        return entries;
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
        return document.Descendants().Where(element => element.Name.LocalName == "Relationship").ToArray();
    }

    private static bool IsInternalDigitalSignatureRelationship(XElement relationship, bool requireOrigin) {
        if (string.Equals((string?)relationship.Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) return false;
        string? type = (string?)relationship.Attribute("Type");
        const string prefix = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/";
        return requireOrigin
            ? string.Equals(type, prefix + "origin", StringComparison.Ordinal)
            : type?.StartsWith(prefix, StringComparison.Ordinal) == true;
    }

    private static string? ResolveRelationshipTarget(string sourcePart, string? target) {
        if (string.IsNullOrWhiteSpace(target) || Uri.TryCreate(target, UriKind.Absolute, out _)) return null;
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

    private static bool IsXlsbSignatureMetadataEntry(string entryName) {
        string normalized = NormalizePartName(entryName);
        return normalized.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("_rels/.rels", StringComparison.OrdinalIgnoreCase) ||
            normalized.Equals("docProps/app.xml", StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] RewriteXlsbSignatureMetadata(
        string entryName,
        byte[] content,
        HashSet<string> signatureEntries,
        OfficeProvenanceOptions limits) {
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
                string? contentType = (string?)item.Attribute("ContentType");
                if (signatureEntries.Contains(partName) || contentType is
                    "application/vnd.openxmlformats-package.digital-signature-origin" or
                    "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml") item.Remove();
            }
        } else {
            XNamespace properties = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
            foreach (XElement signature in document.Descendants(properties + "DigSig").ToArray()) signature.Remove();
        }
        using var output = new MemoryStream();
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
}
