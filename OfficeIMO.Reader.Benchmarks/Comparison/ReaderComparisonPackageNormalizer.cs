using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Reader.Benchmarks.Comparison;

/// <summary>Canonicalizes generated Open XML packages so comparison inputs are byte-for-byte reproducible.</summary>
internal static class ReaderComparisonPackageNormalizer {
    private const string CorePropertiesContentType =
        "application/vnd.openxmlformats-package.core-properties+xml";
    private const string CorePropertiesRelationshipType =
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    private const string CorePropertiesPath = "docProps/core.xml";
    private static readonly XNamespace ContentTypesNamespace =
        "http://schemas.openxmlformats.org/package/2006/content-types";
    private static readonly XNamespace CorePropertiesNamespace =
        "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    private static readonly XNamespace DublinCoreTermsNamespace = "http://purl.org/dc/terms/";
    private static readonly XNamespace RelationshipsNamespace =
        "http://schemas.openxmlformats.org/package/2006/relationships";
    private static readonly XNamespace XmlSchemaInstanceNamespace =
        "http://www.w3.org/2001/XMLSchema-instance";

    internal static byte[] Normalize(byte[] packageBytes, DateTimeOffset timestamp) {
        var entries = ReadEntries(packageBytes);
        string? originalCorePath = entries.Keys.SingleOrDefault(IsCorePropertiesPath);
        if (originalCorePath != null && !string.Equals(originalCorePath, CorePropertiesPath, StringComparison.Ordinal)) {
            entries[CorePropertiesPath] = entries[originalCorePath];
            entries.Remove(originalCorePath);
        }

        if (entries.TryGetValue(CorePropertiesPath, out byte[]? coreProperties)) {
            entries[CorePropertiesPath] = NormalizeCoreProperties(coreProperties, timestamp);
        }
        if (entries.TryGetValue("[Content_Types].xml", out byte[]? contentTypes)) {
            entries["[Content_Types].xml"] = NormalizeContentTypes(contentTypes);
        }

        foreach (string relationshipsPath in entries.Keys
                     .Where(path => path.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
                     .OrderBy(path => path, StringComparer.Ordinal)
                     .ToArray()) {
            NormalizeRelationships(entries, relationshipsPath);
        }

        return WriteEntries(entries, timestamp);
    }

    private static Dictionary<string, byte[]> ReadEntries(byte[] packageBytes) {
        var entries = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        using var stream = new MemoryStream(packageBytes, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        foreach (ZipArchiveEntry entry in archive.Entries) {
            if (entry.FullName.EndsWith("/", StringComparison.Ordinal)) continue;
            using Stream input = entry.Open();
            using var output = new MemoryStream();
            input.CopyTo(output);
            entries.Add(entry.FullName, output.ToArray());
        }
        return entries;
    }

    private static byte[] WriteEntries(IReadOnlyDictionary<string, byte[]> entries, DateTimeOffset timestamp) {
        using var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (KeyValuePair<string, byte[]> item in entries.OrderBy(item => item.Key, StringComparer.Ordinal)) {
                ZipArchiveEntry entry = archive.CreateEntry(item.Key, CompressionLevel.Optimal);
                entry.LastWriteTime = timestamp;
                using Stream output = entry.Open();
                output.Write(item.Value, 0, item.Value.Length);
            }
        }
        return stream.ToArray();
    }

    private static byte[] NormalizeCoreProperties(byte[] content, DateTimeOffset timestamp) {
        XDocument document = LoadXml(content);
        XElement root = document.Root ?? throw new InvalidDataException("The package core-properties part has no root element.");
        string value = timestamp.UtcDateTime.ToString("O", System.Globalization.CultureInfo.InvariantCulture);
        SetCoreTimestamp(root, "created", value);
        SetCoreTimestamp(root, "modified", value);
        return SaveXml(document);
    }

    private static void SetCoreTimestamp(XElement root, string localName, string value) {
        XElement? element = root.Element(DublinCoreTermsNamespace + localName);
        if (element == null) {
            element = new XElement(DublinCoreTermsNamespace + localName);
            root.Add(element);
        }
        element.SetAttributeValue(XmlSchemaInstanceNamespace + "type", "dcterms:W3CDTF");
        element.Value = value;
    }

    private static byte[] NormalizeContentTypes(byte[] content) {
        XDocument document = LoadXml(content);
        XElement root = document.Root ?? throw new InvalidDataException("The package content-types part has no root element.");
        root.Elements(ContentTypesNamespace + "Default")
            .Where(element => string.Equals((string?)element.Attribute("Extension"), "psmdcp", StringComparison.OrdinalIgnoreCase))
            .Remove();
        XElement? coreOverride = root.Elements(ContentTypesNamespace + "Override")
            .FirstOrDefault(element => string.Equals(
                (string?)element.Attribute("ContentType"),
                CorePropertiesContentType,
                StringComparison.Ordinal));
        if (coreOverride == null) {
            coreOverride = new XElement(ContentTypesNamespace + "Override");
            root.Add(coreOverride);
        }
        coreOverride.SetAttributeValue("PartName", "/" + CorePropertiesPath);
        coreOverride.SetAttributeValue("ContentType", CorePropertiesContentType);
        return SaveXml(document);
    }

    private static void NormalizeRelationships(IDictionary<string, byte[]> entries, string relationshipsPath) {
        XDocument document = LoadXml(entries[relationshipsPath]);
        XElement root = document.Root ?? throw new InvalidDataException("A package relationships part has no root element.");
        XElement[] relationships = root.Elements(RelationshipsNamespace + "Relationship").ToArray();
        foreach (XElement relationship in relationships.Where(IsCorePropertiesRelationship)) {
            relationship.SetAttributeValue("Target", CorePropertiesPath);
        }

        XElement[] ordered = relationships
            .OrderBy(element => (string?)element.Attribute("Type"), StringComparer.Ordinal)
            .ThenBy(element => (string?)element.Attribute("Target"), StringComparer.Ordinal)
            .ThenBy(element => (string?)element.Attribute("TargetMode"), StringComparer.Ordinal)
            .ToArray();
        var identifiers = new Dictionary<string, string>(StringComparer.Ordinal);
        for (int index = 0; index < ordered.Length; index++) {
            string? oldIdentifier = (string?)ordered[index].Attribute("Id");
            string newIdentifier = "rId" + (index + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
            if (!string.IsNullOrEmpty(oldIdentifier)) identifiers[oldIdentifier] = newIdentifier;
            ordered[index].SetAttributeValue("Id", newIdentifier);
        }
        root.ReplaceNodes(ordered);
        entries[relationshipsPath] = SaveXml(document);

        string? ownerPath = ResolveRelationshipsOwner(relationshipsPath);
        if (ownerPath == null || !entries.TryGetValue(ownerPath, out byte[]? ownerContent)) return;
        XDocument owner = LoadXml(ownerContent);
        XElement ownerRoot = owner.Root ?? throw new InvalidDataException("A relationship owner part has no root element.");
        foreach (XAttribute attribute in ownerRoot.DescendantsAndSelf().Attributes()) {
            if (identifiers.TryGetValue(attribute.Value, out string? replacement)) attribute.Value = replacement;
        }
        ownerRoot.DescendantsAndSelf()
            .Attributes()
            .Where(attribute =>
                string.Equals(
                    attribute.Name.NamespaceName,
                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                    StringComparison.Ordinal) &&
                attribute.Name.LocalName.StartsWith("rsid", StringComparison.OrdinalIgnoreCase))
            .Remove();
        entries[ownerPath] = SaveXml(owner);
    }

    private static bool IsCorePropertiesRelationship(XElement relationship) =>
        string.Equals(
            (string?)relationship.Attribute("Type"),
            CorePropertiesRelationshipType,
            StringComparison.Ordinal);

    private static bool IsCorePropertiesPath(string path) =>
        string.Equals(path, CorePropertiesPath, StringComparison.OrdinalIgnoreCase) ||
        path.IndexOf("/metadata/core-properties/", StringComparison.OrdinalIgnoreCase) >= 0;

    private static string? ResolveRelationshipsOwner(string relationshipsPath) {
        if (string.Equals(relationshipsPath, "_rels/.rels", StringComparison.OrdinalIgnoreCase)) return null;
        int marker = relationshipsPath.LastIndexOf("/_rels/", StringComparison.Ordinal);
        if (marker < 0 || !relationshipsPath.EndsWith(".rels", StringComparison.OrdinalIgnoreCase)) return null;
        string directory = relationshipsPath.Substring(0, marker);
        string fileName = relationshipsPath.Substring(marker + "/_rels/".Length);
        return directory + "/" + fileName.Substring(0, fileName.Length - ".rels".Length);
    }

    private static XDocument LoadXml(byte[] content) {
        using var stream = new MemoryStream(content, writable: false);
        return XDocument.Load(stream, LoadOptions.None);
    }

    private static byte[] SaveXml(XDocument document) {
        using var stream = new MemoryStream();
        using (XmlWriter writer = XmlWriter.Create(stream, new XmlWriterSettings {
                   Encoding = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false),
                   Indent = false,
                   OmitXmlDeclaration = false
               })) {
            document.Save(writer);
        }
        return stream.ToArray();
    }
}
