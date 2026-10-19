using System.IO.Compression;
using System.IO.Packaging;
using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Excel.Utilities;

internal static partial class ExcelPackageUtilities {
    /// <summary>Resolves the actual workbook carrier instead of assuming the SDK retained its original part name.</summary>
    internal static string? GetWorkbookPartName(ZipArchive archive) => ReadRootParts(archive).Workbook;

    private static (string? Workbook, string? App, string? Core) ReadRootParts(ZipArchive archive, long maxCharactersInPart = 0) {
        ZipArchiveEntry? entry = archive.GetEntry("_rels/.rels");
        if (entry == null) return (null, null, null);
        if (entry.Length > MaximumContentTypesEntryBytes) throw new InvalidDataException("Package root relationships exceed the supported XML limit.");
        using Stream input = entry.Open();
        using var reader = XmlReader.Create(input, new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
            MaxCharactersInDocument = maxCharactersInPart > 0 ? Math.Min(maxCharactersInPart, MaximumContentTypesEntryBytes) : MaximumContentTypesEntryBytes
        });
        XElement? root = XDocument.Load(reader).Root;
        return (FindRootPart(archive, root, "/officeDocument"), FindRootPart(archive, root, "/extended-properties"),
            FindRootPart(archive, root, "/metadata/core-properties"));
    }

    private static string? FindRootPart(ZipArchive archive, XElement? root, string relationshipSuffix) {
        XNamespace ns = "http://schemas.openxmlformats.org/package/2006/relationships";
        XElement[] matches = root?.Elements(ns + "Relationship")
            .Where(item => IsOfficeRootRelationship((string?)item.Attribute("Type"), relationshipSuffix)).ToArray()
            ?? Array.Empty<XElement>();
        if (matches.Length == 0) return null;
        if (matches.Length != 1 || string.Equals((string?)matches[0].Attribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("Package root relationship must name one internal part.");
        string? target = (string?)matches[0].Attribute("Target");
        if (string.IsNullOrEmpty(target) || !Uri.TryCreate(target, UriKind.Relative, out Uri? targetUri))
            throw new InvalidDataException("Package root relationship has an invalid target.");
        Uri partUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), targetUri);
        string partName = partUri.OriginalString;
        if (archive.GetEntry(partName.TrimStart('/')) == null) throw new InvalidDataException("Package root relationship targets a missing part.");
        return partName;
    }

    private static bool IsOfficeRootRelationship(string? type, string suffix) {
        if (suffix == "/metadata/core-properties")
            return string.Equals(type, "http://schemas.openxmlformats.org/package/2006/relationships" + suffix, StringComparison.Ordinal);
        return string.Equals(type, "http://schemas.openxmlformats.org/officeDocument/2006/relationships" + suffix, StringComparison.Ordinal)
            || string.Equals(type, "http://purl.oclc.org/ooxml/officeDocument/relationships" + suffix, StringComparison.Ordinal);
    }

    private static string? GetEffectiveContentType(XElement root, string partName) {
        XNamespace ns = root.Name.Namespace;
        return (string?)root.Elements(ns + "Override").FirstOrDefault(item =>
                string.Equals((string?)item.Attribute("PartName"), partName, StringComparison.OrdinalIgnoreCase))?.Attribute("ContentType")
            ?? (string?)root.Elements(ns + "Default").FirstOrDefault(item =>
                string.Equals((string?)item.Attribute("Extension"), Path.GetExtension(partName).TrimStart('.'), StringComparison.OrdinalIgnoreCase))?.Attribute("ContentType");
    }
}
