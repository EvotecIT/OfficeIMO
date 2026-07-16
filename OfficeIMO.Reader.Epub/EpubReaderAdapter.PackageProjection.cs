using OfficeIMO.Epub;
using OfficeIMO.Reader;
using System.Linq;

namespace OfficeIMO.Reader.Epub;

internal static partial class EpubReaderAdapter {
    private static IReadOnlyList<OfficeDocumentMetadataEntry> BuildEpubMetadata(
        EpubDocument document,
        string sourcePath,
        int blockCount,
        int tableCount,
        int linkCount,
        int assetCount) {
        var metadata = new List<OfficeDocumentMetadataEntry> {
            EpubMetadata("epub-chapter-count", "ChapterCount", document.Chapters.Count, "count"),
            EpubMetadata("epub-resource-count", "ResourceCount", document.Resources.Count, "count"),
            EpubMetadata("epub-remote-resource-count", "RemoteResourceCount", document.Resources.Count(static item => item.IsRemote), "count"),
            EpubMetadata("epub-rootfile-count", "RootfileCount", document.Rootfiles.Count, "count"),
            EpubMetadata("epub-metadata-count", "MetadataCount", document.Metadata.Count, "count"),
            EpubMetadata("epub-toc-item-count", "TableOfContentsItemCount", CountNavigationItems(document.TableOfContents), "count"),
            EpubMetadata("epub-page-list-item-count", "PageListItemCount", CountNavigationItems(document.PageList), "count"),
            EpubMetadata("epub-landmark-item-count", "LandmarkItemCount", CountNavigationItems(document.Landmarks), "count"),
            EpubMetadata("epub-block-count", "BlockCount", blockCount, "count"),
            EpubMetadata("epub-table-count", "TableCount", tableCount, "count"),
            EpubMetadata("epub-link-count", "LinkCount", linkCount, "count"),
            EpubMetadata("epub-asset-count", "AssetCount", assetCount, "count")
        };
        if (!string.IsNullOrWhiteSpace(document.Identifier)) metadata.Add(EpubMetadata("epub-identifier", "Identifier", document.Identifier!, "string"));
        if (!string.IsNullOrWhiteSpace(document.UniqueIdentifierId)) metadata.Add(EpubMetadata("epub-unique-identifier-id", "UniqueIdentifierId", document.UniqueIdentifierId!, "string"));
        if (!string.IsNullOrWhiteSpace(document.Language)) metadata.Add(EpubMetadata("epub-language", "Language", document.Language!, "string"));
        if (!string.IsNullOrWhiteSpace(document.OpfPath)) metadata.Add(EpubMetadata("epub-package-path", "PackagePath", document.OpfPath!, "string"));
        if (!string.IsNullOrWhiteSpace(document.PackageVersion)) metadata.Add(EpubMetadata("epub-package-version", "PackageVersion", document.PackageVersion!, "string"));
        if (document.RenditionLayout.HasValue) metadata.Add(EpubMetadata("epub-rendition-layout", "RenditionLayout", document.RenditionLayout.Value.ToString(), "string"));
        metadata.Add(EpubMetadata("epub-fixed-layout", "IsFixedLayout", document.IsFixedLayout, "boolean"));
        metadata.Add(EpubMetadata("epub-encryption-count", "EncryptionCount", document.Encryption.Count, "count"));
        metadata.Add(EpubMetadata("epub-requires-decryption", "RequiresDecryption", document.RequiresDecryption, "boolean"));
        metadata.Add(EpubMetadata("epub-diagnostic-count", "DiagnosticCount", document.Diagnostics.Count, "count"));

        AddRootfileMetadata(document, sourcePath, metadata);
        AddPackageMetadata(document, sourcePath, metadata);
        AddNavigationMetadata(document.TableOfContents, "toc", "TableOfContents", sourcePath, metadata);
        AddNavigationMetadata(document.PageList, "page-list", "PageList", sourcePath, metadata);
        AddNavigationMetadata(document.Landmarks, "landmarks", "Landmarks", sourcePath, metadata);
        return metadata;
    }

    private static void AddRootfileMetadata(EpubDocument document, string sourcePath, List<OfficeDocumentMetadataEntry> destination) {
        for (int index = 0; index < document.Rootfiles.Count; index++) {
            EpubRootfile rootfile = document.Rootfiles[index];
            var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                ["isAvailable"] = rootfile.IsAvailable ? "true" : "false",
                ["isSelected"] = rootfile.IsSelected ? "true" : "false"
            };
            AddAttribute(attributes, "mediaType", rootfile.MediaType);
            destination.Add(new OfficeDocumentMetadataEntry {
                Id = "epub-rootfile-" + (index + 1).ToString("D4", CultureInfo.InvariantCulture),
                Category = "epub.container.rootfile",
                Name = "Rootfile",
                Value = rootfile.FullPath,
                ValueType = "object",
                SourceObjectId = rootfile.FullPath,
                Location = new ReaderLocation { Path = BuildEpubLocationPath(sourcePath, rootfile.FullPath) },
                Attributes = attributes
            });
        }
    }

    private static void AddPackageMetadata(EpubDocument document, string sourcePath, List<OfficeDocumentMetadataEntry> destination) {
        for (int index = 0; index < document.Metadata.Count; index++) {
            EpubMetadataEntry item = document.Metadata[index];
            var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                ["kind"] = item.Kind.ToString(),
                ["elementName"] = item.Name,
                ["namespaceUri"] = item.NamespaceUri
            };
            AddAttribute(attributes, "id", item.Id);
            AddAttribute(attributes, "property", item.Property);
            AddAttribute(attributes, "refines", item.Refines);
            AddAttribute(attributes, "scheme", item.Scheme);
            AddAttribute(attributes, "language", item.Language);
            AddAttribute(attributes, "legacyName", item.LegacyName);
            AddAttribute(attributes, "role", item.Role);
            AddAttribute(attributes, "fileAs", item.FileAs);
            AddAttribute(attributes, "event", item.Event);
            AddAttribute(attributes, "href", item.Href);
            AddAttribute(attributes, "rel", item.Rel);
            AddAttribute(attributes, "mediaType", item.MediaType);
            destination.Add(new OfficeDocumentMetadataEntry {
                Id = "epub-package-metadata-" + (index + 1).ToString("D4", CultureInfo.InvariantCulture),
                Category = "epub.package.metadata",
                Name = item.Property ?? item.LegacyName ?? item.Name,
                Value = item.Value,
                ValueType = item.Kind == EpubMetadataKind.Link ? "uri" : "string",
                SourceObjectId = item.Id ?? item.Refines,
                Location = string.IsNullOrWhiteSpace(document.OpfPath)
                    ? null
                    : new ReaderLocation { Path = BuildEpubLocationPath(sourcePath, document.OpfPath!) },
                Attributes = attributes
            });
        }
    }

    private static void AddNavigationMetadata(
        IReadOnlyList<EpubNavigationItem> items,
        string categorySuffix,
        string name,
        string sourcePath,
        List<OfficeDocumentMetadataEntry> destination) {
        int index = 0;
        AddNavigationMetadata(items, categorySuffix, name, sourcePath, destination, 1, ref index);
    }

    private static void AddNavigationMetadata(
        IReadOnlyList<EpubNavigationItem> items,
        string categorySuffix,
        string name,
        string sourcePath,
        List<OfficeDocumentMetadataEntry> destination,
        int depth,
        ref int index) {
        foreach (EpubNavigationItem item in items) {
            index++;
            var attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                ["depth"] = depth.ToString(CultureInfo.InvariantCulture),
                ["isRemote"] = item.IsRemote ? "true" : "false",
                ["source"] = item.Source.ToString()
            };
            AddAttribute(attributes, "href", item.Href);
            AddAttribute(attributes, "target", item.Target);
            AddAttribute(attributes, "fragment", item.Fragment);
            AddAttribute(attributes, "semanticType", item.SemanticType);
            if (item.PlayOrder.HasValue) attributes["playOrder"] = item.PlayOrder.Value.ToString(CultureInfo.InvariantCulture);

            destination.Add(new OfficeDocumentMetadataEntry {
                Id = "epub-navigation-" + categorySuffix + "-" + index.ToString("D4", CultureInfo.InvariantCulture),
                Category = "epub.navigation." + categorySuffix,
                Name = name,
                Value = item.Label,
                ValueType = "object",
                SourceObjectId = item.Target,
                Location = string.IsNullOrWhiteSpace(item.Target)
                    ? null
                    : new ReaderLocation {
                        Path = BuildEpubLocationPath(sourcePath, item.Target!),
                        BlockAnchor = item.Fragment
                    },
                Attributes = attributes
            });
            AddNavigationMetadata(item.Children, categorySuffix, name, sourcePath, destination, depth + 1, ref index);
        }
    }

    private static int CountNavigationItems(IEnumerable<EpubNavigationItem> items) {
        int count = 0;
        foreach (EpubNavigationItem item in items) {
            count++;
            count += CountNavigationItems(item.Children);
        }
        return count;
    }

    private static void AddAttribute(Dictionary<string, string> destination, string name, string? value) {
        if (!string.IsNullOrWhiteSpace(value)) destination[name] = value!;
    }

    private static OfficeDocumentMetadataEntry EpubMetadata(string id, string name, object value, string valueType) => new OfficeDocumentMetadataEntry {
        Id = id,
        Category = "epub.package",
        Name = name,
        Value = Convert.ToString(value, CultureInfo.InvariantCulture),
        ValueType = valueType
    };

    private static string BuildEpubLocationPath(string sourcePath, string path) {
        if (path.StartsWith("//", StringComparison.Ordinal)) return path;
        if (Uri.TryCreate(path, UriKind.Absolute, out Uri? uri) && !uri.IsFile) return path;
        return BuildVirtualPath(sourcePath, path);
    }
}
