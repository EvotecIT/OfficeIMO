using OfficeIMO;
using OfficeIMO.Core.Internal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.DocBook;

public sealed partial class DocBookDocument {
    private static readonly IReadOnlyDictionary<string, DocBookNodeKind> SharedNodeKinds =
        new Dictionary<string, DocBookNodeKind>(StringComparer.OrdinalIgnoreCase) {
            ["metadata"] = DocBookNodeKind.Info,
            ["title"] = DocBookNodeKind.Title,
            ["subtitle"] = DocBookNodeKind.Subtitle,
            ["author"] = DocBookNodeKind.Author,
            ["section"] = DocBookNodeKind.Section,
            ["paragraph"] = DocBookNodeKind.Paragraph,
            ["itemized-list"] = DocBookNodeKind.ItemizedList,
            ["ordered-list"] = DocBookNodeKind.OrderedList,
            ["variable-list"] = DocBookNodeKind.VariableList,
            ["list-item"] = DocBookNodeKind.ListItem,
            ["table"] = DocBookNodeKind.Table,
            ["table-group"] = DocBookNodeKind.TableGroup,
            ["table-head"] = DocBookNodeKind.TableHead,
            ["table-body"] = DocBookNodeKind.TableBody,
            ["table-row"] = DocBookNodeKind.Row,
            ["table-cell"] = DocBookNodeKind.Entry,
            ["code"] = DocBookNodeKind.ProgramListing,
            ["screen"] = DocBookNodeKind.Screen,
            ["link"] = DocBookNodeKind.Link,
            ["cross-reference"] = DocBookNodeKind.CrossReference,
            ["note"] = DocBookNodeKind.Note,
            ["tip"] = DocBookNodeKind.Tip,
            ["important"] = DocBookNodeKind.Important,
            ["caution"] = DocBookNodeKind.Caution,
            ["warning"] = DocBookNodeKind.Warning,
            ["figure"] = DocBookNodeKind.Figure,
            ["media"] = DocBookNodeKind.MediaObject,
            ["image-object"] = DocBookNodeKind.ImageObject,
            ["image"] = DocBookNodeKind.ImageData,
            ["caption"] = DocBookNodeKind.Caption,
            ["index"] = DocBookNodeKind.Index,
            ["index-term"] = DocBookNodeKind.IndexTerm
        };

    private static readonly ISet<string> KnownUntypedDocBookElementNames = new HashSet<string>(StringComparer.Ordinal) {
        "address", "affiliation", "appendix", "article", "authorinitials", "authorgroup", "bibliography", "chapter",
        "city", "collab", "collabname", "colophon", "colspec", "contrib", "country", "dedication", "email",
        "entrytbl", "fax", "firstname", "glossary", "honorific", "jobtitle", "lineage", "lot", "orgdiv", "orgname",
        "othername", "part", "personname", "phone", "phrase", "pob", "postcode", "preface", "primary", "reference",
        "setindex", "shortaffil", "spanspec", "state", "street", "surname", "term", "textobject", "tfoot",
        "titleabbrev", "toc", "varlistentry"
    };

    internal static bool IsKnownUntypedDocBookLocalName(string localName) =>
        KnownUntypedDocBookElementNames.Contains(localName);

    private static bool IsKnownUntypedDocBookElement(System.Xml.Linq.XName name, System.Xml.Linq.XNamespace sourceNamespace) =>
        name.Namespace == sourceNamespace && KnownUntypedDocBookElementNames.Contains(name.LocalName);

    private static bool IsDerivedBlock(OfficeDocumentModelBlock block, ILookup<string, OfficeDocumentModelNode> nodesById) =>
        !string.IsNullOrEmpty(block.Id) && block.Marker == null && block.Region == null && nodesById[block.Id].Any(node =>
            string.Equals(node.Kind, block.Kind, StringComparison.OrdinalIgnoreCase) && node.Level == block.Level &&
            (string.Equals(node.Text, block.Text, StringComparison.Ordinal) ||
             (ShouldReplaceChildrenWithPrimaryText(node) &&
              string.Equals(GetRepresentedPrimaryChildText(node), block.Text, StringComparison.Ordinal) ||
              ShouldReplaceRepresentedPrimaryChild(node) &&
              string.Equals(GetRepresentedTypedPrimaryChildText(node), block.Text, StringComparison.Ordinal))));

    private static bool IsOrphanedProjectedBlock(
        OfficeDocumentModelBlock block,
        ILookup<string, OfficeDocumentModelNode> nodesById,
        bool isDocBookProjection) =>
        isDocBookProjection && TryGetProjectedNodeId(block.Id, "docbook-", out string nodeId) && !nodesById[nodeId].Any();

    private static bool IsOrphanedProjectedTable(
        OfficeDocumentModelTable table,
        ILookup<int, OfficeDocumentModelNode> nodesByTableIndex,
        bool isDocBookProjection) =>
        isDocBookProjection && table.Location?.TableIndex is int tableIndex &&
        string.Equals(table.Location.SourceBlockKind, "table", StringComparison.Ordinal) &&
        !string.IsNullOrEmpty(table.PayloadHash) && !nodesByTableIndex[tableIndex].Any();

    private static bool IsOrphanedProjectedAsset(
        OfficeDocumentModelAsset asset,
        ILookup<string, OfficeDocumentModelNode> nodesById,
        bool isDocBookProjection) =>
        isDocBookProjection && TryGetProjectedNodeId(asset.Id, "docbook-image-", out string nodeId) && !nodesById[nodeId].Any();

    private static bool IsOrphanedProjectedLink(
        OfficeDocumentModelLink link,
        ILookup<string, OfficeDocumentModelNode> nodesById,
        bool isDocBookProjection) =>
        isDocBookProjection && TryGetProjectedNodeId(link.Id, "docbook-link-", out string nodeId) && !nodesById[nodeId].Any();

    private static bool TryGetProjectedNodeId(string? projectionId, string prefix, out string nodeId) {
        nodeId = string.Empty;
        if (projectionId == null || projectionId.Length == 0 ||
            !projectionId.StartsWith(prefix, StringComparison.Ordinal)) return false;
        string suffix = projectionId.Substring(prefix.Length);
        if (!int.TryParse(suffix, System.Globalization.NumberStyles.None,
                System.Globalization.CultureInfo.InvariantCulture, out int index) || index < 0) return false;
        nodeId = "docbook-" + suffix;
        return true;
    }

    private static bool ShouldReplaceChildrenWithPrimaryText(OfficeDocumentModelNode source) {
        if (source.Children.Count == 0 || string.Equals(source.Kind, "text", StringComparison.OrdinalIgnoreCase)) return false;
        bool acceptsDirectText = TryMapKind(source.Kind, out DocBookNodeKind kind) && NodeAcceptsDirectText(kind);
        return acceptsDirectText &&
            !string.Equals(source.Text, GetRepresentedPrimaryChildText(source), StringComparison.Ordinal);
    }

    private static bool ShouldReplaceRepresentedPrimaryChild(OfficeDocumentModelNode source) {
        if (string.IsNullOrEmpty(source.Id) || !source.Id.StartsWith("docbook-", StringComparison.Ordinal) ||
            !TryMapKind(source.Kind, out DocBookNodeKind kind) ||
            !NodeUsesTitleText(kind)) return false;
        OfficeDocumentModelNode? primaryChild = GetRepresentedTypedPrimaryChild(source, kind);
        return primaryChild != null && !string.Equals(source.Text,
            GetRepresentedSubtreeText(primaryChild), StringComparison.Ordinal);
    }

    private static OfficeDocumentModelNode? GetRepresentedTypedPrimaryChild(
        OfficeDocumentModelNode source,
        DocBookNodeKind kind) {
        string representedKind = NodeUsesTitleText(kind) ? "title" : NodeUsesParagraphText(kind) ? "paragraph" : string.Empty;
        return representedKind.Length == 0 ? null : source.Children.FirstOrDefault(child =>
            string.Equals(child.Kind, representedKind, StringComparison.OrdinalIgnoreCase));
    }

    private static string GetRepresentedTypedPrimaryChildText(OfficeDocumentModelNode source) =>
        TryMapKind(source.Kind, out DocBookNodeKind kind) && GetRepresentedTypedPrimaryChild(source, kind) is OfficeDocumentModelNode child
            ? GetRepresentedSubtreeText(child)
            : string.Empty;

    private static string GetRepresentedPrimaryChildText(OfficeDocumentModelNode source) =>
        TryMapKind(source.Kind, out DocBookNodeKind nodeKind) && nodeKind == DocBookNodeKind.Author
            ? GetRepresentedAuthorText(source)
            : string.Concat(source.Children.Where(child =>
                !string.Equals(child.Kind, "index-term", StringComparison.OrdinalIgnoreCase)).Select(GetRepresentedSubtreeText));

    private static string GetRepresentedSubtreeText(OfficeDocumentModelNode node) {
        if (string.Equals(node.Kind, "index-term", StringComparison.OrdinalIgnoreCase)) return string.Empty;
        return node.Children.Count == 0 ? node.Text : string.Concat(node.Children.Select(GetRepresentedSubtreeText));
    }

    private static string GetRepresentedAuthorText(OfficeDocumentModelNode author) {
        var nameParts = new HashSet<string>(StringComparer.Ordinal) { "honorific", "firstname", "othername", "surname", "lineage" };
        OfficeDocumentModelNode? personName = author.Children.FirstOrDefault(child =>
            IsDocBookExtensionKind(child.Kind, "personname"));
        if (personName != null) return string.Join(" ", GetDocBookAuthorTextParts(personName));
        string[] parts = author.Children.Where(child => nameParts.Any(name => IsDocBookExtensionKind(child.Kind, name)))
            .SelectMany(GetDocBookAuthorTextParts).ToArray();
        if (parts.Length > 0) return string.Join(" ", parts);
        return string.Concat(author.Children.Where(child =>
            string.Equals(child.Kind, "text", StringComparison.OrdinalIgnoreCase)).Select(child => child.Text));
    }

    private static IEnumerable<string> GetDocBookAuthorTextParts(OfficeDocumentModelNode node) {
        foreach (OfficeDocumentModelNode child in node.Children) {
            if (string.Equals(child.Kind, "text", StringComparison.OrdinalIgnoreCase)) {
                string value = child.Text.Trim();
                if (value.Length > 0) yield return value;
            } else if (child.Kind.StartsWith("extension:", StringComparison.Ordinal) && IsDocBookExtensionKind(child.Kind)) {
                foreach (string value in GetDocBookAuthorTextParts(child)) yield return value;
            }
        }
    }

    private static bool IsDocBookExtensionKind(string kind, string? localName = null) {
        const string extensionPrefix = "extension:";
        if (!kind.StartsWith(extensionPrefix, StringComparison.Ordinal)) return false;
        try {
            System.Xml.Linq.XName name = System.Xml.Linq.XName.Get(kind.Substring(extensionPrefix.Length));
            return (name.NamespaceName.Length == 0 || name.NamespaceName == DocBookSchemaProfiles.DocBook52.NamespaceUri) &&
                (localName == null || string.Equals(name.LocalName, localName, StringComparison.Ordinal));
        } catch (Exception exception) when (exception is ArgumentException || exception is System.Xml.XmlException) {
            return false;
        }
    }

    private static bool IsDerivedAsset(
        OfficeDocumentModelAsset asset,
        ILookup<string, OfficeDocumentModelNode> nodesById,
        IReadOnlyDictionary<OfficeDocumentModelNode, OfficeDocumentModelNode> parents) {
        const string prefix = "docbook-image-";
        if (string.IsNullOrEmpty(asset.Id) || !asset.Id.StartsWith(prefix, StringComparison.Ordinal) ||
            !string.Equals(asset.Kind, "image", StringComparison.OrdinalIgnoreCase)) return false;
        string nodeId = "docbook-" + asset.Id.Substring(prefix.Length);
        OfficeDocumentModelNode[] projectedImages = nodesById[nodeId].Where(node =>
            string.Equals(node.Kind, "image", StringComparison.OrdinalIgnoreCase)).Take(2).ToArray();
        bool unchangedFlatProjection = projectedImages.Length == 1 &&
            string.Equals(asset.Location?.SourceBlockKind, "image", StringComparison.Ordinal) &&
            string.Equals(asset.Location?.BlockAnchor, BuildProjectedAssetBaseline(
                asset.Kind, asset.SourceObjectId, asset.FileName, asset.Title, asset.AltText), StringComparison.Ordinal) &&
            !HasUnsupportedAssetFields(asset,
                string.IsNullOrWhiteSpace(asset.SourceObjectId) ? asset.FileName ?? string.Empty : asset.SourceObjectId!);
        if (unchangedFlatProjection) return true;
        string? reference = string.IsNullOrWhiteSpace(asset.SourceObjectId) ? asset.FileName : asset.SourceObjectId;
        OfficeDocumentModelNode? image = nodesById[nodeId].FirstOrDefault(node =>
            string.Equals(node.Kind, "image", StringComparison.OrdinalIgnoreCase) &&
            node.Attributes.TryGetValue("fileref", out string? value) && string.Equals(value, reference, StringComparison.Ordinal));
        if (image == null || !string.Equals(asset.FileName, GetReferenceFileNameFromReference(reference!), StringComparison.Ordinal)) return false;
        OfficeDocumentModelNode? media = image;
        while (parents.TryGetValue(media, out OfficeDocumentModelNode? parent)) {
            media = parent;
            if (string.Equals(media.Kind, "media", StringComparison.OrdinalIgnoreCase)) break;
        }
        if (media == null || !string.Equals(media.Kind, "media", StringComparison.OrdinalIgnoreCase)) return false;
        string? caption = media.Children.FirstOrDefault(node => string.Equals(node.Kind, "caption", StringComparison.OrdinalIgnoreCase))?.Text;
        OfficeDocumentModelNode? textObject = media.Children.FirstOrDefault(node => IsExtensionKind(node.Kind, "textobject"));
        string? alternateText = textObject == null ? null : FindExtensionDescendant(textObject, "phrase")?.Text;
        if (string.IsNullOrWhiteSpace(alternateText)) alternateText = caption;
        return string.Equals(asset.Title, caption, StringComparison.Ordinal) &&
            string.Equals(asset.AltText, alternateText, StringComparison.Ordinal) &&
            !HasUnsupportedAssetFields(asset, reference!);
    }

    private static string BuildProjectedAssetBaseline(
        string? kind,
        string? sourceObjectId,
        string? fileName,
        string? title,
        string? alternateText) {
        string payload = (kind ?? string.Empty) + "\0" + (sourceObjectId ?? string.Empty) + "\0" +
            (fileName ?? string.Empty) + "\0" + (title ?? string.Empty) + "\0" + (alternateText ?? string.Empty);
        using SHA256 hash = SHA256.Create();
        return "docbook-asset-projection-sha256:" + Convert.ToBase64String(hash.ComputeHash(Encoding.UTF8.GetBytes(payload)));
    }

    private static string? FindProjectedAssetReference(
        OfficeDocumentModelAsset asset,
        ILookup<string, OfficeDocumentModelNode> nodesById) {
        const string prefix = "docbook-image-";
        if (string.IsNullOrEmpty(asset.Id) || !asset.Id.StartsWith(prefix, StringComparison.Ordinal)) return null;
        string nodeId = "docbook-" + asset.Id.Substring(prefix.Length);
        string? reference = null;
        foreach (OfficeDocumentModelNode node in nodesById[nodeId]) {
            if (!string.Equals(node.Kind, "image", StringComparison.OrdinalIgnoreCase) ||
                !node.Attributes.TryGetValue("fileref", out string? candidate)) continue;
            if (reference != null) return null;
            reference = candidate;
        }
        return reference;
    }

    private static bool HasUnsupportedAssetFields(
        OfficeDocumentModelAsset asset,
        string reference,
        string? originalReference = null) {
        string? fileName = GetReferenceFileNameFromReference(reference);
        string? expectedExtension = GetReferenceExtensionFromFileName(fileName);
        string expectedMediaType = OfficeIMO.Drawing.OfficeImageInfo.GetMimeTypeFromExtension(expectedExtension);
        string? normalizedMediaType = expectedMediaType == "application/octet-stream" ? null : expectedMediaType;
        string? originalFileName = originalReference == null ? null : GetReferenceFileNameFromReference(originalReference);
        string? originalExtension = GetReferenceExtensionFromFileName(originalFileName);
        string originalMediaTypeValue = OfficeIMO.Drawing.OfficeImageInfo.GetMimeTypeFromExtension(originalExtension);
        string? originalMediaType = originalMediaTypeValue == "application/octet-stream" ? null : originalMediaTypeValue;
        bool extensionMatches = string.IsNullOrWhiteSpace(asset.Extension) ||
            string.Equals(asset.Extension, expectedExtension, StringComparison.OrdinalIgnoreCase) ||
            originalReference != null && string.Equals(asset.Extension, originalExtension, StringComparison.OrdinalIgnoreCase);
        bool mediaTypeMatches = string.IsNullOrWhiteSpace(asset.MediaType) ||
            string.Equals(asset.MediaType, normalizedMediaType, StringComparison.OrdinalIgnoreCase) ||
            originalReference != null && string.Equals(asset.MediaType, originalMediaType, StringComparison.OrdinalIgnoreCase);
        return !extensionMatches || !mediaTypeMatches ||
            asset.Width.HasValue || asset.Height.HasValue || asset.LengthBytes.HasValue ||
            !string.IsNullOrEmpty(asset.PayloadHash) || asset.PayloadBytes != null || asset.Region != null;
    }

    private static IReadOnlyDictionary<OfficeDocumentModelNode, OfficeDocumentModelNode> BuildStructureParentMap(
        IEnumerable<OfficeDocumentModelNode> nodes) {
        var parents = new Dictionary<OfficeDocumentModelNode, OfficeDocumentModelNode>();
        foreach (OfficeDocumentModelNode parent in nodes) {
            foreach (OfficeDocumentModelNode child in parent.Children) {
                if (!parents.ContainsKey(child)) parents.Add(child, parent);
            }
        }
        return parents;
    }

    private static string? GetReferenceFileNameFromReference(string reference) {
        int delimiter = reference.IndexOfAny(new[] { '?', '#' });
        string clean = delimiter < 0 ? reference : reference.Substring(0, delimiter);
        int separator = Math.Max(clean.LastIndexOf('/'), clean.LastIndexOf('\\'));
        string fileName = separator < 0 ? clean : clean.Substring(separator + 1);
        return string.IsNullOrWhiteSpace(fileName) ? null : fileName;
    }

    private static string? GetReferenceExtensionFromFileName(string? fileName) {
        if (string.IsNullOrWhiteSpace(fileName)) return null;
        int dot = fileName!.LastIndexOf('.');
        return dot < 0 || dot == fileName.Length - 1 ? null : fileName.Substring(dot);
    }

    private static bool IsExtensionKind(string kind, string localName) {
        const string prefix = "extension:";
        if (!kind.StartsWith(prefix, StringComparison.Ordinal)) return false;
        try {
            return string.Equals(System.Xml.Linq.XName.Get(kind.Substring(prefix.Length)).LocalName, localName, StringComparison.Ordinal);
        } catch (Exception exception) when (exception is ArgumentException || exception is System.Xml.XmlException) {
            return false;
        }
    }

    private static OfficeDocumentModelNode? FindExtensionDescendant(OfficeDocumentModelNode node, string localName) {
        foreach (OfficeDocumentModelNode child in node.Children) {
            if (IsExtensionKind(child.Kind, localName)) return child;
            OfficeDocumentModelNode? descendant = FindExtensionDescendant(child, localName);
            if (descendant != null) return descendant;
        }
        return null;
    }

    private static bool TryGetDerivedLinkNode(
        OfficeDocumentModelLink link,
        ILookup<string, OfficeDocumentModelNode> nodesById,
        out OfficeDocumentModelNode? derivedNode) {
        derivedNode = null;
        const string prefix = "docbook-link-";
        if (string.IsNullOrEmpty(link.Id) || !link.Id.StartsWith(prefix, StringComparison.Ordinal) ||
            link.Region != null ||
            link.DestinationPageNumber.HasValue || !string.IsNullOrWhiteSpace(link.DestinationMode) ||
            !string.IsNullOrWhiteSpace(link.NamedAction) || !string.IsNullOrWhiteSpace(link.RemoteFile) ||
            !string.IsNullOrWhiteSpace(link.RemoteDestinationName) || link.RemoteDestinationPageNumber.HasValue) return false;
        string nodeId = "docbook-" + link.Id.Substring(prefix.Length);
        derivedNode = nodesById[nodeId].FirstOrDefault(node =>
            (string.Equals(node.Kind, "link", StringComparison.OrdinalIgnoreCase) ||
             string.Equals(node.Kind, "cross-reference", StringComparison.OrdinalIgnoreCase)) &&
            string.Equals(node.Kind, link.Kind, StringComparison.OrdinalIgnoreCase) &&
            (link.Text == null || string.Equals(node.Text, link.Text, StringComparison.Ordinal) ||
             ShouldReplaceChildrenWithPrimaryText(node) &&
             string.Equals(GetRepresentedPrimaryChildText(node), link.Text, StringComparison.Ordinal)));
        return derivedNode != null;
    }

    private static bool LinkTargetMatches(OfficeDocumentModelNode node, OfficeDocumentModelLink link) {
        node.Attributes.TryGetValue("url", out string? url);
        if (url == null) node.Attributes.TryGetValue("{http://www.w3.org/1999/xlink}href", out url);
        node.Attributes.TryGetValue("linkend", out string? destination);
        return string.Equals(url, link.Uri, StringComparison.Ordinal) &&
            string.Equals(destination, link.DestinationName, StringComparison.Ordinal);
    }

    private static bool IsDerivedTable(
        OfficeDocumentModelTable table,
        ILookup<int, OfficeDocumentModelNode> nodesByTableIndex,
        ISet<OfficeDocumentModelNode> consumedNodes) {
        if (table.Location == null || !table.Location.TableIndex.HasValue || string.IsNullOrEmpty(table.PayloadHash) ||
            !string.Equals(table.PayloadHash, ComputeTablePayloadHash(table), StringComparison.OrdinalIgnoreCase)) return false;
        OfficeDocumentModelLocation tableLocation = table.Location;
        string expectedKind = string.Equals(table.Kind, "informaltable", StringComparison.OrdinalIgnoreCase)
            ? "informal-table" : "table";
        foreach (OfficeDocumentModelNode node in nodesByTableIndex[tableLocation.TableIndex.Value]) {
            if (consumedNodes.Contains(node) || !string.Equals(node.Kind, expectedKind, StringComparison.OrdinalIgnoreCase) ||
                !(string.Equals(node.Text, table.Title ?? string.Empty, StringComparison.Ordinal) ||
                  ShouldReplaceRepresentedPrimaryChild(node) &&
                  string.Equals(GetRepresentedTypedPrimaryChildText(node), table.Title ?? string.Empty, StringComparison.Ordinal)) ||
                node.Location?.TableIndex != tableLocation.TableIndex) continue;
            string? nodePath = node.Location?.HeadingPath;
            string expectedPath = OfficeDocumentHeadingPath.Append(nodePath, table.Title, " / ");
            if (!string.Equals(nodePath, tableLocation.HeadingPath, StringComparison.Ordinal) &&
                !string.Equals(expectedPath, tableLocation.HeadingPath, StringComparison.Ordinal)) continue;
            consumedNodes.Add(node);
            return true;
        }
        return false;
    }

    private static string ComputeTablePayloadHash(OfficeDocumentModelTable table) {
        var value = new StringBuilder();
        AppendValue(table.Summary);
        AppendList(table.Columns);
        foreach (IReadOnlyList<string> row in table.Rows) AppendList(row);
        AppendValue(table.TotalRowCount.ToString(System.Globalization.CultureInfo.InvariantCulture));
        AppendValue(table.Truncated ? "1" : "0");
        using SHA256 algorithm = SHA256.Create();
        return BitConverter.ToString(algorithm.ComputeHash(Encoding.UTF8.GetBytes(value.ToString()))).Replace("-", string.Empty);

        void AppendList(IReadOnlyList<string> items) {
            AppendValue(items.Count.ToString(System.Globalization.CultureInfo.InvariantCulture));
            foreach (string item in items) AppendValue(item);
        }

        void AppendValue(string? item) {
            if (item == null) {
                value.Append("-1:");
                return;
            }
            value.Append(item.Length.ToString(System.Globalization.CultureInfo.InvariantCulture));
            value.Append(':');
            value.Append(item);
        }
    }
}
