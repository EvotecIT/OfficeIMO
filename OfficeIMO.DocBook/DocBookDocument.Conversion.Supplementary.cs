using OfficeIMO;
using OfficeIMO.Core.Internal;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.DocBook;

public sealed partial class DocBookDocument {
    private static readonly ISet<string> KnownUntypedDocBookElementNames = new HashSet<string>(StringComparer.Ordinal) {
        "appendix", "article", "authorgroup", "bibliography", "chapter", "colophon", "colspec", "dedication",
        "entrytbl", "firstname", "glossary", "honorific", "lineage", "lot", "othername", "part", "personname",
        "phrase", "preface", "primary", "reference", "setindex", "spanspec", "surname", "term", "textobject",
        "tfoot", "titleabbrev", "toc", "varlistentry"
    };

    private static bool IsKnownUntypedDocBookElement(System.Xml.Linq.XName name, System.Xml.Linq.XNamespace sourceNamespace) =>
        name.Namespace == sourceNamespace && KnownUntypedDocBookElementNames.Contains(name.LocalName);

    private static bool IsDerivedBlock(OfficeDocumentModelBlock block, ILookup<string, OfficeDocumentModelNode> nodesById) =>
        !string.IsNullOrEmpty(block.Id) && block.Marker == null && block.Region == null && nodesById[block.Id].Any(node =>
            string.Equals(node.Kind, block.Kind, StringComparison.OrdinalIgnoreCase) && node.Level == block.Level &&
            (string.Equals(node.Text, block.Text, StringComparison.Ordinal) ||
             ShouldReplaceChildrenWithPrimaryText(node) &&
             string.Equals(GetRepresentedPrimaryChildText(node), block.Text, StringComparison.Ordinal)));

    private static bool ShouldReplaceChildrenWithPrimaryText(OfficeDocumentModelNode source) {
        if (source.Children.Count == 0 || string.Equals(source.Kind, "text", StringComparison.OrdinalIgnoreCase)) return false;
        bool acceptsDirectText = TryMapKind(source.Kind, out DocBookNodeKind kind) && NodeAcceptsDirectText(kind);
        return acceptsDirectText &&
            !string.Equals(source.Text, GetRepresentedPrimaryChildText(source), StringComparison.Ordinal);
    }

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
            string.Equals(asset.AltText, alternateText, StringComparison.Ordinal);
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

    private static bool IsDerivedLink(OfficeDocumentModelLink link, ILookup<string, OfficeDocumentModelNode> nodesById) {
        const string prefix = "docbook-link-";
        if (string.IsNullOrEmpty(link.Id) || !link.Id.StartsWith(prefix, StringComparison.Ordinal) ||
            link.DestinationPageNumber.HasValue || !string.IsNullOrWhiteSpace(link.DestinationMode) ||
            !string.IsNullOrWhiteSpace(link.NamedAction) || !string.IsNullOrWhiteSpace(link.RemoteFile) ||
            !string.IsNullOrWhiteSpace(link.RemoteDestinationName) || link.RemoteDestinationPageNumber.HasValue) return false;
        string nodeId = "docbook-" + link.Id.Substring(prefix.Length);
        return nodesById[nodeId].Any(node =>
            (string.Equals(node.Kind, "link", StringComparison.OrdinalIgnoreCase) ||
             string.Equals(node.Kind, "cross-reference", StringComparison.OrdinalIgnoreCase)) &&
            string.Equals(node.Kind, link.Kind, StringComparison.OrdinalIgnoreCase) &&
            (link.Text == null || string.Equals(node.Text, link.Text, StringComparison.Ordinal) ||
             ShouldReplaceChildrenWithPrimaryText(node) &&
             string.Equals(GetRepresentedPrimaryChildText(node), link.Text, StringComparison.Ordinal)) &&
            LinkTargetMatches(node, link));
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
                !string.Equals(node.Text, table.Title ?? string.Empty, StringComparison.Ordinal) ||
                node.Location?.TableIndex != tableLocation.TableIndex) continue;
            string expectedPath = OfficeDocumentHeadingPath.Append(node.Location?.HeadingPath, table.Title, " / ");
            if (!string.Equals(expectedPath, tableLocation.HeadingPath, StringComparison.Ordinal)) continue;
            consumedNodes.Add(node);
            return true;
        }
        return false;
    }

    private static string ComputeTablePayloadHash(OfficeDocumentModelTable table) {
        var value = new StringBuilder();
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
