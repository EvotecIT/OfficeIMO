using OfficeIMO;
using OfficeIMO.Core.Internal;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.DocBook;

public sealed partial class DocBookDocument {
    private static bool IsDerivedBlock(OfficeDocumentModelBlock block, IEnumerable<OfficeDocumentModelNode> nodes) =>
        !string.IsNullOrEmpty(block.Id) && block.Marker == null && block.Region == null && nodes.Any(node =>
            string.Equals(node.Id, block.Id, StringComparison.Ordinal) &&
            string.Equals(node.Kind, block.Kind, StringComparison.OrdinalIgnoreCase) &&
            string.Equals(node.Text, block.Text, StringComparison.Ordinal) && node.Level == block.Level);

    private static bool IsDerivedAsset(OfficeDocumentModelAsset asset, IEnumerable<OfficeDocumentModelNode> nodes) {
        const string prefix = "docbook-image-";
        if (string.IsNullOrEmpty(asset.Id) || !asset.Id.StartsWith(prefix, StringComparison.Ordinal)) return false;
        string nodeId = "docbook-" + asset.Id.Substring(prefix.Length);
        string? reference = string.IsNullOrWhiteSpace(asset.SourceObjectId) ? asset.FileName : asset.SourceObjectId;
        return reference != null && nodes.Any(node =>
            string.Equals(node.Id, nodeId, StringComparison.Ordinal) &&
            string.Equals(node.Kind, "image", StringComparison.OrdinalIgnoreCase) &&
            node.Attributes.TryGetValue("fileref", out string? value) && string.Equals(value, reference, StringComparison.Ordinal));
    }

    private static bool IsDerivedLink(OfficeDocumentModelLink link, IEnumerable<OfficeDocumentModelNode> nodes) {
        const string prefix = "docbook-link-";
        if (string.IsNullOrEmpty(link.Id) || !link.Id.StartsWith(prefix, StringComparison.Ordinal) ||
            link.DestinationPageNumber.HasValue || !string.IsNullOrWhiteSpace(link.DestinationMode) ||
            !string.IsNullOrWhiteSpace(link.NamedAction) || !string.IsNullOrWhiteSpace(link.RemoteFile) ||
            !string.IsNullOrWhiteSpace(link.RemoteDestinationName) || link.RemoteDestinationPageNumber.HasValue) return false;
        string nodeId = "docbook-" + link.Id.Substring(prefix.Length);
        return nodes.Any(node =>
            string.Equals(node.Id, nodeId, StringComparison.Ordinal) &&
            (string.Equals(node.Kind, "link", StringComparison.OrdinalIgnoreCase) ||
             string.Equals(node.Kind, "cross-reference", StringComparison.OrdinalIgnoreCase)) &&
            (link.Text == null || string.Equals(node.Text, link.Text, StringComparison.Ordinal)) &&
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
        IEnumerable<OfficeDocumentModelNode> nodes,
        ISet<OfficeDocumentModelNode> consumedNodes) {
        if (table.Location == null || !table.Location.TableIndex.HasValue) return false;
        OfficeDocumentModelLocation tableLocation = table.Location;
        string expectedKind = string.Equals(table.Kind, "informaltable", StringComparison.OrdinalIgnoreCase)
            ? "informal-table" : "table";
        foreach (OfficeDocumentModelNode node in nodes) {
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
}
