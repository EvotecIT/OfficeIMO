using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

public sealed partial class RtfLosslessEditor {
    /// <summary>
    /// Adds, replaces, or removes an XML namespace declaration in the root <c>xmlnstbl</c> destination while preserving the rest of the RTF stream.
    /// </summary>
    public void SetXmlNamespace(int id, string? uri) {
        if (id < 0) {
            throw new ArgumentOutOfRangeException(nameof(id), "XML namespace id cannot be negative.");
        }

        RtfGroup root = SetXmlNamespace(_syntaxTree.Root, id, uri);
        _syntaxTree = new RtfSyntaxTree(root, _syntaxTree.Diagnostics);
    }

    /// <summary>
    /// Removes all XML namespace declarations with the supplied id while preserving the rest of the RTF stream.
    /// </summary>
    public void RemoveXmlNamespace(int id) {
        SetXmlNamespace(id, null);
    }

    private static RtfGroup SetXmlNamespace(RtfGroup root, int id, string? uri) {
        var children = new List<RtfNode>(root.Children);
        int namespaceTableIndex = children.FindIndex(node => node is RtfGroup group && group.Destination == "xmlnstbl");

        if (namespaceTableIndex >= 0) {
            RtfGroup? updatedTable = SetXmlNamespaceInTable((RtfGroup)children[namespaceTableIndex], id, uri);
            if (updatedTable == null) {
                children.RemoveAt(namespaceTableIndex);
            } else {
                children[namespaceTableIndex] = updatedTable;
            }

            return new RtfGroup(root.Position, children);
        }

        if (string.IsNullOrEmpty(uri)) {
            return root;
        }

        children.Insert(GetXmlNamespaceInsertIndex(children), CreateXmlNamespaceTable(id, uri!));
        return new RtfGroup(root.Position, children);
    }

    private static RtfGroup? SetXmlNamespaceInTable(RtfGroup namespaceTable, int id, string? uri) {
        var children = new List<RtfNode>(namespaceTable.Children);
        bool replaced = false;

        for (int index = children.Count - 1; index >= 0; index--) {
            if (children[index] is not RtfGroup namespaceGroup || !XmlNamespaceIdMatches(namespaceGroup, id)) {
                continue;
            }

            if (string.IsNullOrEmpty(uri) || replaced) {
                children.RemoveAt(index);
            } else {
                children[index] = CreateXmlNamespaceGroup(id, uri!);
                replaced = true;
            }
        }

        if (!replaced && !string.IsNullOrEmpty(uri)) {
            children.Add(CreateXmlNamespaceGroup(id, uri!));
        }

        return children.Any(node => node is RtfGroup group && group.Destination == "xmlns")
            ? new RtfGroup(namespaceTable.Position, children)
            : null;
    }

    private static bool XmlNamespaceIdMatches(RtfGroup namespaceGroup, int id) {
        return namespaceGroup.Destination == "xmlns" &&
               namespaceGroup.Children.OfType<RtfControlWord>().Any(control => control.Name == "xmlns" && control.Parameter == id);
    }

    private static RtfGroup CreateXmlNamespaceTable(int id, string uri) {
        return new RtfGroup(0, new RtfNode[] {
            new RtfControlSymbol(0, '*', null, hasParameter: false, rawText: @"\*"),
            new RtfControlWord(0, "xmlnstbl", null, hasParameter: false, rawText: @"\xmlnstbl"),
            CreateXmlNamespaceGroup(id, uri)
        });
    }

    private static RtfGroup CreateXmlNamespaceGroup(int id, string uri) {
        return new RtfGroup(0, new RtfNode[] {
            new RtfControlWord(
                0,
                "xmlns",
                id,
                hasParameter: true,
                rawText: @"\xmlns" + id.ToString(CultureInfo.InvariantCulture) + " "),
            new RtfText(0, uri + ";", RtfTextEncoding.EncodeText(uri + ";"))
        });
    }

    private static int GetXmlNamespaceInsertIndex(List<RtfNode> children) {
        int insertIndex = -1;
        for (int index = 0; index < children.Count; index++) {
            if (children[index] is RtfGroup group &&
                (group.Destination == "xmlnstbl" || group.Destination == "filetbl" ||
                 group.Destination == "fonttbl" || group.Destination == "generator")) {
                insertIndex = index;
            }
        }

        return insertIndex >= 0 ? insertIndex + 1 : GetInfoInsertIndex(children);
    }
}
