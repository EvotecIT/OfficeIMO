using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word.OpenXml;

/// <summary>An element of an owned tree, returned by edits so attributes can follow.</summary>
internal readonly struct WordOwnedElement {
    internal WordOwnedElement(OfficeXmlTree tree, int node) {
        Tree = tree;
        Node = node;
    }

    internal OfficeXmlTree Tree { get; }
    internal int Node { get; }

    internal WordOwnedElement SetValue(int namespaceId, string localName, string value) {
        Tree.SetAttribute(Node, namespaceId, localName, value);
        return this;
    }
}

/// <summary>Edits on owned WordprocessingML trees with the results the SDK's typed members give.</summary>
internal static class WordOwnedEdits {
    private const int W = OfficeXmlNamespaces.WordprocessingML;

    /// <summary>
    /// Brings the body to the state an SDK load leaves it in before the first edit: final section properties present
    /// and last. Content added afterwards then lands where the SDK path puts it.
    /// </summary>
    internal static void PrepareBodyEdit(WordPackageModel model) {
        if (model.BodyPrepared) return;
        WordOwnedSave.EnsureBodySectionProperties(model);
        model.BodyPrepared = true;
    }

    /// <summary>
    /// Puts a new, empty element in place of the first child with this name, or inserts one in schema order, as
    /// assigning a new element to a typed SDK property does.
    /// </summary>
    internal static WordOwnedElement ReplaceChild(OfficeXmlTree tree, int parent, string localName, OfficeXmlChildOrder order) {
        int replacement = tree.CreateElement(W, localName);
        int existing = tree.FindChild(parent, W, localName);
        if (existing != OfficeXmlTree.NoNode) {
            tree.InsertBefore(parent, replacement, existing);
            tree.Remove(existing);
        } else {
            tree.InsertOrdered(parent, replacement, order);
        }
        return new WordOwnedElement(tree, replacement);
    }

    /// <summary>Removes the first child with this name, as clearing a typed SDK property does.</summary>
    internal static void RemoveChild(OfficeXmlTree tree, int parent, string localName) {
        int existing = tree.FindChild(parent, W, localName);
        if (existing != OfficeXmlTree.NoNode) tree.Remove(existing);
    }

    /// <summary>
    /// Reads an on/off property the way OfficeIMO's SDK path does: present with no value, or a true value, is on.
    /// Returns false for a value the SDK would not parse, so the caller can leave that case to it.
    /// </summary>
    internal static bool TryReadOnOff(OfficeXmlTree tree, int element, out bool value) {
        value = false;
        if (element == OfficeXmlTree.NoNode) return true;
        if (!tree.HasAttribute(element, W, "val")) {
            value = true;
            return true;
        }
        // The usual spellings are compared as written, without decoding the value.
        if (tree.AttributeEquals(element, W, "val", "true") || tree.AttributeEquals(element, W, "val", "1")) {
            value = true;
            return true;
        }
        if (tree.AttributeEquals(element, W, "val", "false") || tree.AttributeEquals(element, W, "val", "0")) return true;
        string token = tree.GetAttribute(element, W, "val")!;
        switch (token) {
            case "true": case "1": case "on":
                value = true;
                return true;
            case "false": case "0": case "off":
                return true;
            default:
                return false;
        }
    }
}
