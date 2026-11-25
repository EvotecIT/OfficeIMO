using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.PowerPoint.OpenXml;

/// <summary>
/// Text of DrawingML text bodies on a package tree, read and written as <see cref="PowerPointParagraph"/> and
/// <see cref="PowerPointTextContent"/> do it on SDK elements.
/// </summary>
internal static class PowerPointTreeText {
    /// <summary>The body's paragraphs joined with line breaks; false when a run holds text the SDK reads differently.</summary>
    internal static bool TryReadBody(OfficeXmlTree tree, int body, out string text) {
        text = string.Empty;
        System.Text.StringBuilder? builder = null;
        string? single = null;
        for (int paragraph = tree.FindChild(body, OfficeXmlNamespaces.DrawingML, "p"); paragraph != OfficeXmlTree.NoNode;
             paragraph = tree.FindNextSibling(paragraph, OfficeXmlNamespaces.DrawingML, "p")) {
            if (!TryReadParagraph(tree, paragraph, out string part)) return false;
            if (builder != null) builder.Append(Environment.NewLine).Append(part);
            else if (single == null) single = part;
            else builder = new System.Text.StringBuilder(single).Append(Environment.NewLine).Append(part);
        }
        text = builder?.ToString() ?? single ?? string.Empty;
        return true;
    }

    /// <summary>A paragraph's runs, breaks and fields as text.</summary>
    internal static bool TryReadParagraph(OfficeXmlTree tree, int paragraph, out string text) {
        text = string.Empty;
        System.Text.StringBuilder? builder = null;
        string? single = null;
        for (int child = tree.FirstChildOf(paragraph); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
            if (tree.KindOf(child) != OfficeXmlTreeNodeKind.Element || tree.NamespaceOf(child) != OfficeXmlNamespaces.DrawingML) continue;
            string? part;
            switch (tree.LocalNameOf(child)) {
                case "r":
                case "fld":
                    if (!TryReadText(tree, tree.FindChild(child, OfficeXmlNamespaces.DrawingML, "t"), out part)) return false;
                    break;
                case "br":
                    part = Environment.NewLine;
                    break;
                default:
                    continue;
            }
            if (builder != null) builder.Append(part);
            else if (single == null) single = part;
            else builder = new System.Text.StringBuilder(single).Append(part);
        }
        text = builder?.ToString() ?? single ?? string.Empty;
        return true;
    }

    // The text of an a:t with one text node at most; the SDK reads only part of several (text and CDATA side by side).
    private static bool TryReadText(OfficeXmlTree tree, int element, out string text) {
        text = string.Empty;
        if (element == OfficeXmlTree.NoNode) return true;
        int child = tree.FirstChildOf(element);
        if (child == OfficeXmlTree.NoNode) return true;
        if (tree.KindOf(child) != OfficeXmlTreeNodeKind.Text || tree.NextSiblingOf(child) != OfficeXmlTree.NoNode) return false;
        text = tree.GetText(child);
        return true;
    }

    /// <summary>
    /// Appends one run per line, with <c>a:br</c> between them, each carrying a copy of <paramref name="properties"/> (an
    /// <c>a:rPr</c> in the same tree, or <see cref="OfficeXmlTree.NoNode"/>), as CreateInlineContent does.
    /// </summary>
    internal static void AppendInlineContent(OfficeXmlTree tree, int paragraph, string? text, int properties) {
        bool first = true;
        foreach (string line in PowerPointTextContent.SplitLines(text)) {
            if (!first) {
                int lineBreak = Append(tree, paragraph, "br");
                // A break cannot be clicked, so its properties keep no interaction.
                if (properties != OfficeXmlTree.NoNode) tree.Append(lineBreak, WithoutInteractions(tree, tree.CloneElement(properties)));
            }
            first = false;
            int run = Append(tree, paragraph, "r");
            if (properties != OfficeXmlTree.NoNode) tree.Append(run, tree.CloneElement(properties));
            int value = Append(tree, run, "t");
            if (line.Length > 0) tree.Append(value, tree.CreateText(line));
        }
    }

    /// <summary>Removes the click and hover hyperlinks of a copied <c>a:rPr</c>.</summary>
    internal static int WithoutInteractions(OfficeXmlTree tree, int properties) {
        for (int child = tree.FirstChildOf(properties); child != OfficeXmlTree.NoNode;) {
            int next = tree.NextSiblingOf(child);
            if (tree.KindOf(child) == OfficeXmlTreeNodeKind.Element && tree.NamespaceOf(child) == OfficeXmlNamespaces.DrawingML
                && tree.LocalNameOf(child) is "hlinkClick" or "hlinkMouseOver") {
                tree.Remove(child);
            }
            child = next;
        }
        return properties;
    }

    /// <summary>Whether a subtree refers to an embedded sound, which the model also removes from the part when unused.</summary>
    internal static bool HasSound(OfficeXmlTree tree, int node) {
        for (int child = tree.FirstChildOf(node); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
            if (tree.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
            if (tree.NamespaceOf(child) == OfficeXmlNamespaces.DrawingML && tree.LocalNameOf(child) == "snd") return true;
            if (HasSound(tree, child)) return true;
        }
        return false;
    }

    internal static int Append(OfficeXmlTree tree, int parent, string localName) {
        int element = tree.CreateElement(OfficeXmlNamespaces.DrawingML, localName);
        tree.Append(parent, element);
        return element;
    }
}
