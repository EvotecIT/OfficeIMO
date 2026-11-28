using System.Text;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word.OpenXml;

/// <summary>
/// Run text on owned trees, with the rules of <see cref="WordParagraph.Text"/>: tabs read as <c>\t</c>, text-wrapping
/// breaks and carriage returns as <c>\n</c>, other breaks as U+2028, and text without <c>xml:space="preserve"</c> loses
/// its edge whitespace. Runs holding anything these rules don't cover (drawings, fields, comment references and the
/// like) are left to the SDK path.
/// </summary>
internal static class WordOwnedText {
    private const int W = OfficeXmlNamespaces.WordprocessingML;
    private const char NonTextBreakPlaceholder = '\u2028';
    private static readonly char[] EdgeWhitespace = { ' ', '\t', '\r', '\n' };
    private static readonly char[] SimpleRunUnsupportedCharacters = { '\r', '\n', '\t', NonTextBreakPlaceholder };

    internal static bool TryRead(OfficeXmlTree tree, int run, out string text) {
        text = string.Empty;
        string? single = null;
        StringBuilder? builder = null;
        for (int child = tree.FirstChildOf(run); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
            OfficeXmlTreeNodeKind kind = tree.KindOf(child);
            // Character data and comments between a run's elements carry no text on the SDK path either.
            if (kind != OfficeXmlTreeNodeKind.Element) continue;
            if (tree.NamespaceOf(child) != W) {
                if (tree.FirstChildOf(child) != OfficeXmlTree.NoNode) return false;
                continue;
            }
            string? part;
            switch (tree.LocalNameOf(child)) {
                case "rPr":
                    continue;
                case "t":
                    if (!TryReadTextElement(tree, child, out part)) return false;
                    break;
                case "tab":
                case "ptab":
                    part = "\t";
                    break;
                case "cr":
                    part = "\n";
                    break;
                case "noBreakHyphen":
                    part = "\u2011";
                    break;
                case "softHyphen":
                    part = "\u00ad";
                    break;
                case "br":
                    if (!TryReadBreak(tree, child, out bool textWrapping)) return false;
                    part = textWrapping ? "\n" : NonTextBreakPlaceholder.ToString();
                    break;
                // A hidden comment reference hides the run's text; those runs stay on the SDK path.
                case "commentReference":
                    return false;
                case "fldChar":
                case "instrText":
                    // Field runs are read as fields by the SDK path's paragraph enumeration.
                    return false;
                default:
                    // Elements with content, such as drawings, objects and ruby text, can contribute text of their own.
                    if (tree.FirstChildOf(child) != OfficeXmlTree.NoNode) return false;
                    continue;
            }
            if (part.Length == 0) continue;
            if (single == null) single = part;
            else (builder ??= new StringBuilder(single)).Append(part);
        }
        text = builder?.ToString() ?? single ?? string.Empty;
        return true;
    }

    private static bool TryReadTextElement(OfficeXmlTree tree, int element, out string text) {
        text = string.Empty;
        for (int child = tree.FirstChildOf(element); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
            if (tree.KindOf(child) != OfficeXmlTreeNodeKind.Text) return false;
        }
        string value = tree.GetInnerText(element);
        bool preserve = tree.AttributeEquals(element, OfficeXmlNamespaces.Xml, "space", "preserve");
        text = preserve ? value : value.Trim(EdgeWhitespace);
        return true;
    }

    private static bool TryReadBreak(OfficeXmlTree tree, int element, out bool textWrapping) {
        textWrapping = true;
        if (!tree.TryGetAttribute(element, W, "type", out string type)) return true;
        switch (type) {
            case "textWrapping":
                return true;
            case "page":
            case "column":
                textWrapping = false;
                return true;
            default:
                return false;
        }
    }

    /// <summary>Whether <see cref="Write"/> covers the run: it holds only the elements the text rules know.</summary>
    internal static bool CanWrite(OfficeXmlTree tree, int run) {
        for (int child = tree.FirstChildOf(run); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
            if (tree.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
            if (tree.NamespaceOf(child) != W) return false;
            switch (tree.LocalNameOf(child)) {
                case "rPr":
                case "t":
                case "tab":
                case "ptab":
                case "cr":
                case "noBreakHyphen":
                case "softHyphen":
                case "lastRenderedPageBreak":
                    continue;
                case "br":
                    if (!TryReadBreak(tree, child, out _)) return false;
                    continue;
                default:
                    return false;
            }
        }
        return true;
    }

    /// <summary>
    /// Replaces the text of <paramref name="run"/> as the SDK path's setter does: text, tabs and text-wrapping breaks
    /// are rewritten from <paramref name="value"/>, page and column breaks keep their positions, and every text element
    /// preserves its whitespace.
    /// </summary>
    internal static void Write(OfficeXmlTree tree, int run, string value) {
        if (value.IndexOfAny(SimpleRunUnsupportedCharacters) < 0 && TrySetSimpleText(tree, run, value)) return;

        var preservedBreaks = new List<(int ContentIndex, int Node)>();
        int contentNodesEncountered = 0;
        int child = tree.FirstChildOf(run);
        while (child != OfficeXmlTree.NoNode) {
            int next = tree.NextSiblingOf(child);
            if (tree.KindOf(child) == OfficeXmlTreeNodeKind.Element && tree.NamespaceOf(child) == W) {
                switch (tree.LocalNameOf(child)) {
                    case "t":
                        TryReadTextElement(tree, child, out string text);
                        tree.Remove(child);
                        contentNodesEncountered += CountContentUnits(text);
                        break;
                    case "tab":
                    case "ptab":
                        tree.Remove(child);
                        contentNodesEncountered++;
                        break;
                    case "br":
                        TryReadBreak(tree, child, out bool textWrapping);
                        tree.Remove(child);
                        if (textWrapping) contentNodesEncountered++;
                        else preservedBreaks.Add((contentNodesEncountered, child));
                        break;
                }
            }
            child = next;
        }

        string normalized = value.Replace("\r\n", "\n").Replace("\r", "\n");
        List<(string Text, bool EndsWithTextWrappingBreak)> segments = BuildSegments(normalized);
        int emitted = 0;
        int preservedIndex = 0;

        void AppendPreserved(int contentIndex) {
            while (preservedIndex < preservedBreaks.Count && preservedBreaks[preservedIndex].ContentIndex == contentIndex) {
                tree.Append(run, preservedBreaks[preservedIndex].Node);
                preservedIndex++;
            }
        }

        void AppendTextWithTabs(string segment) {
            if (segment.Length == 0) {
                AppendText(tree, run, string.Empty);
                emitted++;
                AppendPreserved(emitted);
                return;
            }
            string[] parts = segment.Split('\t');
            for (int index = 0; index < parts.Length; index++) {
                if (parts[index].Length > 0) {
                    AppendText(tree, run, parts[index]);
                    emitted++;
                    AppendPreserved(emitted);
                }
                if (index < parts.Length - 1) {
                    tree.Append(run, tree.CreateElement(W, "tab"));
                    emitted++;
                    AppendPreserved(emitted);
                }
            }
        }

        AppendPreserved(0);
        for (int index = 0; index < segments.Count; index++) {
            (string segment, bool endsWithBreak) = segments[index];
            bool isLast = index == segments.Count - 1;
            if (segment.Length > 0 || segments.Count == 1 || isLast) AppendTextWithTabs(segment);
            if (endsWithBreak) {
                tree.Append(run, tree.CreateElement(W, "br"));
                emitted++;
                AppendPreserved(emitted);
            }
        }
        while (preservedIndex < preservedBreaks.Count) {
            tree.Append(run, preservedBreaks[preservedIndex].Node);
            preservedIndex++;
        }
    }

    private static bool TrySetSimpleText(OfficeXmlTree tree, int run, string value) {
        int firstText = OfficeXmlTree.NoNode;
        for (int child = tree.FirstChildOf(run); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
            if (tree.KindOf(child) != OfficeXmlTreeNodeKind.Element) continue;
            if (tree.IsElement(child, W, "rPr")) continue;
            if (!tree.IsElement(child, W, "t")) return false;
            if (firstText == OfficeXmlTree.NoNode) firstText = child;
        }
        if (firstText == OfficeXmlTree.NoNode) {
            AppendText(tree, run, value);
            return true;
        }
        int sibling = tree.NextSiblingOf(firstText);
        while (sibling != OfficeXmlTree.NoNode) {
            int next = tree.NextSiblingOf(sibling);
            if (tree.IsElement(sibling, W, "t")) tree.Remove(sibling);
            sibling = next;
        }
        tree.SetInnerText(firstText, value);
        tree.SetAttribute(firstText, OfficeXmlNamespaces.Xml, "space", "preserve");
        return true;
    }

    private static void AppendText(OfficeXmlTree tree, int run, string text) {
        int element = tree.CreateElement(W, "t");
        tree.SetAttribute(element, OfficeXmlNamespaces.Xml, "space", "preserve");
        if (text.Length > 0) tree.Append(element, tree.CreateText(text));
        tree.Append(run, element);
    }

    private static int CountContentUnits(string text) {
        if (text.Length == 0) return 0;
        int units = 0;
        string[] parts = text.Split('\t');
        for (int index = 0; index < parts.Length; index++) {
            if (parts[index].Length > 0) units++;
            if (index < parts.Length - 1) units++;
        }
        return units;
    }

    private static List<(string Text, bool EndsWithTextWrappingBreak)> BuildSegments(string source) {
        var result = new List<(string, bool)>();
        foreach (string block in source.Split(NonTextBreakPlaceholder)) {
            string[] lines = block.Split('\n');
            for (int index = 0; index < lines.Length; index++) {
                result.Add((lines[index], index < lines.Length - 1));
            }
        }
        return result;
    }
}
