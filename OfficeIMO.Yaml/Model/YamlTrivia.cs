namespace OfficeIMO.Yaml;

/// <summary>
/// Reads comments from the source gaps around nodes on demand. The gaps between node spans hold only whitespace, line
/// breaks, comments, byte-order marks and the indicators <c>- ? : ,</c>, so every <c>#</c> found in one starts a comment.
/// The attachment rules are documented on <see cref="YamlNode.LeadingComments"/> and <see cref="YamlNode.TrailingComment"/>.
/// </summary>
internal static class YamlTrivia {
    public static IReadOnlyList<string> GetLeadingComments(YamlNode node) {
        // A block collection without properties starts where its first child starts; that child owns the comments.
        if (IsTransparentBlock(node)) return Array.Empty<string>();
        int from = GapStart(node, out bool afterNode);
        string source = node.Document._source;
        // A gap that follows a node on the same line starts with that node's trailing comment.
        bool skipFirstLine = afterNode && from > 0 && !IsBreak(source[from - 1]);
        List<string>? comments = null;
        int i = from, end = node._start;
        while (i < end) {
            char ch = source[i];
            if (ch == '\r' || ch == '\n') {
                skipFirstLine = false;
                i++;
            } else if (ch == '#') {
                int stop = LineEnd(source, i, end);
                if (!skipFirstLine) (comments ??= new List<string>()).Add(source.Substring(i + 1, stop - i - 1));
                i = stop;
            } else {
                i++;
            }
        }
        return (IReadOnlyList<string>?)comments ?? Array.Empty<string>();
    }

    public static string? GetTrailingComment(YamlNode node) {
        string source = node.Document._source;
        if (node is YamlScalar { Style: YamlStyle.Literal or YamlStyle.Folded } block) return BlockHeaderComment(source, block);
        if (node._style == YamlStyle.Block) return null;
        int i = node._end;
        while (i < source.Length) {
            char ch = source[i];
            if (ch == ' ' || ch == '\t' || ch == ',' || ch == ':') {
                i++;
                continue;
            }
            // A '#' is a comment only after whitespace; directly after a node or indicator it cannot occur in a gap.
            if (ch != '#' || (source[i - 1] != ' ' && source[i - 1] != '\t')) return null;
            // A key whose omitted value ends before the comment leaves the comment to that value.
            if (node._parent is YamlMapping mapping) {
                int index = mapping.IndexOfNode(node, out bool isKey);
                if (isKey && mapping._entries[index].Value._start <= i) return null;
            }
            return source.Substring(i + 1, LineEnd(source, i, source.Length) - i - 1);
        }
        return null;
    }

    /// <summary>Finds the comment after a block scalar's <c>|</c> or <c>&gt;</c> header, skipping the node's properties.</summary>
    private static string? BlockHeaderComment(string source, YamlScalar scalar) {
        int i = scalar._start, end = scalar._valueStart;
        bool boundary = true;
        while (i < end) {
            char ch = source[i];
            if (ch == ' ' || ch == '\t' || IsBreak(ch)) {
                boundary = true;
                i++;
            } else if (boundary && ch == '#') {
                i = LineEnd(source, i, end);
            } else if (boundary && (ch == '|' || ch == '>')) {
                i++;
                while (i < end && source[i] != ' ' && source[i] != '\t' && !IsBreak(source[i])) i++;
                while (i < end && (source[i] == ' ' || source[i] == '\t')) i++;
                return i < end && source[i] == '#' ? source.Substring(i + 1, LineEnd(source, i, end) - i - 1) : null;
            } else {
                boundary = false;
                i++;
            }
        }
        return null;
    }

    /// <summary>
    /// Returns where the gap before a node starts: the end of the previous sibling or of the node's key, the content start
    /// of its collection, or the start of its document's content. <paramref name="afterNode"/> tells whether that point
    /// is the end of a node.
    /// </summary>
    private static int GapStart(YamlNode node, out bool afterNode) {
        afterNode = false;
        switch (node._parent) {
            case YamlMapping mapping: {
                int index = mapping.IndexOfNode(node, out bool isKey);
                if (!isKey) {
                    afterNode = true;
                    return mapping._entries[index].Key._end;
                }
                if (index > 0) {
                    afterNode = true;
                    return mapping._entries[index - 1].Value._end;
                }
                return IsTransparentBlock(mapping) ? GapStart(mapping, out afterNode) : mapping._contentStart;
            }
            case YamlSequence sequence: {
                int index = sequence.IndexOf(node);
                if (index > 0) {
                    afterNode = true;
                    return sequence._items[index - 1]._end;
                }
                return IsTransparentBlock(sequence) ? GapStart(sequence, out afterNode) : sequence._contentStart;
            }
            default:
                return DocumentContentStart(node._part);
        }
    }

    /// <summary>The root's gap starts after <c>---</c>, or after the previous document when there is no marker.</summary>
    private static int DocumentContentStart(YamlDocumentPart part) {
        if (part.HasStartMarker) return part._markerEnd;
        IReadOnlyList<YamlDocumentPart> parts = part._document.Documents;
        for (int i = 0; i < parts.Count; i++) {
            if (ReferenceEquals(parts[i], part)) return i == 0 ? 0 : parts[i - 1]._end;
        }
        return part._start;
    }

    private static bool IsTransparentBlock(YamlNode node) =>
        node._style == YamlStyle.Block && node._properties is null && node is YamlMapping or YamlSequence;

    private static int LineEnd(string source, int from, int limit) {
        int i = from;
        while (i < limit && !IsBreak(source[i])) i++;
        return i;
    }

    private static bool IsBreak(char ch) => ch == '\r' || ch == '\n';
}
