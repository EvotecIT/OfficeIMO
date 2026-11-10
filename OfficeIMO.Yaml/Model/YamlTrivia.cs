namespace OfficeIMO.Yaml;

/// <summary>
/// Reads and locates comments and blank lines in the source gaps around nodes. The gaps between node spans hold only
/// whitespace, line breaks, comments, byte-order marks and the indicators <c>- ? : ,</c>, so every <c>#</c> found at
/// the start of a gap line starts a comment.
/// </summary>
/// <remarks>
/// <para><b>Ownership.</b> Comment lines and blank lines above an entry belong to the line where the entry starts. The
/// owner of that line is the outermost node starting on it: a mapping value's owner is its key, and the first entry of
/// a block collection without properties that is a sequence item or the root defers to that collection. Above the
/// first entry of a document without <c>---</c>, comment lines followed by a blank line belong to the document part
/// (its header). Comment lines after the root's last line belong to the part (its footer).</para>
/// <para><b>Regions.</b> The region before an owner runs from the line after the previous content to the owner's line:
/// first the blank lines (<see cref="YamlNode.SpaceBefore"/>), then from the first comment line on, the comment block
/// (<see cref="YamlNode.CommentBefore"/>). Removing an entry removes its comment block, its lines and its trailing
/// comment; blank lines above its comments stay as separators for the neighbors.</para>
/// </remarks>
internal static class YamlTrivia {
    // ------------------------------------------------------------------ public getters

    public static string? GetCommentBefore(YamlNode node) {
        YamlNode owner = Owner(node);
        if (owner._properties is { CommentBeforeSet: true } edits) return edits.CommentBefore;
        if (!HasRegion(owner) || InFlow(owner)) return null;
        if (!TryBeforeRegion(owner, out _, out int commentStart, out int lineStart)) return null;
        return Comments(owner.Document!._source, commentStart, lineStart);
    }

    public static bool GetSpaceBefore(YamlNode node) {
        YamlNode owner = Owner(node);
        if (owner._properties is { SpaceBeforeSet: true } edits) return edits.SpaceBefore;
        if (!HasRegion(owner) || InFlow(owner)) return false;
        if (!TryBeforeRegion(owner, out int blankStart, out int commentStart, out _)) return false;
        return HasBlankLine(owner.Document!._source, blankStart, commentStart);
    }

    public static string? GetComment(YamlNode node) {
        YamlNode target = TrailingTarget(node);
        if (target._properties is { CommentSet: true } edits) return edits.Comment;
        // A block collection value and its key share the key's line; an edit may sit on either.
        if (!ReferenceEquals(target, node) && node._properties is { CommentSet: true } own) return own.Comment;
        if (target._parent is YamlMapping { _style: YamlStyle.Block } mapping) {
            int index = mapping.IndexOfNode(target, out bool isKey);
            YamlNode value = mapping._entries[index].Value;
            if (isKey && value is YamlMapping or YamlSequence && value._style == YamlStyle.Block && value._properties is { CommentSet: true } valueEdits) return valueEdits.Comment;
        }
        if (!target.HasSourceSpan) return null;
        int hash = TrailingHash(target, out int lineEnd);
        return hash < 0 ? null : CommentText(target.Document!._source, hash, lineEnd);
    }

    // ------------------------------------------------------------------ ownership

    /// <summary>Whether the node holds a place in the loaded source: its own text, or the span of the node it replaced.</summary>
    public static bool HasRegion(YamlNode node) => node._start >= 0 && node._part is { _new: false };

    /// <summary>Whether a transparent collection now starts with an entry added by editing, so its line has no source comments.</summary>
    private static bool StartsAnew(YamlNode owner) =>
        owner.IsTransparentBlock && owner.HasSourceSpan && FirstChild(owner) is { _start: < 0 };

    private static YamlNode? FirstChild(YamlNode collection) => collection switch {
        YamlMapping { _count: > 0 } m => m._entries[0].Key,
        YamlSequence { _count: > 0 } q => q._items[0],
        _ => null
    };

    /// <summary>Returns the node that owns the comment lines and blank lines above <paramref name="node"/>'s entry.</summary>
    public static YamlNode Owner(YamlNode node) {
        YamlNode current = node;
        while (true) {
            YamlNode? parent = current._parent;
            if (parent is YamlMapping map) {
                int index = map.IndexOfNode(current, out bool isKey);
                if (!isKey) {
                    // A mapping value's lines are its key's.
                    current = map._entries[index].Key;
                    continue;
                }
            }
            if (parent is null || !parent.IsTransparentBlock || !IsFirstChild(parent, current)) return current;
            // The first entry of a transparent sequence item or root shares the line with that collection.
            if (parent._parent is YamlSequence || (parent._parent is null && parent._part is not null && ReferenceEquals(parent._part._root, parent))) {
                current = parent;
                continue;
            }
            return current;
        }
    }

    /// <summary>Returns the node whose line end carries <paramref name="node"/>'s trailing comment.</summary>
    public static YamlNode TrailingTarget(YamlNode node) {
        if (node is YamlScalar && node.HasSourceSpan && node._parent is YamlMapping { _style: YamlStyle.Block } owner) {
            // An omitted value without its ':' has no line of its own; its key's line is the entry's.
            int index = owner.IndexOfNode(node, out bool isKey);
            if (!isKey && OmittedWithoutColon(node.Document!._source, owner._entries[index].Key, node)) return owner._entries[index].Key;
        }
        if (node is YamlMapping or YamlSequence && node._style == YamlStyle.Block && node._parent is YamlMapping map) {
            int index = map.IndexOfNode(node, out bool isKey);
            if (!isKey) return map._entries[index].Key;
        }
        return node;
    }

    /// <summary>Whether the node's own line is inside a flow collection, where comment lines cannot go.</summary>
    public static bool InFlow(YamlNode node) => node._parent is not null && node._parent._style == YamlStyle.Flow;

    private static bool IsFirstChild(YamlNode parent, YamlNode child) {
        return parent switch {
            YamlMapping m => m._count > 0 && ReferenceEquals(m._entries[0].Key, child),
            YamlSequence s => s._count > 0 && ReferenceEquals(s._items[0], child),
            _ => false
        };
    }

    // ------------------------------------------------------------------ gaps and regions

    /// <summary>
    /// Finds the region above a source-backed owner: blank lines from <paramref name="blankStart"/>, comment lines from
    /// <paramref name="commentStart"/>, up to <paramref name="lineStart"/>, the start of the owner's line. Returns false
    /// when the owner shares its line with earlier content (inside a flow collection, or after a key).
    /// </summary>
    public static bool TryBeforeRegion(YamlNode owner, out int blankStart, out int commentStart, out int lineStart) {
        string source = owner.Document!._source;
        // A collection without properties starts on its first entry's line. While that entry is still the loaded first
        // one, the gap is the collection's own; after earlier entries were removed, it is the entry's.
        YamlNode gapOwner = owner, lineOwner = owner;
        while (owner.IsTransparentBlock && FirstChild(owner) is { _start: >= 0 } first) {
            if (OriginalIndex(owner, first) != 0) gapOwner = first;
            // The outermost sequence item on the line decides it through its dash.
            if (!(lineOwner._parent is YamlSequence { _style: YamlStyle.Block }) && owner is YamlSequence { _style: YamlStyle.Block }) lineOwner = first;
            owner = first;
        }
        if (owner._start < 0 || StartsAnew(owner)) {
            blankStart = commentStart = lineStart = 0;
            return false;
        }
        // A sequence item's line is its dash's, whatever its first entry is now.
        int position = lineOwner._parent is YamlSequence { _style: YamlStyle.Block } && lineOwner._start >= 0 ? LinePosition(lineOwner) : LinePosition(owner);
        int gap = GapStart(gapOwner, out bool documentStart);
        return RegionAt(source, gap, documentStart, position, out blankStart, out commentStart, out lineStart);
    }

    /// <summary>
    /// Finds the region between the content ending at <paramref name="gap"/> and the line holding
    /// <paramref name="position"/>; see <see cref="TryBeforeRegion"/>. At a document start, comments that a blank line
    /// separates from the line are left out as the document's header.
    /// </summary>
    public static bool RegionAt(string source, int gap, bool documentStart, int position, out int blankStart, out int commentStart, out int lineStart) {
        lineStart = LineStart(source, position);
        blankStart = commentStart = lineStart;
        if (!IsBlankLine(source, lineStart, position)) return false;
        int from = AtLineStart(source, gap) ? gap : NextLineStart(source, gap);
        if (from > lineStart) return false;
        blankStart = from;
        int firstComment = -1;
        int line = from;
        bool header = false;
        while (line < lineStart) {
            int end = LineContentEnd(source, line);
            if (IsCommentLine(source, line, end)) {
                if (firstComment < 0) firstComment = line;
            } else if (documentStart && firstComment >= 0) {
                // Comments followed by a blank line at the top of a document belong to the document.
                header = true;
                firstComment = -1;
            }
            int next = AfterBreak(source, end);
            if (next == line) break;
            line = next;
        }
        commentStart = firstComment >= 0 ? firstComment : lineStart;
        if (header) blankStart = HeaderEnd(source, from, commentStart);
        return true;
    }

    /// <summary>The position after the last header comment line, where blank lines before the owner begin.</summary>
    private static int HeaderEnd(string source, int from, int limit) {
        int result = from;
        int line = from;
        while (line < limit) {
            int end = LineContentEnd(source, line);
            int next = AfterBreak(source, end);
            if (IsCommentLine(source, line, end)) result = next;
            line = next;
        }
        return result;
    }

    /// <summary>
    /// Returns where the gap before a node starts: the end of the previous loaded sibling (or of the node's key), the
    /// content start of a collection with properties, or the start of the document's content.
    /// </summary>
    public static int GapStart(YamlNode node, out bool documentStart) {
        documentStart = false;
        switch (node._parent) {
            case YamlMapping mapping: {
                int index = mapping.IndexOfNode(node, out bool isKey);
                if (!isKey) return ColonEnd(node.Document!._source, mapping._entries[index].Key._end);
                int original = OriginalIndex(mapping, node);
                if (original > 0) return OriginalEnd(mapping, original - 1);
                return CollectionGapStart(mapping, out documentStart);
            }
            case YamlSequence sequence: {
                int original = OriginalIndex(sequence, node);
                if (original > 0) return OriginalEnd(sequence, original - 1);
                return CollectionGapStart(sequence, out documentStart);
            }
            default:
                documentStart = !node._part!.HasStartMarker;
                return DocumentContentStart(node._part!);
        }
    }

    public static int CollectionGapStart(YamlNode collection, out bool documentStart) {
        documentStart = false;
        if (!collection.IsTransparentBlock) return collection is YamlMapping m ? m._contentStart : ((YamlSequence)collection)._contentStart;
        // The first entry of a sequence item starts after the item's dash, even on a later line.
        if (collection._parent is YamlSequence { _style: YamlStyle.Block }) return DashPosition(collection.Document!._source, collection._start) + 1;
        return GapStart(collection, out documentStart);
    }

    /// <summary>The position of a loaded child among the collection's loaded children, or its current index.</summary>
    public static int OriginalIndex(YamlNode collection, YamlNode child) {
        int[]? spans = Spans(collection);
        if (spans is not null && child._start >= 0) {
            // Loaded children keep source order, so their starts are sorted.
            int low = 0, high = spans.Length / 2 - 1;
            while (low <= high) {
                int mid = (low + high) >> 1;
                int start = spans[2 * mid];
                if (start == child._start) return mid;
                if (start < child._start) low = mid + 1;
                else high = mid - 1;
            }
            return -1;
        }
        return collection switch {
            YamlMapping m => m.IndexOfNode(child, out _),
            YamlSequence s => s.IndexOf(child),
            _ => -1
        };
    }

    /// <summary>The source end of the loaded child at an original index: a mapping entry's value end or an item's end.</summary>
    public static int OriginalEnd(YamlNode collection, int index) {
        int[]? spans = Spans(collection);
        if (spans is not null) return spans[index * 2 + 1];
        return collection is YamlMapping m ? m._entries[index].Value._end : ((YamlSequence)collection)._items[index]._end;
    }

    /// <summary>The snapshot of loaded child spans as (start, end) pairs, or null while nothing was inserted or removed.</summary>
    public static int[]? Spans(YamlNode collection) => collection switch {
        YamlMapping m => m._original,
        YamlSequence s => s._original,
        _ => null
    };

    /// <summary>The root's gap starts after <c>---</c>, or after the previous document when there is no marker.</summary>
    public static int DocumentContentStart(YamlDocumentPart part) {
        if (part.HasStartMarker) return part._markerEnd;
        IReadOnlyList<YamlDocumentPart> parts = part._document.Documents;
        for (int i = 0; i < parts.Count; i++) {
            if (ReferenceEquals(parts[i], part)) return i == 0 ? 0 : parts[i - 1]._end;
        }
        return part._start;
    }

    /// <summary>
    /// The position that decides an owner's line: the dash of a block sequence item, otherwise the node's start.
    /// </summary>
    public static int LinePosition(YamlNode owner) {
        if (owner._parent is YamlSequence { _style: YamlStyle.Block }) return DashPosition(owner.Document!._source, owner._start);
        if (owner._parent is YamlMapping { _style: YamlStyle.Block } mapping) return EntryStart(owner.Document!._source, owner._start, MappingStart(mapping));
        return owner._start;
    }

    /// <summary>
    /// Returns the <c>?</c> of an explicit block entry whose key starts at <paramref name="keyStart"/>, or the key's start.
    /// A <c>?</c> before the mapping's own start introduces an enclosing entry whose key is this mapping.
    /// </summary>
    public static int EntryStart(string source, int keyStart, int mappingStart) {
        int i = keyStart - 1;
        while (i >= 0 && (source[i] == ' ' || source[i] == '\t')) i--;
        if (i >= mappingStart && i >= 0 && source[i] == '?' && (i == 0 || source[i - 1] == ' ' || source[i - 1] == '\t' || IsBreak(source[i - 1]))) return i;
        return keyStart;
    }

    /// <summary>Where a mapping's content starts in the source: its first entry, after any anchor or tag.</summary>
    public static int MappingStart(YamlMapping mapping) => mapping.IsTransparentBlock ? mapping._start : mapping._contentStart;

    /// <summary>Finds the <c>-</c> before a block sequence item that starts at <paramref name="itemStart"/>.</summary>
    public static int DashPosition(string source, int itemStart) {
        int i = itemStart - 1;
        while (i >= 0 && source[i] != '-') i--;
        return i < 0 ? itemStart : i;
    }

    /// <summary>
    /// Finds the trailing comment's <c>#</c> for a source-backed node, or -1. <paramref name="lineEnd"/> receives the
    /// end of that line's content.
    /// </summary>
    public static int TrailingHash(YamlNode node, out int lineEnd) {
        string source = node.Document!._source;
        lineEnd = -1;
        if (node is YamlScalar { Style: YamlStyle.Literal or YamlStyle.Folded } block) return BlockHeaderHash(source, block, out lineEnd);
        if (node._style == YamlStyle.Block && node is YamlMapping or YamlSequence) return -1;
        if (node._parent is YamlMapping mapping && node._parent._style == YamlStyle.Block) {
            int index = mapping.IndexOfNode(node, out bool isKey);
            if (isKey) {
                // A key owns the comment on its line only when its value is a block collection, starts on a later line,
                // or is omitted together with its ':' (an explicit `? key # comment`).
                YamlNode value = mapping._entries[index].Value;
                int end = LineContentEnd(source, node._end);
                if (value._start <= end && !(value is YamlMapping or YamlSequence && value._style == YamlStyle.Block) && !OmittedWithoutColon(source, node, value)) return -1;
                lineEnd = end;
                return FindHash(source, node._end, end);
            }
            if (OmittedWithoutColon(source, mapping._entries[index].Key, node)) return -1;
        }
        int i = node._end;
        while (i < source.Length) {
            char ch = source[i];
            if (ch == ' ' || ch == '\t' || ch == ',' || ch == ':') {
                i++;
                continue;
            }
            // A '#' is a comment only after whitespace; directly after a node or indicator it cannot occur in a gap.
            if (ch != '#' || (source[i - 1] != ' ' && source[i - 1] != '\t')) return -1;
            lineEnd = LineContentEnd(source, i);
            return i;
        }
        return -1;
    }

    /// <summary>Whether a value is an empty node with no ':' after its key, which leaves nothing on the line of its own.</summary>
    private static bool OmittedWithoutColon(string source, YamlNode key, YamlNode value) =>
        value is YamlScalar && value._start == value._end && value._start == key._end && ColonEnd(source, key._end) == key._end;

    /// <summary>Finds a comment that starts after whitespace between two positions on one line.</summary>
    private static int FindHash(string source, int from, int end) {
        for (int i = from; i < end; i++) {
            if (source[i] == '#' && i > 0 && (source[i - 1] == ' ' || source[i - 1] == '\t')) return i;
        }
        return -1;
    }

    /// <summary>Finds the comment after a block scalar's <c>|</c> or <c>&gt;</c> header, skipping the node's properties.</summary>
    private static int BlockHeaderHash(string source, YamlScalar scalar, out int lineEnd) {
        int header = BlockHeaderPosition(source, scalar);
        lineEnd = LineContentEnd(source, header);
        return FindHash(source, header, lineEnd);
    }

    /// <summary>Returns the position of a block scalar's <c>|</c> or <c>&gt;</c> indicator.</summary>
    public static int BlockHeaderPosition(string source, YamlScalar scalar) {
        int i = scalar._start, end = scalar._valueStart;
        bool boundary = true;
        while (i < end) {
            char ch = source[i];
            if (ch == ' ' || ch == '\t' || IsBreak(ch)) {
                boundary = true;
                i++;
            } else if (boundary && ch == '#') {
                i = LineContentEnd(source, i);
            } else if (boundary && (ch == '|' || ch == '>')) {
                return i;
            } else {
                boundary = false;
                i++;
            }
        }
        return scalar._start;
    }

    // ------------------------------------------------------------------ document part comments

    /// <summary>The region holding a part's header comments: before <c>---</c>, or above the root's own comment block.</summary>
    public static bool TryHeaderRegion(YamlDocumentPart part, out int start, out int end) {
        string source = part._document._source;
        start = end = 0;
        if (part._root is null || !part._root.HasSourceSpan) return false;
        if (part.HasStartMarker) {
            int index = PartIndex(part);
            YamlDocumentPart? previous = index > 0 ? part._document.Documents[index - 1] : null;
            start = previous is null ? 0 : previous.HasEndMarker ? NextLineStart(source, previous._end - 1) : part._start;
            end = LineStart(source, part._start);
            return start <= end;
        }
        YamlNode owner = Owner(part._root);
        if (!owner.HasSourceSpan) return false;
        int gap = DocumentContentStart(part);
        start = AtLineStart(source, gap) ? gap : NextLineStart(source, gap);
        if (!TryBeforeRegion(owner, out int blankStart, out _, out _)) return false;
        end = blankStart;
        return true;
    }

    /// <summary>The region holding a part's footer comments: the lines after the root's last line, before the next marker.</summary>
    public static bool TryFooterRegion(YamlDocumentPart part, out int start, out int end) {
        string source = part._document._source;
        start = end = 0;
        if (part._root is null || !part._root.HasSourceSpan) return false;
        int rootEnd = part._root._end;
        start = AtLineStart(source, rootEnd) && rootEnd > 0 ? rootEnd : NextLineStart(source, rootEnd);
        int index = PartIndex(part);
        IReadOnlyList<YamlDocumentPart> parts = part._document.Documents;
        if (part.HasEndMarker) end = LineStart(source, part._end - 1);
        else if (index + 1 < parts.Count) end = LineStart(source, parts[index + 1]._start);
        else end = source.Length;
        if (end < start) end = start;
        return true;
    }

    public static string? PartComments(YamlDocumentPart part, bool header) {
        if (header ? part._commentBeforeSet : part._commentAfterSet) return header ? part._commentBefore : part._commentAfter;
        bool found = header ? TryHeaderRegion(part, out int start, out int end) : TryFooterRegion(part, out start, out end);
        return found ? Comments(part._document._source, start, end) : null;
    }

    private static int PartIndex(YamlDocumentPart part) {
        IReadOnlyList<YamlDocumentPart> parts = part._document.Documents;
        for (int i = 0; i < parts.Count; i++) {
            if (ReferenceEquals(parts[i], part)) return i;
        }
        return -1;
    }

    // ------------------------------------------------------------------ text helpers

    /// <summary>Finds the end of the <c>:</c> after a key, skipping blanks, line breaks and comments.</summary>
    public static int ColonEnd(string source, int keyEnd) {
        int i = keyEnd;
        while (i < source.Length) {
            char ch = source[i];
            if (ch == ':') return i + 1;
            if (ch == '#') {
                i = LineContentEnd(source, i);
                continue;
            }
            if (ch != ' ' && ch != '\t' && !IsBreak(ch)) return keyEnd;
            i++;
        }
        return keyEnd;
    }

    /// <summary>Joins the comment lines in a range, without <c>#</c> and one following space; null when there are none.</summary>
    public static string? Comments(string source, int start, int end) {
        StringBuilder? text = null;
        int line = start;
        while (line < end) {
            int lineEnd = LineContentEnd(source, line);
            if (lineEnd > end) lineEnd = end;
            int i = line;
            while (i < lineEnd && (source[i] == ' ' || source[i] == '\t')) i++;
            if (i < lineEnd && source[i] == '#') {
                if (text is null) text = new StringBuilder();
                else text.Append('\n');
                text.Append(CommentText(source, i, lineEnd));
            }
            line = AfterBreak(source, lineEnd);
            if (line == lineEnd) break;
        }
        return text?.ToString();
    }

    /// <summary>Returns a comment's text without its <c>#</c> and one following space.</summary>
    public static string CommentText(string source, int hash, int lineEnd) {
        int start = hash + 1;
        if (start < lineEnd && source[start] == ' ') start++;
        int end = lineEnd;
        if (end > start && source[end - 1] == '\r') end--;
        return source.Substring(start, end - start);
    }

    public static bool HasBlankLine(string source, int start, int end) {
        int line = start;
        while (line < end) {
            int lineEnd = LineContentEnd(source, line);
            if (IsBlankLine(source, line, lineEnd)) return true;
            int next = AfterBreak(source, lineEnd);
            if (next == lineEnd) break;
            line = next;
        }
        return false;
    }

    public static bool IsBreak(char ch) => ch == '\r' || ch == '\n';

    public static int LineStart(string source, int position) {
        while (position > 0 && !IsBreak(source[position - 1])) position--;
        return position;
    }

    /// <summary>The position of the line break (or the end of the text) on the line holding <paramref name="position"/>.</summary>
    public static int LineContentEnd(string source, int position) {
        while (position < source.Length && !IsBreak(source[position])) position++;
        return position;
    }

    /// <summary>Skips one line break at <paramref name="position"/>: <c>\r\n</c>, <c>\n</c> or <c>\r</c>.</summary>
    public static int AfterBreak(string source, int position) {
        if (position < source.Length && source[position] == '\r') {
            position++;
            if (position < source.Length && source[position] == '\n') position++;
        } else if (position < source.Length && source[position] == '\n') {
            position++;
        }
        return position;
    }

    public static int NextLineStart(string source, int position) => AfterBreak(source, LineContentEnd(source, position));

    public static bool AtLineStart(string source, int position) => position == 0 || IsBreak(source[position - 1]);

    public static bool IsBlankLine(string source, int start, int end) {
        for (int i = start; i < end; i++) {
            if (source[i] != ' ' && source[i] != '\t') return false;
        }
        return true;
    }

    public static bool IsCommentLine(string source, int start, int end) {
        int i = start;
        while (i < end && (source[i] == ' ' || source[i] == '\t')) i++;
        return i < end && source[i] == '#';
    }

    /// <summary>Whether only blanks and the block indicators <c>-</c> and <c>?</c> precede a position on its line.</summary>
    public static bool OnlyBlanksOrIndicators(string source, int lineStart, int position) {
        for (int i = lineStart; i < position; i++) {
            char ch = source[i];
            if (ch != ' ' && ch != '\t' && ch != '-' && ch != '?') return false;
        }
        return true;
    }
}
