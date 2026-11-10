namespace OfficeIMO.Yaml;

/// <summary>
/// Turns an edited document into output segments: ranges of the retained source for everything untouched, and text the
/// emitter writes for edited nodes, inserted entries, removed entries' gaps and changed comments. The walk descends only
/// into nodes that contain an edit, so its cost follows the edits, apart from copying the source ranges.
/// </summary>
/// <remarks>
/// Each edit owns a precise region of the source:
/// <list type="bullet">
/// <item><description>A changed scalar, alias or flow collection that stays on its line owns its own span, properties
/// included; its trailing comment stays.</description></item>
/// <item><description>A mapping value that changes between inline and block layout owns the text from the key's
/// <c>:</c> to the end of its last line; the comment on the key's line is written again after the new value or on the key
/// line.</description></item>
/// <item><description>A removed block entry owns its comment block, its lines and its trailing comment; the blank lines
/// above its comments stay. A removed first entry that shares its line with a dash owns the text up to the next entry,
/// which moves onto the dash line.</description></item>
/// <item><description>An inserted block entry goes above the comment block of the entry it precedes, or after the last
/// entry's last line. A removed flow item owns the separator before it, or after it when it is first.</description></item>
/// <item><description>A collection left with none of its loaded children is written whole in place of its own span.</description></item>
/// </list>
/// </remarks>
internal sealed class YamlSaveWalk {
    private readonly YamlDocument _document;
    private readonly string _source;
    private readonly List<YamlTextSegment> _segments = new List<YamlTextSegment>();
    private readonly YamlWriterOptions _options;
    private readonly string _newLine;
    private readonly StringWriter _text = new StringWriter(CultureInfo.InvariantCulture);
    private int _cursor;

    public YamlSaveWalk(YamlDocument document) {
        _document = document;
        _source = document._source;
        _options = document.WriterOptions;
        _newLine = _options.NewLine;
    }

    public List<YamlTextSegment> Run() {
        IReadOnlyList<YamlDocumentPart> parts = _document.Documents;
        for (int i = 0; i < parts.Count; i++) {
            YamlDocumentPart part = parts[i];
            if (part._new) {
                WriteNewPart(part, i);
                continue;
            }
            if (part._commentBeforeSet) EmitPartComments(part, header: true);
            if (part._root is not null) Visit(part._root);
            if (part._commentAfterSet) EmitPartComments(part, header: false);
        }
        Copy(_source.Length);
        return _segments;
    }

    // ------------------------------------------------------------------ nodes

    private void Visit(YamlNode node) {
        YamlNodeState state = node._state;
        bool trailingHandled = false;
        if ((state & YamlNodeState.CommentsDirty) != 0 && node._properties is { } edits && (edits.CommentBeforeSet || edits.SpaceBeforeSet)) {
            EmitBefore(node, edits);
        }
        if ((state & YamlNodeState.Dirty) != 0) {
            trailingHandled = EmitReplacement(node);
        } else {
            if ((state & YamlNodeState.PropertiesDirty) != 0) EmitProperties(node);
            if ((state & (YamlNodeState.ContainsDirty | YamlNodeState.Restructured)) != 0) VisitChildren(node);
        }
        if (!trailingHandled && (state & YamlNodeState.CommentsDirty) != 0 && node._properties is { CommentSet: true } trailing) {
            EmitTrailing(node, trailing.Comment);
        }
    }

    private void VisitChildren(YamlNode node) {
        if (node is YamlMapping mapping) {
            if (mapping._original is not null) {
                Restructured(mapping);
                return;
            }
            for (int i = 0; i < mapping._count; i++) {
                Visit(mapping._entries[i].Key);
                Visit(mapping._entries[i].Value);
            }
        } else if (node is YamlSequence sequence) {
            if (sequence._original is not null) {
                Restructured(sequence);
                return;
            }
            for (int i = 0; i < sequence._count; i++) Visit(sequence._items[i]);
        }
    }

    // ------------------------------------------------------------------ inserted and removed children

    private void Restructured(YamlNode collection) {
        int[] spans = YamlTrivia.Spans(collection)!;
        int originals = spans.Length / 2;
        int count = ChildCount(collection);
        // Match the children still loaded from the source with their original positions; order never changes.
        var match = new int[count];
        int survivors = 0, next = 0;
        for (int j = 0; j < count; j++) {
            match[j] = -1;
            int start = ChildStart(collection, j);
            if (start < 0) continue;
            for (int i = next; i < originals; i++) {
                if (spans[2 * i] != start) continue;
                match[j] = i;
                next = i + 1;
                survivors++;
                break;
            }
        }
        if (survivors == 0) {
            // Nothing of the loaded collection is left: write it whole.
            EmitReplacement(collection);
            return;
        }
        if (collection._style == YamlStyle.Flow) {
            // A single pair inside a flow sequence (`[a: b]`) holds one entry only; with more it needs braces.
            int content = FlowContentStart(collection);
            if (collection is YamlMapping && (_source[content] != '{' || content >= spans[0])) {
                EmitReplacement(collection);
                return;
            }
            RestructuredFlow(collection, spans, match);
            return;
        }
        bool mapping = collection is YamlMapping;
        int indent = Column(LinePosition(collection, spans[0]));
        int current = 0;
        for (int i = 0; i < originals; i++) {
            int j = IndexOfOriginal(match, i, current);
            if (j < 0) {
                RemoveBlockChild(collection, spans, i, survivorFollows: HasLaterSurvivor(match, i));
                continue;
            }
            if (j > current) InsertBlockChildren(collection, current, j, InsertionPoint(collection, spans, i, out bool inline), indent, inline);
            VisitChild(collection, j);
            current = j + 1;
        }
        if (current < count) {
            int point = OwnedEnd(spans[2 * originals - 1]);
            InsertBlockChildren(collection, current, count, point, indent, inline: false);
        }
    }

    private static int IndexOfOriginal(int[] match, int original, int from) {
        for (int j = from; j < match.Length; j++) {
            if (match[j] == original) return j;
        }
        return -1;
    }

    private static bool HasLaterSurvivor(int[] match, int original) {
        foreach (int value in match) {
            if (value > original) return true;
        }
        return false;
    }

    private void RemoveBlockChild(YamlNode collection, int[] spans, int index, bool survivorFollows) {
        int linePosition = LinePosition(collection, spans[2 * index]);
        int lineStart = YamlTrivia.LineStart(_source, linePosition);
        int end = OwnedEnd(spans[2 * index + 1]);
        if (!YamlTrivia.IsBlankLine(_source, lineStart, linePosition)) {
            // The entry shares its line with a dash: the next entry moves up onto that line.
            int start = linePosition;
            if (survivorFollows) {
                while (end < _source.Length && (_source[end] == ' ' || _source[end] == '\t')) end++;
            }
            Skip(start, end);
            return;
        }
        int gap = index > 0 ? spans[2 * index - 1] : CollectionGap(collection, out bool documentStart);
        bool atDocumentStart = index == 0 && IsDocumentStart(collection);
        // Blank lines above the entry go with it, except below a document header, which needs them to stay a header.
        YamlTrivia.RegionAt(_source, gap, atDocumentStart, linePosition, out int blankStart, out int commentStart, out _);
        Skip(atDocumentStart ? commentStart : blankStart, end);
    }

    private int InsertionPoint(YamlNode collection, int[] spans, int index, out bool inline) {
        int linePosition = LinePosition(collection, spans[2 * index]);
        int lineStart = YamlTrivia.LineStart(_source, linePosition);
        inline = !YamlTrivia.IsBlankLine(_source, lineStart, linePosition);
        if (inline) return linePosition;
        int gap = index > 0 ? spans[2 * index - 1] : CollectionGap(collection, out _);
        bool atDocumentStart = index == 0 && IsDocumentStart(collection);
        // New entries go above the next entry's blank lines and comments, which stay with that entry.
        YamlTrivia.RegionAt(_source, gap, atDocumentStart, linePosition, out int blankStart, out int commentStart, out _);
        return atDocumentStart ? commentStart : blankStart;
    }

    private void InsertBlockChildren(YamlNode collection, int from, int to, int point, int indent, bool inline) {
        Copy(point);
        bool needBreak = point > 0 && !YamlTrivia.AtLineStart(_source, point) && !inline;
        YamlEmitter emitter = Emitter();
        if (needBreak) _text.Write(_newLine);
        emitter.BeginEntries(collection is YamlMapping, indent, inline ? indent : 0);
        WriteChildren(emitter, collection, from, to);
        if (inline) {
            emitter.EndFragment(endLine: false);
            emitter.ContinueLine(indent);
        } else {
            emitter.EndFragment(endLine: true);
        }
        Flush();
    }

    private void RestructuredFlow(YamlNode collection, int[] spans, int[] match) {
        int originals = spans.Length / 2;
        int current = 0;
        int lastKept = -1;
        int i = 0;
        while (i < originals) {
            int j = IndexOfOriginal(match, i, current);
            if (j < 0) {
                if (lastKept >= 0) {
                    // Remove the separator before the item together with the item.
                    Skip(spans[2 * i - 1], spans[2 * i + 1]);
                    i++;
                    continue;
                }
                // Removed items at the head take the separator after them, up to the next kept item.
                int k = i;
                while (k < originals && IndexOfOriginal(match, k, current) < 0) k++;
                Skip(spans[2 * i], spans[2 * k]);
                i = k;
                continue;
            }
            if (j > current) {
                Copy(spans[2 * i]);
                for (int n = current; n < j; n++) {
                    Write(RenderFlowChild(collection, n));
                    Write(", ");
                }
            }
            VisitChild(collection, j);
            current = j + 1;
            lastKept = i;
            i++;
        }
        int count = ChildCount(collection);
        if (current < count) {
            int end = spans[2 * lastKept + 1];
            Copy(end);
            bool lines = lastKept > 0 && YamlTrivia.LineStart(_source, spans[2 * lastKept]) != YamlTrivia.LineStart(_source, spans[2 * lastKept - 1]);
            string separator = lines ? "," + _newLine + new string(' ', Column(spans[2 * lastKept])) : ", ";
            // An empty node right after '?' needs a space, or the '?' would read as text.
            if (end > 0 && _source[end - 1] == '?') Write(" ");
            for (int n = current; n < count; n++) {
                Write(separator);
                Write(RenderFlowChild(collection, n));
            }
        }
    }

    private string RenderFlowChild(YamlNode collection, int index) {
        if (collection is YamlMapping mapping) {
            YamlNode key = mapping._entries[index].Key, value = mapping._entries[index].Value;
            string keyText = RenderSlot(key, YamlSlot.Flow, 0, 0);
            if (value is YamlScalar { _style: YamlStyle.Plain } scalar && scalar.Value.Length == 0 && scalar._properties is null) return keyText;
            return keyText + (key is YamlAlias ? " : " : ": ") + RenderSlot(value, YamlSlot.Flow, 0, 0);
        }
        return RenderSlot(((YamlSequence)collection)._items[index], YamlSlot.Flow, 0, 0);
    }

    // ------------------------------------------------------------------ replacements

    /// <summary>Writes an edited node in place of its region. Returns whether its trailing comment was written with it.</summary>
    private bool EmitReplacement(YamlNode node) {
        int start = node._start, end = node._end;
        if (start < 0) return false;
        YamlNode? parent = node._parent;
        bool endsAtLineStart = end > start && YamlTrivia.IsBreak(_source[end - 1]);
        if (parent is YamlMapping mapping) {
            int index = mapping.IndexOfNode(node, out bool isKey);
            if (isKey) {
                string keyText = RenderSlot(node, parent._style == YamlStyle.Flow ? YamlSlot.Flow : YamlSlot.Inline, 0, Column(start));
                // A block scalar key ended with its line break; the ':' that followed stays on the next line.
                if (endsAtLineStart && !keyText.EndsWith("\n", StringComparison.Ordinal)) keyText += _newLine;
                Replace(start, end, keyText);
                return false;
            }
            YamlNode key = mapping._entries[index].Key;
            int colonEnd = ColonEnd(key._end);
            bool hasColon = colonEnd > key._end;
            if (parent._style == YamlStyle.Flow) {
                // An omitted value may have no ':' at all (`{a}`), or no space after it.
                string flow = RenderSlot(node, YamlSlot.Flow, 0, Column(start));
                if (!hasColon) {
                    // A key that spans lines cannot be implicit once it has a ':'; it becomes explicit.
                    if (ContainsBreak(key._start, key._end) && _cursor <= key._start && YamlTrivia.EntryStart(_source, key._start, 0) == key._start && !PrecededByQuestion(key._start)) Insert(key._start, "? ");
                    flow = (key is YamlAlias || (start > 0 && _source[start - 1] == '?') ? " : " : ": ") + flow;
                }
                else if (start == colonEnd) flow = " " + flow;
                Replace(start, end, flow);
                return false;
            }
            int entry = YamlTrivia.EntryStart(_source, key._start, YamlTrivia.MappingStart(mapping));
            int question = entry < key._start ? entry : -1;
            if (question >= 0 || key is YamlMapping or YamlSequence) {
                // An explicit entry: the value follows a ':' at the '?' column and is laid out like a sequence item.
                int column = question >= 0 ? Column(question) : Column(YamlTrivia.LineStart(_source, key._start));
                if (!hasColon) {
                    // Add the ':' line below the key, leaving the key's line and its comment as they are.
                    int point = YamlTrivia.LineContentEnd(_source, start);
                    return EmitAfterIndicator(node, point, point, column, column + 2, endsAtLineStart: false, _newLine + new string(' ', column) + ": ");
                }
                return EmitAfterIndicator(node, start, end, Column(colonEnd - 1), Column(start) + (start == colonEnd ? 1 : 0), endsAtLineStart, start == colonEnd ? " " : string.Empty);
            }
            int keyIndent = Column(key._start);
            bool oldInline = !ContainsBreak(colonEnd, start);
            string inline = RenderSlot(node, start == colonEnd ? YamlSlot.MappingValue : YamlSlot.Inline, keyIndent, Column(start));
            if (oldInline && !endsAtLineStart && !ContainsBreak(inline)) {
                Replace(start, end, inline);
                return false;
            }
            if (!oldInline && !endsAtLineStart && start > colonEnd && node is YamlScalar or YamlAlias) {
                // A value on its own line below the key stays there when it fits on one line.
                string below = RenderSlot(node, YamlSlot.Inline, keyIndent, Column(start));
                if (!ContainsBreak(below)) {
                    Replace(start, end, below);
                    return false;
                }
            }
            // The value changes between inline and block layout: rewrite from the colon, keeping the line's comment.
            int regionEnd = endsAtLineStart ? end : YamlTrivia.LineContentEnd(_source, end);
            int hash = FindHash(colonEnd, YamlTrivia.LineContentEnd(_source, colonEnd));
            string? comment = Edited(node, out string? edited) ? edited : hash >= 0 ? YamlTrivia.CommentText(_source, hash, YamlTrivia.LineContentEnd(_source, hash)) : null;
            string text = RenderSlot(node, YamlSlot.MappingValue, keyIndent, Column(colonEnd), comment, endLine: endsAtLineStart);
            Replace(colonEnd, endsAtLineStart ? end : Math.Max(regionEnd, end), text);
            return true;
        }
        if (parent is YamlSequence { _style: YamlStyle.Flow }) {
            Replace(start, end, RenderSlot(node, YamlSlot.Flow, 0, Column(start)));
            return false;
        }
        // A sequence item or a document root.
        int parentIndent = parent is YamlSequence ? Column(YamlTrivia.DashPosition(_source, start)) : 0;
        bool space = NeedsSpace(start);
        return EmitAfterIndicator(node, start, end, parentIndent, Column(start) + (space ? 1 : 0), endsAtLineStart, space ? " " : string.Empty);
    }

    /// <summary>
    /// Writes a node that follows an indicator (<c>-</c>, an explicit <c>:</c>) or starts a document: it stays in its span
    /// when it fits on one line, and otherwise takes its line's trailing comment along to its first line.
    /// </summary>
    private bool EmitAfterIndicator(YamlNode node, int start, int end, int parentIndent, int column, bool endsAtLineStart, string prefix) {
        if (prefix.Length == 0 && node is YamlMapping or YamlSequence && node._style == YamlStyle.Block) {
            // A tab may separate an indicator from a scalar, but cannot indent a compact block collection.
            int blanks = start;
            while (blanks > 0 && (_source[blanks - 1] == ' ' || _source[blanks - 1] == '\t')) blanks--;
            if (_source.IndexOf('\t', blanks, start - blanks) >= 0) {
                start = blanks;
                prefix = " ";
                column = Column(blanks) + 1;
            }
        }
        string single = RenderSlot(node, YamlSlot.Inline, parentIndent, column);
        if (!endsAtLineStart && !ContainsBreak(single)) {
            Replace(start, end, prefix + single);
            return false;
        }
        // A block collection has no line of its own: the replaced node's trailing comment goes with it.
        bool block = node is YamlMapping or YamlSequence && node._style == YamlStyle.Block;
        int lineEnd = endsAtLineStart ? end : YamlTrivia.LineContentEnd(_source, end);
        int trailing = endsAtLineStart || block ? -1 : FindHash(end, lineEnd);
        string? note = block ? null : Edited(node, out string? editedNote) ? editedNote : trailing >= 0 ? YamlTrivia.CommentText(_source, trailing, lineEnd) : null;
        string multi = RenderSlot(node, YamlSlot.Inline, parentIndent, column, note, endLine: endsAtLineStart);
        Replace(start, endsAtLineStart ? end : lineEnd, prefix + multi);
        return true;
    }

    private bool PrecededByQuestion(int start) {
        int i = start - 1;
        while (i >= 0 && (_source[i] == ' ' || _source[i] == '\t' || YamlTrivia.IsBreak(_source[i]))) i--;
        return i >= 0 && _source[i] == '?';
    }

    private bool NeedsSpace(int start) => start > 0 && _source[start - 1] != ' ' && _source[start - 1] != '\t' && !YamlTrivia.IsBreak(_source[start - 1]);

    private static bool Edited(YamlNode node, out string? comment) {
        comment = null;
        if (node._properties is not { CommentSet: true } edits) return false;
        comment = edits.Comment;
        return true;
    }

    /// <summary>Rewrites the anchor and tag of a node whose content is still copied from the source.</summary>
    private void EmitProperties(YamlNode node) {
        string properties = PropertiesText(node);
        int start = node._start;
        char first = _source[start];
        bool had = first == '&' || first == '!';
        bool block = node is YamlMapping or YamlSequence && node._style == YamlStyle.Block;
        if (had) {
            int end = start;
            while (end < _source.Length && (_source[end] == '&' || _source[end] == '!')) {
                while (end < _source.Length && _source[end] != ' ' && _source[end] != '\t' && !YamlTrivia.IsBreak(_source[end])) end++;
                int afterBlanks = end;
                while (afterBlanks < _source.Length && (_source[afterBlanks] == ' ' || _source[afterBlanks] == '\t')) afterBlanks++;
                if (afterBlanks < _source.Length && (_source[afterBlanks] == '&' || _source[afterBlanks] == '!')) end = afterBlanks;
            }
            if (properties.Length > 0) {
                Replace(start, end, properties);
            } else if (block) {
                int from = start;
                while (from > 0 && (_source[from - 1] == ' ' || _source[from - 1] == '\t')) from--;
                Skip(from, end);
            } else {
                while (end < _source.Length && (_source[end] == ' ' || _source[end] == '\t')) end++;
                Skip(start, end);
            }
            return;
        }
        if (properties.Length == 0) return;
        if (!block) {
            Insert(start, properties + " ");
            return;
        }
        if (node._parent is YamlMapping mapping) {
            int index = mapping.IndexOfNode(node, out _);
            Insert(ColonEnd(mapping._entries[index].Key._end), " " + properties);
            return;
        }
        Insert(start, properties + _newLine + new string(' ', Column(start)));
    }

    private static string PropertiesText(YamlNode node) {
        string text = string.Empty;
        if (node.Anchor is { } anchor) text = "&" + anchor;
        if (node.Tag is { } tag) text += (text.Length > 0 ? " " : string.Empty) + YamlScalarFormat.FormatTag(tag, node._part?.RedefinesSecondaryHandle ?? false);
        return text;
    }

    // ------------------------------------------------------------------ comments

    private void EmitBefore(YamlNode owner, YamlNodeProperties edits) {
        if (!owner.HasSourceSpan && owner._start < 0) return;
        if (!YamlTrivia.TryBeforeRegion(owner, out int blankStart, out int commentStart, out int lineStart)) return;
        int indent = Column(YamlTrivia.LinePosition(owner));
        if (edits.SpaceBeforeSet) {
            bool has = YamlTrivia.HasBlankLine(_source, blankStart, commentStart);
            if (edits.SpaceBefore && !has) Insert(Math.Max(blankStart, _cursor), _newLine);
            else if (!edits.SpaceBefore && has) Skip(Math.Max(blankStart, _cursor), commentStart);
        }
        if (edits.CommentBeforeSet) Replace(Math.Max(commentStart, _cursor), lineStart, CommentLines(edits.CommentBefore, indent));
    }

    private void EmitTrailing(YamlNode node, string? text) {
        if (!node.HasSourceSpan) return;
        int hash = YamlTrivia.TrailingHash(node, out int lineEnd);
        if (hash >= 0) {
            if (text is null) {
                int from = hash;
                while (from > 0 && (_source[from - 1] == ' ' || _source[from - 1] == '\t')) from--;
                Skip(Math.Max(from, _cursor), lineEnd);
            } else {
                Replace(hash, lineEnd, CommentText(text));
            }
            return;
        }
        if (text is null) return;
        int position;
        if (node is YamlScalar { Style: YamlStyle.Literal or YamlStyle.Folded } block) position = YamlTrivia.LineContentEnd(_source, YamlTrivia.BlockHeaderPosition(_source, block));
        else if (YamlEdit.IsKey(node, out _)) position = YamlTrivia.LineContentEnd(_source, node._end);
        else position = node._end;
        // Blanks at the end of the line stay before the new comment instead of ending up inside it.
        int contentEnd = YamlTrivia.LineContentEnd(_source, position);
        if (YamlTrivia.IsBlankLine(_source, position, contentEnd)) position = contentEnd;
        Insert(Math.Max(position, _cursor), " " + CommentText(text));
    }

    private void EmitPartComments(YamlDocumentPart part, bool header) {
        string? text = header ? part._commentBefore : part._commentAfter;
        if (header) {
            if (!YamlTrivia.TryHeaderRegion(part, out int start, out int end)) return;
            bool had = YamlTrivia.Comments(_source, start, end) is not null;
            string lines = CommentLines(text, 0);
            if (!part.HasStartMarker && !had && text is not null) lines += _newLine;
            if (!part.HasStartMarker && had && text is null) {
                // Drop the header and the blank lines that set it apart.
                Skip(Math.Max(start, _cursor), end);
                return;
            }
            Replace(Math.Max(start, _cursor), end, lines);
            return;
        }
        if (!YamlTrivia.TryFooterRegion(part, out int from, out int to)) return;
        int first = FirstCommentLine(from, to);
        string footer = CommentLines(text, 0);
        if (first < 0) {
            if (text is null) return;
            int point = Math.Max(from, _cursor);
            if (point > 0 && !YamlTrivia.AtLineStart(_source, point)) footer = _newLine + footer;
            Insert(point, footer);
            return;
        }
        Replace(Math.Max(first, _cursor), LastCommentLineEnd(first, to), footer);
    }

    private int FirstCommentLine(int from, int to) {
        int line = from;
        while (line < to) {
            int end = YamlTrivia.LineContentEnd(_source, line);
            if (YamlTrivia.IsCommentLine(_source, line, end)) return line;
            int next = YamlTrivia.AfterBreak(_source, end);
            if (next == line) break;
            line = next;
        }
        return -1;
    }

    private int LastCommentLineEnd(int from, int to) {
        int result = from, line = from;
        while (line < to) {
            int end = YamlTrivia.LineContentEnd(_source, line);
            int next = YamlTrivia.AfterBreak(_source, end);
            if (YamlTrivia.IsCommentLine(_source, line, end)) result = next;
            if (next == line) break;
            line = next;
        }
        return result;
    }

    private string CommentLines(string? text, int indent) {
        if (text is null) return string.Empty;
        var lines = new StringBuilder();
        string pad = new string(' ', indent);
        foreach (string line in text.Split('\n')) {
            string content = line.TrimEnd('\r');
            lines.Append(pad).Append('#');
            if (content.Length > 0) lines.Append(' ').Append(content);
            lines.Append(_newLine);
        }
        return lines.ToString();
    }

    private static string CommentText(string text) => text.Length == 0 ? "#" : "# " + text;

    // ------------------------------------------------------------------ new parts

    private void WriteNewPart(YamlDocumentPart part, int index) {
        Copy(_source.Length);
        YamlEmitter emitter = Emitter();
        if (_source.Length > 0 && !YamlTrivia.AtLineStart(_source, _source.Length) && _segments.Count > 0) _text.Write(_newLine);
        bool marker = part.HasStartMarker || index > 0;
        YamlModelWriter.WritePart(emitter, part, marker && (index > 0 || part.HasStartMarker));
        emitter.Finish();
        Flush();
    }

    // ------------------------------------------------------------------ helpers

    private static int ChildCount(YamlNode collection) => collection is YamlMapping m ? m._count : ((YamlSequence)collection)._count;

    private static int ChildStart(YamlNode collection, int index) =>
        collection is YamlMapping m ? (m._entries[index].Key._start >= 0 && m._entries[index].Key._part is not null ? m._entries[index].Key._start : -1)
        : ((YamlSequence)collection)._items[index]._start;

    private void VisitChild(YamlNode collection, int index) {
        if (collection is YamlMapping mapping) {
            Visit(mapping._entries[index].Key);
            Visit(mapping._entries[index].Value);
        } else {
            Visit(((YamlSequence)collection)._items[index]);
        }
    }

    private static void WriteChildren(YamlEmitter emitter, YamlNode collection, int from, int to) {
        for (int i = from; i < to; i++) {
            if (collection is YamlMapping mapping) {
                YamlNode key = mapping._entries[i].Key, value = mapping._entries[i].Value;
                YamlModelWriter.CommentLines(emitter, key);
                YamlModelWriter.WriteNode(emitter, key, ownComments: false);
                if (value is YamlMapping { _count: > 0, _style: YamlStyle.Block } || value is YamlSequence { _count: > 0, _style: YamlStyle.Block }) {
                    if (key.Comment is { } note) emitter.SetTrailingComment(note);
                }
                YamlModelWriter.WriteNode(emitter, value, ownComments: true);
            } else {
                YamlNode item = ((YamlSequence)collection)._items[i];
                YamlModelWriter.CommentLines(emitter, item);
                YamlModelWriter.WriteNode(emitter, item, ownComments: true);
            }
        }
    }

    /// <summary>The position deciding a loaded child's line: the dash of a block item, or the key's start.</summary>
    private int LinePosition(YamlNode collection, int childStart) =>
        collection._style != YamlStyle.Block ? childStart
        : collection is YamlSequence ? YamlTrivia.DashPosition(_source, childStart) : YamlTrivia.EntryStart(_source, childStart, YamlTrivia.MappingStart((YamlMapping)collection));

    private static int CollectionGap(YamlNode collection, out bool documentStart) => YamlTrivia.CollectionGapStart(collection, out documentStart);

    private bool IsDocumentStart(YamlNode collection) {
        CollectionGap(collection, out bool documentStart);
        return documentStart;
    }

    /// <summary>The end of a child's last line: after its trailing comment and line break.</summary>
    private int OwnedEnd(int end) {
        if (end > 0 && YamlTrivia.IsBreak(_source[end - 1])) return end;
        int position = end;
        while (position < _source.Length && (_source[position] == ' ' || _source[position] == '\t')) position++;
        if (position < _source.Length && _source[position] == '#') position = YamlTrivia.LineContentEnd(_source, position);
        if (position < _source.Length && YamlTrivia.IsBreak(_source[position])) return YamlTrivia.AfterBreak(_source, position);
        return position >= _source.Length ? _source.Length : end;
    }

    /// <summary>Finds the end of the <c>:</c> after a key, skipping blanks, line breaks and comments.</summary>
    private int ColonEnd(int keyEnd) => YamlTrivia.ColonEnd(_source, keyEnd);

    /// <summary>The first character of a flow collection after its anchor and tag.</summary>
    private int FlowContentStart(YamlNode collection) {
        int i = collection._start;
        while (i < _source.Length && (_source[i] == '&' || _source[i] == '!')) {
            while (i < _source.Length && _source[i] != ' ' && _source[i] != '\t' && !YamlTrivia.IsBreak(_source[i])) i++;
            while (i < _source.Length && (_source[i] == ' ' || _source[i] == '\t' || YamlTrivia.IsBreak(_source[i]))) i++;
        }
        return Math.Min(i, _source.Length - 1);
    }

    private int FindHash(int from, int end) {
        for (int i = from; i < end; i++) {
            if (_source[i] == '#' && i > 0 && (_source[i - 1] == ' ' || _source[i - 1] == '\t')) return i;
        }
        return -1;
    }

    private bool ContainsBreak(int from, int to) {
        for (int i = from; i < to; i++) {
            if (YamlTrivia.IsBreak(_source[i])) return true;
        }
        return false;
    }

    private static bool ContainsBreak(string text) => text.IndexOf('\n') >= 0 || text.IndexOf('\r') >= 0;

    private int Column(int position) => position - YamlTrivia.LineStart(_source, position);

    private string RenderSlot(YamlNode node, YamlSlot slot, int parentIndent, int column, string? comment = null, bool endLine = false) {
        _text.GetStringBuilder().Clear();
        var emitter = new YamlEmitter(_text, _options);
        emitter.BeginSlot(slot, parentIndent, column);
        if (comment is not null && slot != YamlSlot.Flow && IsBlockScalar(node)) emitter.SetTrailingComment(comment);
        YamlModelWriter.WriteNode(emitter, node, ownComments: false);
        if (comment is not null && slot != YamlSlot.Flow && !IsBlockScalar(node)) {
            // On the first line: right after an inline value, or on the key line before a block collection.
            if (node is YamlMapping or YamlSequence && node._style == YamlStyle.Block && ChildCount(node) > 0) { } else emitter.SetTrailingComment(comment);
        }
        emitter.EndFragment(endLine);
        string text = _text.ToString();
        _text.GetStringBuilder().Clear();
        if (comment is not null && node is YamlMapping or YamlSequence && node._style == YamlStyle.Block && ChildCount(node) > 0) {
            text = PlaceOnFirstLine(text, comment);
        }
        return endLine ? text : YamlModelWriter.TrimFinalBreak(text);
    }

    private static bool IsBlockScalar(YamlNode node) => node is YamlScalar { _style: YamlStyle.Literal or YamlStyle.Folded };

    /// <summary>Puts a comment at the end of a rendered block's first line, the key line of a mapping value.</summary>
    private string PlaceOnFirstLine(string text, string comment) {
        int lineEnd = text.IndexOf('\n');
        if (lineEnd < 0) return text + " " + CommentText(comment);
        int contentEnd = lineEnd > 0 && text[lineEnd - 1] == '\r' ? lineEnd - 1 : lineEnd;
        return text.Substring(0, contentEnd) + " " + CommentText(comment) + text.Substring(contentEnd);
    }

    private YamlEmitter Emitter() {
        _text.GetStringBuilder().Clear();
        return new YamlEmitter(_text, _options);
    }

    /// <summary>Adds the emitter's buffered text as a segment.</summary>
    private void Flush() {
        string text = _text.ToString();
        _text.GetStringBuilder().Clear();
        if (text.Length > 0) _segments.Add(new YamlTextSegment(text, 0, text.Length));
    }

    private void Copy(int to) {
        if (to > _cursor) {
            _segments.Add(new YamlTextSegment(_source, _cursor, to - _cursor));
            _cursor = to;
        }
    }

    private void Write(string text) {
        if (text.Length > 0) _segments.Add(new YamlTextSegment(text, 0, text.Length));
    }

    private void Insert(int position, string text) {
        Copy(position);
        Write(text);
    }

    private void Replace(int start, int end, string text) {
        Copy(start);
        Write(text);
        if (end > _cursor) _cursor = end;
    }

    private void Skip(int start, int end) {
        Copy(start);
        if (end > _cursor) _cursor = end;
    }
}
