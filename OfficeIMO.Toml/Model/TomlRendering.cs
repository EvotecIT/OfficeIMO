namespace OfficeIMO.Toml;

/// <summary>A run of output text: a range of a string, usually of the retained source.</summary>
internal readonly struct TomlTextSegment {
    public TomlTextSegment(string text, int start, int length) {
        Text = text;
        Start = start;
        Length = length;
    }

    public string Text { get; }
    public int Start { get; }
    public int Length { get; }
}

/// <summary>
/// Produces the saved text of a document. Unchanged text is copied from the source; the walk descends only into
/// changed nodes and turns each change into a text edit at a source position: a replaced value, a deleted line or
/// section, or inserted lines written in the document's own layout.
/// </summary>
internal sealed class TomlRendering {
    private readonly TomlDocument _document;
    private readonly string _source;
    private readonly TomlText _text;
    private readonly List<Edit> _edits = new List<Edit>();

    private TomlRendering(TomlDocument document) {
        _document = document;
        _source = document._source;
        _text = new TomlText(_source);
    }

    private readonly struct Edit {
        public Edit(int start, int end, string text, int sequence) {
            Start = start;
            End = end;
            Text = text;
            Sequence = sequence;
        }

        public int Start { get; }
        public int End { get; }
        public string Text { get; }
        public int Sequence { get; }
    }

    private TomlLayout Layout => _document.Layout;

    public static List<TomlTextSegment> Collect(TomlDocument document) {
        string source = document._source;
        var segments = new List<TomlTextSegment>(1);
        if ((document._root._state & TomlNodeState.Changed) == 0) {
            if (source.Length > 0) segments.Add(new TomlTextSegment(source, 0, source.Length));
            return segments;
        }
        var rendering = new TomlRendering(document);
        rendering.Visit(document._root);
        rendering._edits.Sort((a, b) => a.Start != b.Start ? a.Start.CompareTo(b.Start) : a.End != b.End ? a.End.CompareTo(b.End) : a.Sequence.CompareTo(b.Sequence));
        int cursor = 0;
        foreach (Edit edit in rendering._edits) {
            if (edit.Start < cursor) {
                // Inside text another edit already replaced or deleted: overlapping deletions merge, the rest is gone.
                if (edit.End > cursor && edit.Text.Length == 0) cursor = edit.End;
                continue;
            }
            if (edit.Start > cursor) segments.Add(new TomlTextSegment(source, cursor, edit.Start - cursor));
            if (edit.Text.Length > 0) segments.Add(new TomlTextSegment(edit.Text, 0, edit.Text.Length));
            cursor = edit.End;
        }
        if (cursor < source.Length) {
            segments.Add(new TomlTextSegment(source, cursor, source.Length - cursor));
        } else if (segments.Count > 0 && ReferenceEquals(segments[segments.Count - 1].Text, source) && rendering.DeletedToEnd()) {
            // Text deleted up to the end of the file leaves no blank lines behind it.
            TomlTextSegment last = segments[segments.Count - 1];
            int end = rendering._text.SkipBlankLinesUp(last.Start + last.Length, last.Start);
            segments[segments.Count - 1] = new TomlTextSegment(source, last.Start, end - last.Start);
        }
        return segments;
    }

    private bool DeletedToEnd() {
        foreach (Edit edit in _edits) {
            if (edit.End == _source.Length && edit.Text.Length == 0 && edit.End > edit.Start) return true;
        }
        return false;
    }

    private void Replace(int start, int end, string text) => _edits.Add(new Edit(start, end, text, _edits.Count));

    private void Insert(int position, string text) {
        if (text.Length > 0) _edits.Add(new Edit(position, position, text, _edits.Count));
    }

    private void Delete(int start, int end) {
        if (end > start) _edits.Add(new Edit(start, end, string.Empty, _edits.Count));
    }

    private bool HasInsertAt(int position) {
        foreach (Edit edit in _edits) {
            if (edit.Start == position && edit.End == position) return true;
        }
        return false;
    }

    // ---- Walk --------------------------------------------------------------------------------------------------

    private void Visit(TomlNode node) {
        if ((node._state & TomlNodeState.Detached) != 0 || !node.IsLoaded) return;
        if ((node._state & TomlNodeState.Dirty) != 0) {
            Replace(node._start, node._end, RenderValue(node));
            if ((node._state & TomlNodeState.CommentChanged) != 0) WriteComment(node);
            return;
        }
        if ((node._state & TomlNodeState.CommentChanged) != 0) WriteComment(node);
        if ((node._state & TomlNodeState.Changed) == 0) return;
        switch (node) {
            case TomlTable table when table._tableKind == TomlTableKind.Inline:
                VisitInlineTable(table);
                break;
            case TomlTable table:
                VisitSectionTable(table);
                break;
            case TomlArray array:
                VisitArray(array);
                break;
            case TomlTableArray tables:
                VisitTableArray(tables);
                break;
        }
    }

    private void VisitSectionTable(TomlTable table) {
        if (table._removed != null) {
            foreach (TomlNode removed in table._removed) DeleteUnits(removed);
        }
        List<TomlNode>? lines = null, sections = null;
        for (int i = 0; i < table._count; i++) {
            TomlNode node = table._nodes[i];
            if (node.IsLoaded) continue;
            if (IsSection(node)) (sections ??= new List<TomlNode>()).Add(node);
            else (lines ??= new List<TomlNode>()).Add(node);
        }
        if (lines != null) InsertLines(table, lines);
        if (sections != null) InsertSections(table, sections);
        for (int i = 0; i < table._count; i++) Visit(table._nodes[i]);
    }

    private void VisitTableArray(TomlTableArray tables) {
        if (tables._removed != null) {
            foreach (TomlNode removed in tables._removed) DeleteUnits(removed);
        }
        List<TomlNode>? added = null;
        foreach (TomlTable item in tables._items) {
            if (!item.IsLoaded) (added ??= new List<TomlNode>()).Add(item);
        }
        if (added != null) {
            int anchor = -1;
            foreach (TomlTable item in tables._items) anchor = Math.Max(anchor, SubtreeEnd(item));
            if (tables._removed != null) {
                foreach (TomlNode removed in tables._removed) anchor = Math.Max(anchor, SubtreeEnd(removed));
            }
            InsertSectionText(anchor < 0 ? _source.Length : anchor, added, PathOf(tables));
        }
        foreach (TomlTable item in tables._items) Visit(item);
    }

    /// <summary>New children written as their own <c>[header]</c> sections rather than key/value lines.</summary>
    private static bool IsSection(TomlNode node) =>
        node is TomlTableArray || (node is TomlTable table && table._tableKind is TomlTableKind.Header or TomlTableKind.ArrayItem);

    // ---- Deleting ----------------------------------------------------------------------------------------------

    /// <summary>Deletes the text of a removed loaded node: its key/value lines and the sections of its tables.</summary>
    private void DeleteUnits(TomlNode node) {
        if (!node.IsLoaded) return;
        if (node.IsKeyValue) {
            int start = _text.BlockStart(_text.LineStart(node._keyStart), 0);
            int end = _text.LineEnd(node._end);
            // The last line of a file without a final line break takes the break before it, so the file still ends
            // without one and text added at the end starts on a new line of its own, not after a blank one.
            if (end == _source.Length && start > 0 && _source[end - 1] != '\n') {
                start--;
                if (start > 0 && _source[start - 1] == '\r') start--;
            }
            Delete(start, end);
            return;
        }
        switch (node) {
            case TomlTable table:
                if (table._section >= 0) {
                    (int start, int end) = SectionRange(table);
                    Delete(start, end);
                }
                for (int i = 0; i < table._count; i++) DeleteUnits(table._nodes[i]);
                if (table._removed != null) {
                    foreach (TomlNode removed in table._removed) DeleteUnits(removed);
                }
                break;
            case TomlTableArray tables:
                foreach (TomlTable item in tables._items) DeleteUnits(item);
                if (tables._removed != null) {
                    foreach (TomlNode removed in tables._removed) DeleteUnits(removed);
                }
                break;
        }
    }

    /// <summary>
    /// Returns a header's whole section: from the comment lines directly above the header to the comment block of the
    /// next header. The last section also takes the blank lines before it, so the file does not end with them.
    /// </summary>
    private (int Start, int End) SectionRange(TomlTable header) {
        int start = _text.BlockStart(_text.LineStart(header._start), 0);
        int end = NextSectionStart(header._section);
        if (end == _source.Length) start = _text.SkipBlankLinesUp(start, 0);
        return (start, end);
    }

    /// <summary>Returns where the section after the given header index starts, including its leading comments.</summary>
    private int NextSectionStart(int section) {
        List<TomlTable> headers = _document._headers;
        if (section + 1 >= headers.Count) return _source.Length;
        TomlTable next = headers[section + 1];
        return _text.BlockStart(_text.LineStart(next._start), _text.LineEnd(headers[section]._end));
    }

    /// <summary>Returns the end of a section's content: before trailing blank lines and the next header's comments.</summary>
    private int SectionContentEnd(TomlTable header) {
        int floor = _text.LineEnd(header._end);
        return Math.Max(floor, _text.SkipBlankLinesUp(NextSectionStart(header._section), floor));
    }

    private int RootContentEnd() {
        if (_document._headers.Count == 0) return _text.SkipBlankLinesUp(_source.Length, 0);
        int first = _text.BlockStart(_text.LineStart(_document._headers[0]._start), 0);
        return _text.SkipBlankLinesUp(first, 0);
    }

    /// <summary>Returns the latest section end in a loaded subtree, or -1 when it has no header sections.</summary>
    private int SubtreeEnd(TomlNode node) {
        int end = -1;
        if (node is TomlTable table) {
            if (table._section >= 0) end = SectionContentEnd(table);
            for (int i = 0; i < table._count; i++) {
                if (table._nodes[i].IsLoaded) end = Math.Max(end, SubtreeEnd(table._nodes[i]));
            }
        } else if (node is TomlTableArray tables) {
            foreach (TomlTable item in tables._items) {
                if (item.IsLoaded) end = Math.Max(end, SubtreeEnd(item));
            }
        }
        return end;
    }

    // ---- Inserting key/value lines -----------------------------------------------------------------------------

    private void InsertLines(TomlTable table, List<TomlNode> nodes) {
        if (table._tableKind == TomlTableKind.Implicit) {
            InsertImplicitHeader(table, nodes);
            return;
        }
        TomlTable owner = table;
        var prefix = new List<string>();
        while (owner._tableKind == TomlTableKind.Dotted && owner._parent is TomlTable parent) {
            prefix.Insert(0, owner._key!);
            owner = parent;
        }
        TomlNode? last = LastLine(table, owner) ?? (ReferenceEquals(table, owner) ? null : LastLine(owner, owner));
        int anchor;
        bool separateFromHeader = false;
        if (last != null) {
            anchor = _text.LineEnd(last._end);
        } else if (owner._tableKind == TomlTableKind.Root) {
            if (_document._headers.Count == 0) {
                anchor = _source.Length;
            } else {
                anchor = _text.BlockStart(_text.LineStart(_document._headers[0]._start), 0);
                separateFromHeader = Layout.BlankLineBeforeHeaders;
            }
        } else {
            anchor = _text.LineEnd(owner._end);
        }
        string indent = last != null ? _text.Indentation(last._keyStart) : owner._tableKind == TomlTableKind.Root ? string.Empty : Layout.KeyIndent;
        string separator = last != null ? _source.Substring(last.PairKeyEnd, last._start - last.PairKeyEnd) : Layout.KeySeparator;
        var text = new StringBuilder();
        StartLine(text, anchor);
        foreach (TomlNode node in nodes) AppendLine(text, indent, prefix, node, separator);
        if (separateFromHeader) text.Append(Layout.NewLine);
        Insert(anchor, text.ToString());
    }

    /// <summary>Gives an implicit table its own <c>[header]</c>, placed before its first sub-table, for new key/value lines.</summary>
    private void InsertImplicitHeader(TomlTable table, List<TomlNode> nodes) {
        int first = int.MaxValue;
        FirstHeader(table, ref first);
        int anchor = first == int.MaxValue ? _source.Length : _text.BlockStart(_text.LineStart(first), 0);
        var text = new StringBuilder();
        StartLine(text, anchor);
        AppendHeader(text, table, PathOf(table));
        foreach (TomlNode node in nodes) AppendLine(text, Layout.KeyIndent, null, node, Layout.KeySeparator);
        if (first != int.MaxValue && Layout.BlankLineBeforeHeaders) text.Append(Layout.NewLine);
        Insert(anchor, text.ToString());
    }

    private static void FirstHeader(TomlNode node, ref int first) {
        if (node is TomlTable table) {
            if (table._section >= 0) first = Math.Min(first, table._start);
            for (int i = 0; i < table._count; i++) FirstHeader(table._nodes[i], ref first);
        } else if (node is TomlTableArray tables) {
            foreach (TomlTable item in tables._items) FirstHeader(item, ref first);
        }
    }

    /// <summary>Returns the loaded key/value node in a table's subtree whose line ends last within the owner's section.</summary>
    private static TomlNode? LastLine(TomlTable table, TomlTable owner) {
        TomlNode? last = null;
        for (int i = 0; i < table._count; i++) {
            TomlNode node = table._nodes[i];
            if (node.IsKeyValue) {
                if (last == null || node._end > last._end) last = node;
            } else if (node is TomlTable dotted && dotted._tableKind == TomlTableKind.Dotted) {
                TomlNode? inner = LastLine(dotted, owner);
                if (inner != null && (last == null || inner._end > last._end)) last = inner;
            }
        }
        return last;
    }

    /// <summary>Starts inserted text on a new line when it lands at the end of text without a final line break.</summary>
    private void StartLine(StringBuilder text, int anchor) {
        if (anchor == _source.Length && anchor > 0 && _source[anchor - 1] != '\n') text.Append(Layout.NewLine);
    }

    private void AppendLine(StringBuilder text, string indent, List<string>? prefix, TomlNode node, string separator) {
        text.Append(indent);
        if (prefix != null) {
            foreach (string key in prefix) {
                TomlFormatter.AppendKey(text, key);
                text.Append('.');
            }
        }
        TomlFormatter.AppendKey(text, node._key!);
        text.Append(separator);
        AppendValue(text, node);
        AppendNewComment(text, node);
        text.Append(Layout.NewLine);
    }

    private void AppendNewComment(StringBuilder text, TomlNode node) {
        string? comment = node.ChangedComment;
        if ((node._state & TomlNodeState.CommentChanged) == 0 || comment is null) return;
        text.Append(Layout.CommentGap).Append(CommentText(comment));
    }

    // ---- Inserting sections ------------------------------------------------------------------------------------

    private void InsertSections(TomlTable table, List<TomlNode> nodes) {
        int anchor;
        if (table._tableKind == TomlTableKind.Root) {
            anchor = _text.SkipBlankLinesUp(_source.Length, 0);
        } else {
            anchor = SubtreeEnd(table);
            TomlTable owner = table;
            while (owner._tableKind == TomlTableKind.Dotted && owner._parent is TomlTable parent) owner = parent;
            if (!ReferenceEquals(owner, table) || anchor < 0) {
                int ownerEnd = owner._tableKind == TomlTableKind.Root ? RootContentEnd() : owner._section >= 0 ? SectionContentEnd(owner) : -1;
                anchor = Math.Max(anchor, ownerEnd);
            }
            if (anchor < 0) anchor = _source.Length;
        }
        List<string> path = PathOf(table);
        var items = new List<TomlNode>();
        foreach (TomlNode node in nodes) {
            if (node is TomlTableArray tables) items.AddRange(tables._items);
            else items.Add(node);
        }
        InsertSectionText(anchor, items, path, nodes);
    }

    private void InsertSectionText(int anchor, List<TomlNode> tables, List<string> parentPath, List<TomlNode>? declared = null) {
        var text = new StringBuilder();
        StartLine(text, anchor);
        bool content = anchor > 0 || HasInsertAt(anchor);
        foreach (TomlNode node in tables) {
            var path = new List<string>(parentPath);
            if (node._key != null) path.Add(node._key);
            else if (node._parent is TomlTableArray owner && owner._key != null && declared != null) path.Add(owner._key);
            AppendSection(text, (TomlTable)node, path, ref content);
        }
        Insert(anchor, text.ToString());
    }

    /// <summary>Writes a new table as a header section followed by its sub-table sections.</summary>
    private void AppendSection(StringBuilder text, TomlTable table, List<string> path, ref bool content) {
        bool values = false, sections = false;
        for (int i = 0; i < table._count; i++) {
            if (IsSectionChild(table._nodes[i])) sections = true;
            else values = true;
        }
        // A table that only holds sub-tables needs no header of its own, as [b.c] makes b.
        if (values || !sections || table._tableKind == TomlTableKind.ArrayItem || (table._state & TomlNodeState.CommentChanged) != 0) {
            if (content && Layout.BlankLineBeforeHeaders) text.Append(Layout.NewLine);
            content = true;
            AppendHeader(text, table, path);
        }
        for (int i = 0; i < table._count; i++) {
            TomlNode node = table._nodes[i];
            if (!IsSectionChild(node)) AppendLine(text, Layout.KeyIndent, null, node, Layout.KeySeparator);
        }
        for (int i = 0; i < table._count; i++) {
            TomlNode node = table._nodes[i];
            if (node is TomlTableArray tables) {
                foreach (TomlTable item in tables._items) AppendSection(text, item, Append(path, table._nodes[i]._key!), ref content);
            } else if (IsSectionChild(node)) {
                AppendSection(text, (TomlTable)node, Append(path, table._nodes[i]._key!), ref content);
            }
        }
    }

    private static bool IsSectionChild(TomlNode node) =>
        node is TomlTableArray || (node is TomlTable table && table._tableKind != TomlTableKind.Inline);

    private void AppendHeader(StringBuilder text, TomlTable table, List<string> path) {
        bool array = table._tableKind == TomlTableKind.ArrayItem;
        text.Append(array ? "[[" : "[");
        TomlFormatter.AppendKeys(text, path);
        text.Append(array ? "]]" : "]");
        AppendNewComment(text, table);
        text.Append(Layout.NewLine);
    }

    private static List<string> Append(List<string> path, string key) => new List<string>(path) { key };

    /// <summary>Returns the header path of a table: the keys from the root, without array-of-tables indexes.</summary>
    private static List<string> PathOf(TomlNode node) {
        var path = new List<string>();
        for (TomlNode? current = node; current != null; current = current._parent) {
            if (current._key != null) path.Insert(0, current._key);
        }
        return path;
    }

    // ---- Arrays and inline tables ------------------------------------------------------------------------------

    /// <summary>One element of a bracketed container as written: an array item or an inline-table key/value pair.</summary>
    private readonly struct Element {
        public Element(TomlNode node, List<string>? prefix) {
            Node = node;
            Prefix = prefix;
        }

        public TomlNode Node { get; }

        /// <summary>For inline-table entries, the dotted keys between the inline table and the entry's table.</summary>
        public List<string>? Prefix { get; }

        public int Start => Node._keyStart >= 0 ? Node._keyStart : Node._start;

        public int End => Node._end;
    }

    private void VisitArray(TomlArray array) {
        var current = new List<Element>(array._count);
        for (int i = 0; i < array._count; i++) current.Add(new Element(array._items[i], null));
        var removed = new List<Element>();
        if (array._removed != null) {
            foreach (TomlNode node in array._removed) {
                if (node.IsSource) removed.Add(new Element(node, null));
            }
        }
        EditContainer(array, current, removed, inline: false);
    }

    private void VisitInlineTable(TomlTable table) {
        var current = new List<Element>();
        var removed = new List<Element>();
        CollectInline(table, null, current, removed);
        EditContainer(table, current, removed, inline: true);
    }

    /// <summary>Lists the key/value pairs written inside an inline table, looking through its dotted tables.</summary>
    private static void CollectInline(TomlTable table, List<string>? prefix, List<Element> current, List<Element> removed) {
        for (int i = 0; i < table._count; i++) {
            TomlNode node = table._nodes[i];
            if (node is TomlTable dotted && dotted._tableKind == TomlTableKind.Dotted) {
                var inner = prefix == null ? new List<string>() : new List<string>(prefix);
                inner.Add(table._nodes[i]._key!);
                CollectInline(dotted, inner, current, removed);
            } else {
                current.Add(new Element(node, prefix));
            }
        }
        if (table._removed != null) {
            foreach (TomlNode node in table._removed) CollectRemoved(node, removed);
        }
    }

    private static void CollectRemoved(TomlNode node, List<Element> removed) {
        if (node.IsKeyValue) {
            removed.Add(new Element(node, null));
        } else if (node is TomlTable dotted) {
            for (int i = 0; i < dotted._count; i++) CollectRemoved(dotted._nodes[i], removed);
            if (dotted._removed != null) {
                foreach (TomlNode inner in dotted._removed) CollectRemoved(inner, removed);
            }
        }
    }

    /// <summary>Applies removed, added and changed elements of an array or inline table.</summary>
    private void EditContainer(TomlNode container, List<Element> current, List<Element> removed, bool inline) {
        int open = container._start, close = container._end - 1;
        var survivors = new List<Element>();
        foreach (Element element in current) {
            if (element.Node.IsSource) survivors.Add(element);
        }
        if (survivors.Count == 0) {
            if (removed.Count > 0 || current.Count > 0) Replace(open + 1, close, RenderElements(current, inline, padded: inline && current.Count > 0));
            return;
        }
        bool multiline = _source.IndexOf('\n', open, close - open) >= 0;
        string separator = InferSeparator(survivors, inline);
        if (removed.Count > 0) DeleteElements(survivors, removed, multiline);
        InsertElements(current, survivors, multiline, separator, inline);
        foreach (Element survivor in survivors) Visit(survivor.Node);
    }

    private string InferSeparator(List<Element> survivors, bool inline) {
        if (!inline) return " = ";
        TomlNode first = survivors[0].Node;
        int keyEnd = first.PairKeyEnd;
        return keyEnd >= 0 ? _source.Substring(keyEnd, first._start - keyEnd) : " = ";
    }

    private void DeleteElements(List<Element> survivors, List<Element> removed, bool multiline) {
        var all = new List<Element>(survivors.Count + removed.Count);
        all.AddRange(survivors);
        all.AddRange(removed);
        all.Sort((a, b) => a.Start.CompareTo(b.Start));
        var gone = new HashSet<TomlNode>();
        foreach (Element element in removed) gone.Add(element.Node);
        for (int k = 0; k < all.Count; k++) {
            Element element = all[k];
            if (!gone.Contains(element.Node)) continue;
            if (multiline && OwnsLine(element, out int lineStart, out int lineEnd)) {
                Delete(lineStart, lineEnd);
            } else if (k + 1 < all.Count) {
                Delete(element.Start, all[k + 1].Start);
            } else {
                int previous = -1;
                for (int p = k - 1; p >= 0; p--) {
                    if (!gone.Contains(all[p].Node)) {
                        previous = all[p].End;
                        break;
                    }
                }
                if (previous >= 0) Delete(previous, element.End);
            }
        }
    }

    /// <summary>Returns whether an element is alone on its line (with its comma and comment), and that line's range.</summary>
    private bool OwnsLine(Element element, out int lineStart, out int lineEnd) {
        lineStart = _text.LineStart(element.Start);
        lineEnd = 0;
        if (!_text.IsWhitespace(lineStart, element.Start)) return false;
        int i = _text.SkipSpaces(element.End);
        if (i < _source.Length && _source[i] == ',') i = _text.SkipSpaces(i + 1);
        if (i < _source.Length && _source[i] == '#') i = _text.ContentEnd(i);
        if (i < _source.Length && _source[i] != '\n' && _source[i] != '\r') return false;
        lineEnd = _text.LineEnd(i);
        return true;
    }

    private void InsertElements(List<Element> current, List<Element> survivors, bool multiline, string separator, bool inline) {
        var pending = new List<Element>();
        for (int i = 0; i < current.Count; i++) {
            Element element = current[i];
            if (!element.Node.IsSource) {
                pending.Add(element);
                continue;
            }
            if (pending.Count == 0) continue;
            // New elements before a loaded one go in front of it, each followed by a separator.
            var text = new StringBuilder();
            string indent = multiline ? _text.Indentation(element.Start) : string.Empty;
            foreach (Element added in pending) {
                AppendElement(text, added, separator, inline);
                text.Append(multiline ? "," + Layout.NewLine + indent : ", ");
            }
            Insert(element.Start, text.ToString());
            pending.Clear();
        }
        if (pending.Count == 0) return;
        Element last = survivors[survivors.Count - 1];
        int after = _text.SkipSpaces(last.End);
        bool comma = after < _source.Length && _source[after] == ',';
        int lineBreak = LineBreakAfter(comma ? after + 1 : last.End);
        if (!multiline || lineBreak < 0) {
            var text = new StringBuilder();
            foreach (Element added in pending) {
                text.Append(", ");
                AppendElement(text, added, separator, inline);
            }
            Insert(last.End, text.ToString());
            return;
        }
        string lineIndent = _text.Indentation(last.Start);
        var lines = new StringBuilder();
        for (int i = 0; i < pending.Count; i++) {
            lines.Append(Layout.NewLine).Append(lineIndent);
            AppendElement(lines, pending[i], separator, inline);
            if (comma || i < pending.Count - 1) lines.Append(',');
        }
        if (!comma) Insert(last.End, ",");
        Insert(lineBreak, lines.ToString());
    }

    /// <summary>Returns the line break that follows only spaces and a comment after a position, or -1.</summary>
    private int LineBreakAfter(int position) {
        int i = _text.SkipSpaces(position);
        if (i < _source.Length && _source[i] == '#') i = _text.ContentEnd(i);
        if (i < _source.Length && (_source[i] == '\n' || _source[i] == '\r')) return i;
        return -1;
    }

    private void AppendElement(StringBuilder text, Element element, string separator, bool inline) {
        if (inline) {
            if (element.Prefix != null) {
                foreach (string key in element.Prefix) {
                    TomlFormatter.AppendKey(text, key);
                    text.Append('.');
                }
            }
            TomlFormatter.AppendKey(text, element.Node._key!);
            text.Append(separator);
        }
        AppendValue(text, element.Node);
    }

    private string RenderElements(List<Element> elements, bool inline, bool padded) {
        var text = new StringBuilder();
        if (padded) text.Append(' ');
        for (int i = 0; i < elements.Count; i++) {
            if (i > 0) text.Append(", ");
            AppendElement(text, elements[i], " = ", inline);
        }
        if (padded) text.Append(' ');
        return text.ToString();
    }

    // ---- Values and comments -----------------------------------------------------------------------------------

    private string RenderValue(TomlNode node) {
        var text = new StringBuilder();
        AppendValue(text, node);
        return text.ToString();
    }

    /// <summary>Writes a value: unchanged loaded values as their source text, everything else from the model.</summary>
    internal static void AppendValue(StringBuilder text, TomlNode node) {
        if (node.IsSource && (node._state & (TomlNodeState.Dirty | TomlNodeState.Changed)) == 0) {
            text.Append(node._document._source, node._start, node._end - node._start);
            return;
        }
        switch (node) {
            case TomlValue value:
                TomlFormatter.AppendScalar(text, value);
                break;
            case TomlArray array:
                text.Append('[');
                for (int i = 0; i < array._count; i++) {
                    if (i > 0) text.Append(", ");
                    AppendValue(text, array._items[i]);
                }
                text.Append(']');
                break;
            case TomlTable table:
                AppendInlineTable(text, table);
                break;
            case TomlTableArray tables:
                text.Append('[');
                for (int i = 0; i < tables._items.Count; i++) {
                    if (i > 0) text.Append(", ");
                    AppendInlineTable(text, tables._items[i]);
                }
                text.Append(']');
                break;
        }
    }

    private static void AppendInlineTable(StringBuilder text, TomlTable table) {
        if (table._count == 0) {
            text.Append("{}");
            return;
        }
        text.Append("{ ");
        for (int i = 0; i < table._count; i++) {
            if (i > 0) text.Append(", ");
            TomlFormatter.AppendKey(text, table._nodes[i]._key!);
            text.Append(" = ");
            AppendValue(text, table._nodes[i]);
        }
        text.Append(" }");
    }

    /// <summary>Sets, replaces or removes the comment after a loaded node on its line.</summary>
    private void WriteComment(TomlNode node) {
        string? comment = node.ChangedComment;
        int hash = _text.FindTrailingComment(node._end);
        if (hash >= 0) {
            int end = _text.ContentEnd(hash);
            if (comment is null) Delete(_text.SkipSpacesBack(hash, node._end), end);
            else Replace(hash, end, CommentText(comment));
        } else if (comment != null) {
            Insert(node._end, Layout.CommentGap + CommentText(comment));
        }
    }

    private static string CommentText(string comment) => comment.Length > 0 && comment[0] == ' ' ? "#" + comment : "# " + comment;

    // ---- ToString of a node ------------------------------------------------------------------------------------

    public static string NodeText(TomlNode node) {
        if (node is TomlTable table && table._tableKind != TomlTableKind.Inline) {
            if (table._tableKind == TomlTableKind.Root) return node._document.ToString();
            var rendering = new TomlRendering(node._document);
            var text = new StringBuilder();
            bool content = false;
            rendering.AppendSectionCopy(text, table, PathOf(table), ref content);
            return text.ToString();
        }
        if (node is TomlTableArray tables) {
            var rendering = new TomlRendering(node._document);
            var text = new StringBuilder();
            bool content = false;
            foreach (TomlTable item in tables._items) rendering.AppendSectionCopy(text, item, PathOf(tables), ref content);
            return text.ToString();
        }
        var value = new StringBuilder();
        AppendValue(value, node);
        return value.ToString();
    }

    private void AppendSectionCopy(StringBuilder text, TomlTable table, List<string> path, ref bool content) => AppendSection(text, table, path, ref content);
}
