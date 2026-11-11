namespace OfficeIMO.Toml;

/// <summary>Change tracking shared by the editing methods.</summary>
internal static class TomlEditing {
    /// <summary>Marks a node and its ancestors as changed so saving walks into them.</summary>
    public static void MarkChanged(TomlNode node) {
        for (TomlNode? current = node; current != null; current = current._parent) {
            if ((current._state & TomlNodeState.Detached) != 0) return;
            if ((current._state & TomlNodeState.Changed) != 0 && !ReferenceEquals(current, node)) return;
            current._state |= TomlNodeState.Changed;
        }
    }

    /// <summary>Marks a removed node; loaded nodes are remembered so saving deletes their text.</summary>
    public static void Detach(TomlNode node, ref List<TomlNode>? removed) {
        node._state |= TomlNodeState.Detached;
        if (node.IsLoaded) (removed ??= new List<TomlNode>()).Add(node);
    }

    /// <summary>Gives a new node the position of the loaded node it replaces, so saving writes it in the same place.</summary>
    public static void Adopt(TomlNode old, TomlNode node) {
        old._state |= TomlNodeState.Detached;
        if (!old.IsSource) return;
        node._start = old._start;
        node._end = old._end;
        node._keyStart = old._keyStart;
        node._keyEnd = old._keyEnd;
        node._state |= TomlNodeState.Dirty;
    }

    /// <summary>
    /// Returns the quoting new strings in a container should use: literal when most of its loaded single-line strings
    /// are literal, otherwise basic.
    /// </summary>
    public static TomlStringStyle PreferredStringStyle(TomlNode container) {
        int basic = 0, literal = 0;
        if (container is TomlTable table) {
            for (int i = 0; i < table._count; i++) Count(table._nodes[i], ref basic, ref literal);
        } else if (container is TomlArray array) {
            for (int i = 0; i < array._count; i++) Count(array._items[i], ref basic, ref literal);
        }
        return literal > basic ? TomlStringStyle.Literal : TomlStringStyle.Basic;
    }

    private static void Count(TomlNode node, ref int basic, ref int literal) {
        if (node is not TomlValue value || value._kind != TomlNodeKind.String || !value.IsSource) return;
        if (value._style == TomlStringStyle.Literal) literal++;
        else if (value._style == TomlStringStyle.Basic) basic++;
    }
}

/// <summary>Reads comments around nodes from the retained source.</summary>
internal static class TomlTrivia {
    public static IReadOnlyList<string> GetLeadingComments(TomlNode node) {
        int anchor = node._keyStart >= 0 ? node._keyStart : node._start;
        if (anchor < 0) return Array.Empty<string>();
        string source = node._document._source;
        TomlText text = new TomlText(source);
        int line = text.LineStart(anchor);
        if (!text.IsWhitespace(line, anchor)) return Array.Empty<string>();
        List<string>? comments = null;
        while (line > 0) {
            int previous = text.LineStart(line - 1);
            int hash = text.SkipSpaces(previous);
            if (hash >= line || source[hash] != '#') break;
            (comments ??= new List<string>()).Insert(0, source.Substring(hash + 1, text.ContentEnd(hash) - hash - 1));
            line = previous;
        }
        return (IReadOnlyList<string>?)comments ?? Array.Empty<string>();
    }

    public static string? GetTrailingComment(TomlNode node) {
        if (node._end < 0) return null;
        string source = node._document._source;
        var text = new TomlText(source);
        int hash = text.FindTrailingComment(node._end);
        return hash < 0 ? null : source.Substring(hash + 1, text.ContentEnd(hash) - hash - 1);
    }
}

/// <summary>Line-oriented helpers over the source text.</summary>
internal readonly struct TomlText {
    public TomlText(string source) {
        Source = source;
    }

    public string Source { get; }

    public int Length => Source.Length;

    /// <summary>Returns the start of the line containing <paramref name="offset"/>.</summary>
    public int LineStart(int offset) {
        int i = Math.Min(offset, Source.Length);
        while (i > 0 && Source[i - 1] != '\n') i--;
        return i;
    }

    /// <summary>Returns the position after the line break ending the line that contains <paramref name="offset"/>, or the end of the text.</summary>
    public int LineEnd(int offset) {
        int i = offset;
        while (i < Source.Length && Source[i] != '\n') i++;
        return i < Source.Length ? i + 1 : i;
    }

    /// <summary>Returns the position of the line break (or end of text) ending the line containing <paramref name="offset"/>.</summary>
    public int ContentEnd(int offset) {
        int i = offset;
        while (i < Source.Length && Source[i] != '\n') i++;
        if (i > offset && i <= Source.Length && i > 0 && Source[i - 1] == '\r') i--;
        return i;
    }

    public int SkipSpaces(int offset) {
        int i = offset;
        while (i < Source.Length && (Source[i] == ' ' || Source[i] == '\t')) i++;
        return i;
    }

    /// <summary>Moves back over spaces and tabs before <paramref name="offset"/>, not past <paramref name="floor"/>.</summary>
    public int SkipSpacesBack(int offset, int floor) {
        int i = offset;
        while (i > floor && (Source[i - 1] == ' ' || Source[i - 1] == '\t')) i--;
        return i;
    }

    public bool IsWhitespace(int start, int end) {
        for (int i = start; i < end; i++) {
            if (Source[i] != ' ' && Source[i] != '\t') return false;
        }
        return true;
    }

    /// <summary>Returns whether the line starting at <paramref name="lineStart"/> is empty or only spaces.</summary>
    public bool IsBlankLine(int lineStart) {
        int i = SkipSpaces(lineStart);
        return i >= Source.Length || Source[i] == '\n' || (Source[i] == '\r' && i + 1 < Source.Length && Source[i + 1] == '\n');
    }

    public bool IsCommentLine(int lineStart) {
        int i = SkipSpaces(lineStart);
        return i < Source.Length && Source[i] == '#';
    }

    /// <summary>Moves a line start up over the comment lines directly above it.</summary>
    public int BlockStart(int lineStart, int floor) {
        int line = lineStart;
        while (line > floor) {
            int previous = LineStart(line - 1);
            if (previous < floor || !IsCommentLine(previous)) break;
            line = previous;
        }
        return line;
    }

    /// <summary>Moves a line start up over the blank lines directly above it.</summary>
    public int SkipBlankLinesUp(int lineStart, int floor) {
        int line = lineStart;
        while (line > floor) {
            int previous = LineStart(line - 1);
            if (previous < floor || !IsBlankLine(previous)) break;
            line = previous;
        }
        return line;
    }

    /// <summary>Returns the <c>#</c> of the comment after a node on its line (past an optional comma), or -1.</summary>
    public int FindTrailingComment(int end) {
        int i = SkipSpaces(end);
        if (i < Source.Length && Source[i] == ',') i = SkipSpaces(i + 1);
        return i < Source.Length && Source[i] == '#' ? i : -1;
    }

    /// <summary>Returns the leading whitespace of the line containing <paramref name="offset"/>.</summary>
    public string Indentation(int offset) {
        int start = LineStart(offset);
        return Source.Substring(start, SkipSpaces(start) - start);
    }
}

/// <summary>Formatting conventions inferred once from a loaded document, used for text added by editing.</summary>
internal sealed class TomlLayout {
    public TomlLayout(TomlDocument document) {
        string source = document._source;
        int newline = source.IndexOf('\n');
        NewLine = newline > 0 && source[newline - 1] == '\r' ? "\r\n" : "\n";
        var text = new TomlText(source);
        int separated = 0, total = 0;
        foreach (TomlTable header in document._headers) {
            int line = text.BlockStart(text.LineStart(header._start), 0);
            if (line == 0) continue;
            total++;
            if (text.IsBlankLine(text.LineStart(line - 1))) separated++;
        }
        BlankLineBeforeHeaders = total == 0 || separated * 2 >= total;
        KeySeparator = " = ";
        KeyIndent = string.Empty;
        TomlNode? sample = FindKeyValue(document._root);
        if (sample != null) KeySeparator = source.Substring(sample._keyEnd, sample._start - sample._keyEnd);
        foreach (TomlTable header in document._headers) {
            TomlNode? first = FindKeyValue(header);
            if (first != null) {
                KeyIndent = text.Indentation(first._keyStart);
                if (sample == null) KeySeparator = source.Substring(first._keyEnd, first._start - first._keyEnd);
                break;
            }
        }
        if (KeySeparator.IndexOf('=') < 0) KeySeparator = " = ";
        CommentGap = FindCommentGap(document, text) ?? " ";
    }

    /// <summary>Whitespace between a value and its trailing comment, such as the two spaces PEP 8 style uses.</summary>
    public string CommentGap { get; }

    private static string? FindCommentGap(TomlDocument document, TomlText text) {
        string? gap = Gap(document._root, text);
        if (gap != null) return gap;
        foreach (TomlTable header in document._headers) {
            gap = Gap(header, text);
            if (gap != null) return gap;
        }
        return null;
    }

    private static string? Gap(TomlTable table, TomlText text) {
        for (int i = 0; i < table._count; i++) {
            TomlNode node = table._nodes[i];
            if (!node.IsKeyValue) continue;
            int hash = text.FindTrailingComment(node._end);
            if (hash > node._end && text.Source[node._end] != ',') return text.Source.Substring(node._end, hash - node._end);
        }
        return null;
    }

    public string NewLine { get; }

    public bool BlankLineBeforeHeaders { get; }

    /// <summary>Text between a key and its value, such as <c> = </c>.</summary>
    public string KeySeparator { get; }

    /// <summary>Indentation of key/value lines under headers.</summary>
    public string KeyIndent { get; }

    private static TomlNode? FindKeyValue(TomlTable table) {
        for (int i = 0; i < table._count; i++) {
            TomlNode node = table._nodes[i];
            if (node.IsKeyValue) return node;
            if (node is TomlTable dotted && dotted._tableKind == TomlTableKind.Dotted) {
                TomlNode? inner = FindKeyValue(dotted);
                if (inner != null) return inner;
            }
        }
        return null;
    }
}
