namespace OfficeIMO.Toml;

/// <summary>
/// Finds the source text of the key segment that names a loaded node. Nodes store only where their key/value pair or
/// header starts, so a key is found by reading the segments of that key again; a table made by dotted keys or by deeper
/// headers is named by a segment of its first entry's key.
/// </summary>
internal static class TomlKeyLocator {
    public static bool TryFind(TomlNode node, out int start, out int end) {
        start = end = -1;
        if (node._key is null || (node._state & TomlNodeState.Detached) != 0) return false;
        string source = node._document._source;
        // Segments of the defining key that come after the node's own segment.
        int after = 0;
        TomlNode current = node;
        while (true) {
            if (current.IsKeyValue) return FindSegment(source, current._keyStart, after, out start, out end);
            switch (current) {
                case TomlTable { _tableKind: TomlTableKind.Header or TomlTableKind.ArrayItem } table when table.IsSource:
                    return FindSegment(source, HeaderKeyStart(source, table._start), after, out start, out end);
                case TomlTable { _tableKind: TomlTableKind.Dotted or TomlTableKind.Implicit } table: {
                    TomlNode? first = FirstLoaded(table._nodes, table._count);
                    if (first is null) return false;
                    after++;
                    current = first;
                    break;
                }
                case TomlTableArray tables: {
                    // A table of the array is named by the last segment of its [[header]], like the array itself.
                    TomlTable? first = null;
                    for (int i = 0; i < tables.Count && first is null; i++) {
                        if (tables[i].IsSource) first = tables[i];
                    }
                    if (first is null) return false;
                    current = first;
                    break;
                }
                default:
                    return false;
            }
        }
    }

    private static TomlNode? FirstLoaded(TomlNode[] nodes, int count) {
        for (int i = 0; i < count; i++) {
            if (nodes[i].IsLoaded) return nodes[i];
        }
        return null;
    }

    private static int HeaderKeyStart(string source, int start) {
        int i = start + 1;
        if (i < source.Length && source[i] == '[') i++;
        return SkipBlank(source, i);
    }

    /// <summary>Finds a segment of the dotted key starting at <paramref name="keyStart"/>, counted from the last one.</summary>
    private static bool FindSegment(string source, int keyStart, int fromLast, out int start, out int end) {
        int count = 0;
        int i = keyStart;
        while (true) {
            count++;
            i = SkipBlank(source, SegmentEnd(source, i));
            if (i >= source.Length || source[i] != '.') break;
            i = SkipBlank(source, i + 1);
        }
        int target = count - 1 - fromLast;
        start = end = -1;
        if (target < 0) return false;
        i = keyStart;
        for (int segment = 0; segment < target; segment++) i = SkipBlank(source, SkipBlank(source, SegmentEnd(source, i)) + 1);
        start = i;
        end = SegmentEnd(source, i);
        return true;
    }

    private static int SegmentEnd(string source, int i) {
        if (i >= source.Length) return i;
        char quote = source[i];
        if (quote == '"' || quote == '\'') {
            for (i++; i < source.Length && source[i] != quote; i++) {
                if (quote == '"' && source[i] == '\\') i++;
            }
            return Math.Min(i + 1, source.Length);
        }
        while (i < source.Length && TomlScalars.IsBareKeyChar(source[i])) i++;
        return i;
    }

    private static int SkipBlank(string source, int i) {
        while (i < source.Length && (source[i] == ' ' || source[i] == '\t')) i++;
        return i;
    }
}
