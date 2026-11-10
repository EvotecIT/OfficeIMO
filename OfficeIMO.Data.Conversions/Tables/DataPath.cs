namespace OfficeIMO.Data.Conversions;

/// <summary>One step of a path: a property name or an array index.</summary>
internal readonly struct DataPathSegment {
    public DataPathSegment(string name) {
        Name = name;
        Index = -1;
    }

    public DataPathSegment(int index) {
        Name = null;
        Index = index;
    }

    public string? Name { get; }
    public int Index { get; }
    public bool IsIndex => Name == null;
}

/// <summary>
/// The path grammar shared with the YAML engine's <c>Select</c>: <c>name</c>, <c>.name</c>, <c>[n]</c> and
/// <c>["key with . or ["]</c>. Table headers are written in it, so a header can be selected unchanged.
/// </summary>
internal static class DataPath {
    public static bool TryParse(string path, out List<DataPathSegment> segments) {
        segments = new List<DataPathSegment>();
        int i = 0;
        if (path.Length > 0 && path[0] == '$') i++;
        while (i < path.Length) {
            char c = path[i];
            if (c == '[') {
                i++;
                if (i < path.Length && path[i] == '"') {
                    var name = new StringBuilder();
                    i++;
                    while (i < path.Length && path[i] != '"') {
                        if (path[i] == '\\' && i + 1 < path.Length) {
                            char escaped = path[++i];
                            if (escaped == 'u' && i + 4 < path.Length && int.TryParse(path.Substring(i + 1, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int code)) {
                                name.Append((char)code);
                                i += 4;
                            } else {
                                name.Append(escaped switch { 'n' => '\n', 't' => '\t', 'r' => '\r', _ => escaped });
                            }
                        } else {
                            name.Append(path[i]);
                        }

                        i++;
                    }

                    if (i + 1 >= path.Length || path[i + 1] != ']') return false;
                    i += 2;
                    segments.Add(new DataPathSegment(name.ToString()));
                    continue;
                }

                int start = i;
                while (i < path.Length && path[i] >= '0' && path[i] <= '9') i++;
                if (i == start || i >= path.Length || path[i] != ']') return false;
                if (i - start > 1 && path[start] == '0') return false;
                if (!int.TryParse(path.Substring(start, i - start), NumberStyles.None, CultureInfo.InvariantCulture, out int index)) return false;
                segments.Add(new DataPathSegment(index));
                i++;
                continue;
            }

            if (c == '.') {
                if (segments.Count == 0 && i > 0 && path[0] != '$') return false;
                i++;
                if (i >= path.Length) return false;
            } else if (segments.Count > 0) {
                return false;
            }

            int nameStart = i;
            while (i < path.Length && path[i] != '.' && path[i] != '[' && path[i] != ']') i++;
            if (i == nameStart) return false;
            segments.Add(new DataPathSegment(path.Substring(nameStart, i - nameStart)));
        }

        return segments.Count > 0;
    }

    /// <summary>Whether a property name can be written as a plain header step without changing what it means.</summary>
    public static bool IsPlainName(string name) => name.Length > 0 && name.IndexOfAny(Reserved) < 0;

    private static readonly char[] Reserved = { '.', '[', ']' };
}
