namespace OfficeIMO.Yaml;

/// <summary>
/// Evaluates node paths (<c>spec.template.spec.containers[0].image</c>, <c>metadata.labels["app.kubernetes.io/name"]</c>).
/// The grammar is documented on <see cref="YamlNode.Select"/>. A malformed path throws even when an earlier step
/// already found nothing, so the outcome never depends on the data.
/// </summary>
internal static class YamlPath {
    public static YamlNode? Select(YamlDocument document, string path) {
        if (path is null) throw new ArgumentNullException(nameof(path));
        IReadOnlyList<YamlDocumentPart> parts = document.Documents;
        if (parts.Count > 1 && path.Length > 1 && path[0] == '$' && path[1] == '[') {
            int i = 2;
            long index = ReadIndex(path, ref i);
            YamlNode? root = index < parts.Count ? parts[(int)index].Root : null;
            return Walk(root, path, i);
        }
        return Select(parts.Count == 0 ? null : parts[0].Root, path, 0);
    }

    public static YamlNode? Select(YamlNode? node, string path, int position) {
        if (path is null) throw new ArgumentNullException(nameof(path));
        int i = position;
        if (i < path.Length && path[i] == '$') return Walk(node, path, i + 1);
        // The first step may be a bare name without its leading dot.
        if (i < path.Length && path[i] != '.' && path[i] != '[') {
            string name = ReadName(path, ref i);
            node = Key(node, name);
        }
        return Walk(node, path, i);
    }

    private static YamlNode? Walk(YamlNode? node, string path, int i) {
        while (i < path.Length) {
            char ch = path[i];
            if (ch == '.') {
                i++;
                node = Key(node, ReadName(path, ref i));
            } else if (ch == '[') {
                i++;
                if (i < path.Length && path[i] == '"') {
                    string key = ReadQuoted(path, ref i);
                    Expect(path, ref i, ']');
                    node = Key(node, key);
                } else {
                    long index = ReadIndex(path, ref i);
                    node = node?.Resolved is YamlSequence sequence && index < sequence._count ? sequence._items[index] : null;
                }
            } else {
                throw Malformed(path, i, "expected '.' or '['");
            }
        }
        return node;
    }

    private static YamlNode? Key(YamlNode? node, string key) =>
        node?.Resolved is YamlMapping mapping && mapping.TryGetValue(key, out YamlNode? value) ? value : null;

    private static string ReadName(string path, ref int i) {
        int start = i;
        while (i < path.Length && path[i] != '.' && path[i] != '[') i++;
        if (i == start) throw Malformed(path, start, "expected a key name");
        return path.Substring(start, i - start);
    }

    /// <summary>Reads <c>n]</c> after an opening bracket. Indexes beyond the int range simply match nothing.</summary>
    private static long ReadIndex(string path, ref int i) {
        int start = i;
        long value = 0;
        while (i < path.Length && path[i] >= '0' && path[i] <= '9') {
            if (value < int.MaxValue) value = value * 10 + (path[i] - '0');
            i++;
        }
        if (i == start) throw Malformed(path, start, "expected an index or a quoted key");
        Expect(path, ref i, ']');
        return value;
    }

    /// <summary>Reads a double-quoted key with JSON escapes, as <c>OfficeIMO.Reader.Yaml</c> writes them.</summary>
    private static string ReadQuoted(string path, ref int i) {
        int open = i++;
        var text = new StringBuilder();
        while (true) {
            if (i >= path.Length) throw Malformed(path, open, "the quoted key is not closed");
            char ch = path[i++];
            if (ch == '"') return text.ToString();
            if (ch != '\\') {
                text.Append(ch);
                continue;
            }
            if (i >= path.Length) throw Malformed(path, i, "the escape is incomplete");
            char escape = path[i++];
            switch (escape) {
                case '"': case '\\': case '/': text.Append(escape); break;
                case 'b': text.Append('\b'); break;
                case 'f': text.Append('\f'); break;
                case 'n': text.Append('\n'); break;
                case 'r': text.Append('\r'); break;
                case 't': text.Append('\t'); break;
                case 'u':
                    int code = 0;
                    for (int k = 0; k < 4; k++) {
                        int digit = i < path.Length ? YamlChars.HexValue(path[i]) : -1;
                        if (digit < 0) throw Malformed(path, i, "expected four hexadecimal digits after \\u");
                        code = code * 16 + digit;
                        i++;
                    }
                    text.Append((char)code);
                    break;
                default:
                    throw Malformed(path, i - 1, "unknown escape '\\" + escape + "'");
            }
        }
    }

    private static void Expect(string path, ref int i, char ch) {
        if (i >= path.Length || path[i] != ch) throw Malformed(path, i, "expected '" + ch + "'");
        i++;
    }

    private static ArgumentException Malformed(string path, int position, string reason) =>
        new ArgumentException("The YAML path '" + path + "' is malformed at position " + position.ToString(CultureInfo.InvariantCulture) + ": " + reason + ".", nameof(path));
}
