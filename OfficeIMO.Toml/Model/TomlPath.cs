namespace OfficeIMO.Toml;

/// <summary>
/// Evaluates node paths (<c>tool.ruff.line-length</c>, <c>package[0].name</c>, <c>tool.uv.sources["my.pkg"]</c>). The
/// grammar is documented on <see cref="TomlNode.Select"/>. A malformed path throws even when an earlier step already
/// found nothing, so the outcome never depends on the data.
/// </summary>
internal static class TomlPath {
    private readonly struct Step {
        public Step(string key) {
            Key = key;
            Index = -1;
        }

        public Step(long index) {
            Key = null;
            Index = index;
        }

        public string? Key { get; }

        public long Index { get; }
    }

    public static TomlNode? Select(TomlNode node, string path) {
        List<Step> steps = Parse(path);
        TomlNode? current = node;
        foreach (Step step in steps) current = current is null ? null : Child(current, step);
        return current;
    }

    public static TomlNode Set(TomlTable root, string path, object value) {
        List<Step> steps = Parse(path);
        if (steps.Count == 0) throw new ArgumentException("The path must address a key or an index, not the root.", nameof(path));
        if (value is null) throw new ArgumentNullException(nameof(value));
        TomlNode current = root;
        for (int i = 0; i < steps.Count - 1; i++) {
            Step step = steps[i];
            TomlNode? next = Child(current, step);
            if (next is null) {
                if (step.Key is null || current is not TomlTable table) throw new InvalidOperationException("The path step " + Describe(step) + " does not exist.");
                next = table.AddTable(step.Key);
            }
            if (next is TomlValue) throw new InvalidOperationException("The path step " + Describe(step) + " is a value, not a table or array.");
            current = next;
        }
        Step last = steps[steps.Count - 1];
        if (last.Key != null) {
            if (current is TomlTable table) return table.Set(last.Key, value);
            throw new InvalidOperationException("The path step " + Describe(last) + " needs a table but found " + TomlNode.Describe(current.Kind) + ".");
        }
        if (current is TomlArray array) {
            if (last.Index > array.Count) throw new InvalidOperationException("The index " + Describe(last) + " is outside the array.");
            return array.Set((int)last.Index, value);
        }
        throw new InvalidOperationException("The path step " + Describe(last) + " needs an array but found " + TomlNode.Describe(current.Kind) + ".");
    }

    public static bool Remove(TomlTable root, string path) {
        List<Step> steps = Parse(path);
        if (steps.Count == 0) throw new ArgumentException("The root cannot be removed.", nameof(path));
        TomlNode? parent = root;
        for (int i = 0; i < steps.Count - 1 && parent != null; i++) parent = Child(parent, steps[i]);
        Step last = steps[steps.Count - 1];
        switch (parent) {
            case TomlTable table when last.Key != null:
                return table.Remove(last.Key);
            case TomlArray array when last.Key is null && last.Index < array.Count:
                array.RemoveAt((int)last.Index);
                return true;
            case TomlTableArray tables when last.Key is null && last.Index < tables.Count:
                tables.RemoveAt((int)last.Index);
                return true;
            default:
                return false;
        }
    }

    private static TomlNode? Child(TomlNode node, Step step) {
        if (step.Key != null) return node is TomlTable table && table.TryGetValue(step.Key, out TomlNode? value) ? value : null;
        return node switch {
            TomlArray array when step.Index < array.Count => array[(int)step.Index],
            TomlTableArray tables when step.Index < tables.Count => tables[(int)step.Index],
            _ => null
        };
    }

    private static string Describe(Step step) => step.Key != null ? "'" + step.Key + "'" : "[" + step.Index.ToString(CultureInfo.InvariantCulture) + "]";

    private static List<Step> Parse(string path) {
        if (path is null) throw new ArgumentNullException(nameof(path));
        var steps = new List<Step>();
        int i = 0;
        if (i < path.Length && path[i] == '$') {
            i++;
        } else if (i < path.Length && path[i] != '.' && path[i] != '[') {
            steps.Add(new Step(ReadName(path, ref i)));
        }
        while (i < path.Length) {
            char ch = path[i];
            if (ch == '.') {
                i++;
                steps.Add(new Step(ReadName(path, ref i)));
            } else if (ch == '[') {
                i++;
                if (i < path.Length && path[i] == '"') {
                    string key = ReadQuoted(path, ref i);
                    Expect(path, ref i, ']');
                    steps.Add(new Step(key));
                } else {
                    steps.Add(new Step(ReadIndex(path, ref i)));
                }
            } else {
                throw Malformed(path, i, "expected '.' or '['");
            }
        }
        return steps;
    }

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

    /// <summary>Reads a double-quoted key with JSON escapes, as <c>OfficeIMO.Reader.Toml</c> writes them.</summary>
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
                        int digit = i < path.Length ? TomlScalars.HexValue(path[i]) : -1;
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
        new ArgumentException("The TOML path '" + path + "' is malformed at position " + position.ToString(CultureInfo.InvariantCulture) + ": " + reason + ".", nameof(path));
}
