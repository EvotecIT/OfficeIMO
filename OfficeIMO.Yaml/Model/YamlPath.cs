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
                    node = node?.Resolved is YamlSequence sequence && index < sequence._count ? sequence.ItemAt((int)index) : null;
                }
            } else {
                throw Malformed(path, i, "expected '.' or '['");
            }
        }
        return node;
    }

    /// <summary>One step of a parsed path: a key by text or a sequence index.</summary>
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

    /// <summary>Parses a path into its document index (-1 for none) and steps, with the grammar <see cref="Select(YamlNode?, string, int)"/> reads.</summary>
    private static List<Step> Parse(string path, out long document) {
        if (path is null) throw new ArgumentNullException(nameof(path));
        var steps = new List<Step>();
        document = -1;
        int i = 0;
        if (i < path.Length && path[i] == '$') {
            i++;
            if (path.Length > 2 && path[1] == '[' && path[2] >= '0' && path[2] <= '9') {
                i = 2;
                document = ReadIndex(path, ref i);
            }
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

    public static YamlNode Set(YamlDocument document, string path, YamlNode value) {
        List<Step> steps = Parse(path, out long index);
        if (steps.Count == 0) {
            YamlDocumentPart? part = PartFor(document, index, create: true, null);
            if (part is null) return document.AddDocument(value).Root;
            part.Root = value;
            return part.Root;
        }
        YamlNode parent = Walk(document, path, steps, index, steps.Count - 1, create: true)!;
        return SetStep(parent, steps[steps.Count - 1], value, path);
    }

    public static YamlNode Add(YamlDocument document, string path, YamlNode value) {
        List<Step> steps = Parse(path, out long index);
        YamlNode target = steps.Count == 0
            ? (PartFor(document, index, create: true, new YamlSequence()) ?? document.AddDocument(new YamlSequence())).Root
            : Walk(document, path, steps, index, steps.Count, create: true, lastIsSequence: true)!;
        if (target.Resolved is not YamlSequence sequence) throw new InvalidOperationException("The path '" + path + "' addresses " + Describe(target) + ", not a sequence.");
        return sequence.Add(value);
    }

    public static bool Remove(YamlDocument document, string path) {
        List<Step> steps = Parse(path, out long index);
        if (steps.Count == 0) throw new ArgumentException("A path to a document root cannot be removed; set the root instead.", nameof(path));
        YamlNode? parent = Walk(document, path, steps, index, steps.Count - 1, create: false);
        if (parent is null) return false;
        Step last = steps[steps.Count - 1];
        YamlNode resolved = parent.Resolved;
        if (last.Key is not null) {
            if (resolved is YamlMapping mapping) {
                int own = OwnIndex(mapping, last.Key);
                if (own < 0) return false;
                mapping.RemoveAt(own);
                return true;
            }
            if (resolved is YamlScalar) throw ThroughScalar(path, parent);
            return false;
        }
        if (resolved is YamlSequence sequence) {
            if (last.Index >= sequence._count) return false;
            sequence.RemoveAt((int)last.Index);
            return true;
        }
        if (resolved is YamlScalar) throw ThroughScalar(path, parent);
        return false;
    }

    private static YamlDocumentPart? PartFor(YamlDocument document, long index, bool create, YamlNode? root) {
        IReadOnlyList<YamlDocumentPart> parts = document.Documents;
        long position = index < 0 ? 0 : index;
        if (position < parts.Count) return parts[(int)position];
        if (!create) return null;
        if (position != parts.Count) throw new ArgumentOutOfRangeException(nameof(index), index, "The document index is past the end of the stream.");
        return root is null ? null : document.AddDocument(root);
    }

    /// <summary>
    /// Follows <paramref name="count"/> steps from a document root. With <paramref name="create"/>, missing keys get new
    /// mappings (or a sequence when the next step is an index) and an index equal to a sequence's count appends.
    /// </summary>
    private static YamlNode? Walk(YamlDocument document, string path, List<Step> steps, long documentIndex, int count, bool create, bool lastIsSequence = false) {
        YamlDocumentPart? part = PartFor(document, documentIndex, create, create ? NewContainer(steps[0]) : null);
        if (part is null) return null;
        YamlNode node = part.Root;
        if (create && node is YamlScalar { _style: YamlStyle.Plain } empty && empty.IsNull && empty.Tag is null && empty.Anchor is null) {
            part.Root = NewContainer(steps[0]);
            node = part.Root;
        }
        for (int s = 0; s < count; s++) {
            Step step = steps[s];
            YamlNode resolved = node.Resolved;
            YamlNode? next;
            if (step.Key is not null) {
                if (resolved is YamlScalar) throw ThroughScalar(path, node);
                if (resolved is not YamlMapping mapping) {
                    if (!create) return null;
                    throw new InvalidOperationException("The path '" + path + "' has a key step where " + Describe(node) + " is.");
                }
                // Own scalar keys first, then keys merges provide (reading only), then collection keys, as Select finds them.
                int own = mapping.IndexOf(step.Key);
                next = own >= 0 ? mapping.EntryAt(own).Value : null;
                if (next is null && !create && mapping.TryGetValue(step.Key, out YamlNode? merged)) next = merged;
                if (next is null && (own = mapping.IndexOfCollectionKey(step.Key)) >= 0) next = mapping.EntryAt(own).Value;
                if (next is null || (create && next is YamlScalar && next.IsNull && s + 1 < steps.Count)) {
                    if (!create) return null;
                    next = mapping.Set(step.Key, ContainerFor(steps, s + 1, lastIsSequence));
                }
            } else {
                if (resolved is YamlScalar) throw ThroughScalar(path, node);
                if (resolved is not YamlSequence sequence) {
                    if (!create) return null;
                    throw new InvalidOperationException("The path '" + path + "' has an index step where " + Describe(node) + " is.");
                }
                if (step.Index < sequence._count) {
                    next = sequence.ItemAt((int)step.Index);
                } else if (create && step.Index == sequence._count) {
                    next = sequence.Add(ContainerFor(steps, s + 1, lastIsSequence));
                } else if (create) {
                    throw new ArgumentOutOfRangeException(nameof(path), step.Index, "The index in '" + path + "' is past the end of its sequence.");
                } else {
                    return null;
                }
            }
            node = next;
        }
        return node;
    }

    private static YamlNode ContainerFor(List<Step> steps, int next, bool lastIsSequence) =>
        next < steps.Count ? NewContainer(steps[next]) : lastIsSequence ? new YamlSequence() : new YamlMapping();

    private static YamlNode NewContainer(Step step) => step.Key is null ? new YamlSequence() : new YamlMapping();

    private static YamlNode SetStep(YamlNode parent, Step step, YamlNode value, string path) {
        YamlNode resolved = parent.Resolved;
        if (resolved is YamlScalar) throw ThroughScalar(path, parent);
        if (step.Key is not null) {
            if (resolved is not YamlMapping mapping) throw new InvalidOperationException("The path '" + path + "' has a key step where " + Describe(parent) + " is.");
            int own = mapping.IndexOf(step.Key) >= 0 ? -1 : mapping.IndexOfCollectionKey(step.Key);
            return own >= 0 ? YamlEdit.ReplaceValue(mapping, own, value) : mapping.Set(step.Key, value);
        }
        if (resolved is not YamlSequence sequence) throw new InvalidOperationException("The path '" + path + "' has an index step where " + Describe(parent) + " is.");
        if (step.Index == sequence._count) return sequence.Add(value);
        if (step.Index > sequence._count) throw new ArgumentOutOfRangeException(nameof(path), step.Index, "The index in '" + path + "' is past the end of its sequence.");
        sequence[(int)step.Index] = value;
        return sequence.ItemAt((int)step.Index);
    }

    private static InvalidOperationException ThroughScalar(string path, YamlNode node) =>
        new InvalidOperationException("The path '" + path + "' goes through the scalar" + node.Where() + "; a scalar has no keys or items.");

    private static string Describe(YamlNode node) => node.Resolved switch {
        YamlMapping => "a mapping",
        YamlSequence => "a sequence",
        _ => "a scalar"
    };

    /// <summary>Formats a node's path from its document root, the inverse of <see cref="Select(YamlDocument, string)"/>.</summary>
    public static string Format(YamlNode node) {
        var steps = new List<string>();
        YamlNode current = node;
        while (current._parent is { } parent) {
            if (parent is YamlSequence sequence) {
                steps.Add("[" + sequence.IndexOf(current).ToString(CultureInfo.InvariantCulture) + "]");
            } else {
                var mapping = (YamlMapping)parent;
                int index = mapping.IndexOfNode(current, out bool isKey);
                YamlNode key = mapping.EntryAt(index).Key;
                if (isKey) throw new InvalidOperationException("No path addresses a mapping key; paths select values by key text.");
                // A collection key is addressed by its YAML text, as OfficeIMO.Reader.Yaml reports it.
                string text = key.Resolved is YamlScalar scalar ? scalar.Value : key.Resolved.ToString();
                steps.Add(IsName(text) ? "." + text : "[" + Quote(text) + "]");
            }
            current = parent;
        }
        var path = new StringBuilder();
        YamlDocument? document = current.Document;
        if (document is not null && document.Documents.Count > 1) {
            for (int i = 0; i < document.Documents.Count; i++) {
                if (ReferenceEquals(document.Documents[i].RootNode, current)) {
                    path.Append("$[").Append(i.ToString(CultureInfo.InvariantCulture)).Append(']');
                    break;
                }
            }
        }
        for (int i = steps.Count - 1; i >= 0; i--) {
            string step = steps[i];
            if (path.Length == 0 && step[0] == '.') path.Append(step, 1, step.Length - 1);
            else path.Append(step);
        }
        return path.Length == 0 ? "$" : path.ToString();
    }

    /// <summary>Appends a key step to a path in the grammar <see cref="Format"/> writes; an empty path is the root.</summary>
    public static string AppendKey(string path, string key) {
        if (!IsName(key)) return path + "[" + Quote(key) + "]";
        return path.Length == 0 ? key : path + "." + key;
    }

    /// <summary>Appends an index step to a path.</summary>
    public static string AppendIndex(string path, int index) => path + "[" + index.ToString(CultureInfo.InvariantCulture) + "]";

    private static bool IsName(string text) {
        if (text.Length == 0 || text[0] == '$') return false;
        foreach (char c in text) {
            if (c == '.' || c == '[' || c == ']' || c == '"' || c == '\\' || char.IsWhiteSpace(c) || char.IsControl(c)) return false;
        }
        return true;
    }

    private static string Quote(string text) {
        var quoted = new StringBuilder(text.Length + 2).Append('"');
        foreach (char c in text) {
            switch (c) {
                case '"': quoted.Append("\\\""); break;
                case '\\': quoted.Append("\\\\"); break;
                case '\n': quoted.Append("\\n"); break;
                case '\r': quoted.Append("\\r"); break;
                case '\t': quoted.Append("\\t"); break;
                default:
                    if (c < 0x20) quoted.Append("\\u").Append(((int)c).ToString("X4", CultureInfo.InvariantCulture));
                    else quoted.Append(c);
                    break;
            }
        }
        return quoted.Append('"').ToString();
    }

    private static YamlNode? Key(YamlNode? node, string key) {
        if (node?.Resolved is not YamlMapping mapping) return null;
        if (mapping.TryGetValue(key, out YamlNode? value)) return value;
        int index = mapping.IndexOfCollectionKey(key);
        return index >= 0 ? mapping.EntryAt(index).Value : null;
    }

    /// <summary>The mapping's own entry a key step names: a scalar key by its text, otherwise a collection key by its YAML text.</summary>
    private static int OwnIndex(YamlMapping mapping, string key) {
        int index = mapping.IndexOf(key);
        return index >= 0 ? index : mapping.IndexOfCollectionKey(key);
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
