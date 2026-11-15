using OfficeIMO.Patching;

namespace OfficeIMO.Yaml;

/// <summary>
/// RFC 6902 JSON Patch and RFC 7386 JSON Merge Patch through the editing operations, so a patched document saves as a
/// minimal diff. Pointers (RFC 6901) address the first document of the stream through its typed view: steps follow
/// aliases and keys that <c>&lt;&lt;</c> merges provide, as <see cref="YamlNode.Select"/> does.
/// </summary>
/// <remarks>
/// <para>A patch is all or nothing. It runs first on a copy loaded from the document's saved text; only when every
/// operation succeeds there does it run on the document, so a failed test, a pointer to nothing or an edit the model
/// refuses leaves the document as it was.</para>
/// <para>Values from the patch are copied as their typed view: aliases are expanded, <c>&lt;&lt;</c> merges flattened and
/// anchors left out, so JSON and YAML patches mean the same thing. New collections take block layout, or flow layout
/// inside a flow collection; strings take the document's quoting. <c>move</c> moves the node itself, with its comments
/// and spelling.</para>
/// </remarks>
internal sealed class YamlPatch {
    private const int CompareBudget = 1_000_000;
    private const int MaxDepth = 1000;

    /// <summary>
    /// How patch text loads: a repeated member name fails, since RFC 6902 leaves an operation such as
    /// <c>{"op": "add", ..., "op": "remove"}</c> without a meaning.
    /// </summary>
    internal static YamlLoadOptions TextOptions { get; } = new YamlLoadOptions { DuplicateKeys = YamlDuplicateKeyPolicy.Error };

    private readonly YamlDocument _document;
    private readonly int _diagnosticsBefore;
    private readonly List<string> _added = new List<string>();
    private readonly List<string> _replaced = new List<string>();
    private readonly List<string> _removed = new List<string>();

    private YamlPatch(YamlDocument document) {
        _document = document;
        _diagnosticsBefore = document.Diagnostics.Count;
    }

    // ------------------------------------------------------------------ entry points

    public static YamlPatchReport ApplyJsonPatch(YamlDocument document, YamlDocument patch) {
        if (patch is null) throw new ArgumentNullException(nameof(patch));
        if (ReferenceEquals(document, patch)) throw new ArgumentException("A document cannot be patched with itself.", nameof(patch));
        List<Operation> operations = ReadOperations(patch);
        return Atomically(document, target => {
            var run = new YamlPatch(target);
            foreach (Operation operation in operations) run.Apply(operation);
            return run.Report();
        });
    }

    public static YamlPatchReport ApplyMergePatch(YamlDocument document, YamlDocument patch) {
        if (patch is null) throw new ArgumentNullException(nameof(patch));
        if (ReferenceEquals(document, patch)) throw new ArgumentException("A document cannot be patched with itself.", nameof(patch));
        if (patch.Documents.Count > 1) throw new YamlPatchException(-1, null, "A merge patch is one document; this one has " + patch.Documents.Count.ToString(CultureInfo.InvariantCulture) + ".", null);
        // An empty stream (an empty values file) changes nothing; a document holding null replaces the target with null.
        if (patch.Documents.Count == 0) return new YamlPatchReport(new List<string>(), new List<string>(), new List<string>(), Array.Empty<OfficeDiagnostic>());
        YamlNode value = patch.Documents[0].Root;
        return Atomically(document, target => {
            var run = new YamlPatch(target);
            try {
                run.MergePatch(value);
            } catch (Exception exception) when (exception is InvalidOperationException or ArgumentException) {
                throw new YamlPatchException(-1, null, "The merge patch cannot be applied: " + exception.Message, exception);
            }
            return run.Report();
        });
    }

    /// <summary>Runs a patch on a copy of the document, then on the document, so a failure changes nothing.</summary>
    private static YamlPatchReport Atomically(YamlDocument document, Func<YamlDocument, YamlPatchReport> apply) {
        using (YamlDocument copy = document.LoadCopy()) apply(copy);
        return apply(document);
    }

    private YamlPatchReport Report() {
        var diagnostics = new List<OfficeDiagnostic>();
        IReadOnlyList<OfficeDiagnostic> all = _document.Diagnostics;
        for (int i = _diagnosticsBefore; i < all.Count; i++) diagnostics.Add(all[i]);
        return new YamlPatchReport(_added, _replaced, _removed, diagnostics);
    }

    // ------------------------------------------------------------------ reading a JSON Patch

    private sealed class Operation {
        public Operation(int index, JsonPatchOperationKind kind, string path, string[] tokens) {
            Index = index;
            Kind = kind;
            Path = path;
            Tokens = tokens;
        }

        public int Index { get; }
        public JsonPatchOperationKind Kind { get; }
        public string Path { get; }
        public string[] Tokens { get; }
        public string[]? From { get; set; }
        public YamlNode? Value { get; set; }
    }

    private static List<Operation> ReadOperations(YamlDocument patch) {
        if (patch.Documents.Count != 1 || patch.Documents[0].Root.Resolved is not YamlSequence list) {
            throw new YamlPatchException(-1, null, "A JSON Patch is one document holding an array of operations.", null);
        }
        var operations = new List<Operation>(list._count);
        for (int i = 0; i < list._count; i++) {
            if (list.ItemAt(i).Resolved is not YamlMapping item) throw new YamlPatchException(i, null, "The operation is not an object.", null);
            string name = Member(item, "op", i, null);
            string path = Member(item, "path", i, null);
            if (!JsonPatchOperations.TryParse(name, out JsonPatchOperationKind kind)) {
                throw new YamlPatchException(i, path, "'" + name + "' is not an operation; use add, remove, replace, move, copy or test.", null);
            }
            var operation = new Operation(i, kind, path, Pointer(path, i, path));
            if (JsonPatchOperations.NeedsFrom(kind)) {
                string from = Member(item, "from", i, path);
                operation.From = Pointer(from, i, path);
            }
            if (JsonPatchOperations.NeedsValue(kind)) {
                if (!item.TryGetValue("value", out YamlNode? value)) throw new YamlPatchException(i, path, "The '" + name + "' operation has no 'value' member.", null);
                operation.Value = value;
            }
            operations.Add(operation);
        }
        return operations;
    }

    private static string Member(YamlMapping item, string name, int index, string? path) {
        if (!item.TryGetValue(name, out YamlNode? node)) throw new YamlPatchException(index, path, "The operation has no '" + name + "' member.", null);
        if (node!.Resolved is YamlScalar scalar && scalar.TryGetKind(out YamlScalarKind kind) && kind == YamlScalarKind.String) return scalar.Value;
        throw new YamlPatchException(index, path, "The operation's '" + name + "' member is not a string.", null);
    }

    private static string[] Pointer(string pointer, int index, string path) =>
        JsonPointer.Parse(pointer, out string? error) ?? throw new YamlPatchException(index, path, "'" + pointer + "' is not a JSON Pointer: " + error + ".", null);

    // ------------------------------------------------------------------ applying a JSON Patch

    private void Apply(Operation operation) {
        try {
            switch (operation.Kind) {
                case JsonPatchOperationKind.Add:
                    Add(operation, operation.Tokens, operation.Value!, moved: null);
                    break;
                case JsonPatchOperationKind.Remove:
                    Remove(operation, operation.Tokens);
                    break;
                case JsonPatchOperationKind.Replace:
                    if (Find(operation.Tokens, operation.Tokens.Length) is null) throw Fail(operation, "There is no value to replace at '" + operation.Path + "'.");
                    Add(operation, operation.Tokens, operation.Value!, moved: null);
                    break;
                case JsonPatchOperationKind.Move:
                    Move(operation);
                    break;
                case JsonPatchOperationKind.Copy: {
                    YamlNode source = Find(operation.From!, operation.From!.Length) ?? throw Fail(operation, "There is no value to copy at the 'from' location.");
                    Add(operation, operation.Tokens, source, moved: null);
                    break;
                }
                default: {
                    YamlNode actual = Find(operation.Tokens, operation.Tokens.Length) ?? throw Fail(operation, "There is no value to test at '" + operation.Path + "'.");
                    int budget = CompareBudget;
                    if (!JsonEqual(actual, operation.Value!, ref budget, 0)) throw Fail(operation, "The value at '" + operation.Path + "' is not the tested value.");
                    break;
                }
            }
        } catch (Exception exception) when (exception is InvalidOperationException or ArgumentException) {
            throw Fail(operation, exception.Message, exception);
        }
    }

    private void Move(Operation operation) {
        string[] from = operation.From!;
        if (JsonPointer.AreEqual(from, operation.Tokens)) {
            if (Find(from, from.Length) is null) throw Fail(operation, "There is no value to move at the 'from' location.");
            return;
        }
        if (JsonPointer.IsProperPrefix(from, operation.Tokens)) throw Fail(operation, "A value cannot be moved into itself.");
        YamlNode node = Remove(operation, from);
        Add(operation, operation.Tokens, node, moved: node);
        CheckAliasOrder(operation);
    }

    /// <summary>
    /// Puts a value at a pointer: a copy of <paramref name="source"/>, or the <paramref name="moved"/> node itself. A key
    /// is set (a key that only a merge provides gets an entry of its own), a sequence index inserts, <c>-</c> appends and
    /// the empty pointer replaces the document root.
    /// </summary>
    private void Add(Operation operation, string[] tokens, YamlNode source, YamlNode? moved) {
        if (tokens.Length == 0) {
            YamlNode root = moved ?? Copy(source, flow: false, comments: true, 0);
            bool existed = _document.Documents.Count > 0;
            if (existed) _document.Documents[0].Root = root;
            else _document.AddDocument(root);
            (existed ? _replaced : _added).Add(Display(YamlPath.Format(root)));
            return;
        }
        YamlNode parent = Find(tokens, tokens.Length - 1, forEdit: true) ?? throw Fail(operation, "The location that would hold '" + operation.Path + "' does not exist.");
        string last = tokens[tokens.Length - 1];
        YamlNode container = parent.Resolved;
        string at = ContainerPath(container);
        bool flow = container._style == YamlStyle.Flow;
        if (container is YamlMapping mapping) {
            bool existed = mapping.TryGetValue(last, out _);
            mapping.Set(last, moved ?? Copy(source, flow, comments: true, 0));
            (existed ? _replaced : _added).Add(Display(YamlPath.AppendKey(at, last)));
            return;
        }
        if (container is YamlSequence sequence) {
            bool replace = operation.Kind == JsonPatchOperationKind.Replace;
            int count = sequence._count;
            if (!JsonPointer.TryGetIndex(last, count, replace ? count : count + 1, allowEnd: !replace, out int index)) {
                throw Fail(operation, "'" + last + "' is not an index " + (replace ? "of" : "for") + " the sequence at '" + operation.Path + "', which has " + count.ToString(CultureInfo.InvariantCulture) + " items.");
            }
            YamlNode node = moved ?? Copy(source, flow, comments: true, 0);
            if (replace) sequence[index] = node;
            else sequence.Insert(index, node);
            (replace ? _replaced : _added).Add(Display(YamlPath.AppendIndex(at, index)));
            return;
        }
        throw Fail(operation, "The location that would hold '" + operation.Path + "' is a scalar.");
    }

    /// <summary>Removes the value a pointer addresses and returns it, detached.</summary>
    private YamlNode Remove(Operation operation, string[] tokens) {
        if (tokens.Length == 0) throw Fail(operation, "The document root cannot be removed; replace it instead.");
        YamlNode parent = Find(tokens, tokens.Length - 1, forEdit: true) ?? throw Fail(operation, "There is no value to remove at '" + Format(tokens) + "'.");
        string last = tokens[tokens.Length - 1];
        YamlNode container = parent.Resolved;
        string at = ContainerPath(container);
        if (container is YamlMapping mapping) {
            int own = mapping.IndexOf(last);
            if (own < 0) {
                throw Fail(operation, mapping.TryGetValue(last, out _)
                    ? "The key '" + last + "' comes from a << merge key and cannot be removed; merged mappings are never edited."
                    : "There is no value to remove at '" + Format(tokens) + "'.");
            }
            YamlNode value = mapping.EntryAt(own).Value;
            mapping.RemoveAt(own);
            _removed.Add(Display(YamlPath.AppendKey(at, last)));
            return value;
        }
        if (container is YamlSequence sequence && JsonPointer.TryGetIndex(last, sequence._count, sequence._count, allowEnd: false, out int index)) {
            YamlNode item = sequence.ItemAt(index);
            sequence.RemoveAt(index);
            _removed.Add(Display(YamlPath.AppendIndex(at, index)));
            return item;
        }
        throw Fail(operation, "There is no value to remove at '" + Format(tokens) + "'.");
    }

    /// <summary>
    /// Fails a move that left an alias before its anchor, where a reader could not resolve it. Only a move can: copies and
    /// added values have no aliases, and removing or replacing an anchored node that aliases use already fails.
    /// </summary>
    private void CheckAliasOrder(Operation operation) {
        if (_document.Documents.Count == 0) return;
        var anchored = new HashSet<YamlNode>();
        foreach (YamlNode node in _document.Documents[0].Root.Descendants()) {
            if (node is YamlAlias alias) {
                if (!anchored.Contains(alias.Target)) throw Fail(operation, "The move would put the alias *" + alias.Name + " before its anchor.");
            } else if (node.Anchor is not null) {
                anchored.Add(node);
            }
        }
    }

    /// <summary>
    /// Follows <paramref name="count"/> tokens from the first document's root; <see langword="null"/> when nothing is
    /// there. With <paramref name="forEdit"/>, a collection that only a <c>&lt;&lt;</c> merge provides is first copied into
    /// an entry of its own, so the edit below it changes this mapping and never the merged one, which others share.
    /// </summary>
    private YamlNode? Find(string[] tokens, int count, bool forEdit = false) {
        if (_document.Documents.Count == 0) return null;
        YamlNode node = _document.Documents[0].Root;
        for (int i = 0; i < count; i++) {
            YamlNode container = node.Resolved;
            YamlNode? next = null;
            if (container is YamlMapping mapping) {
                int own = mapping.IndexOf(tokens[i]);
                if (own >= 0) {
                    next = mapping.EntryAt(own).Value;
                } else if (mapping.TryGetValue(tokens[i], out YamlNode? merged)) {
                    next = merged;
                    if (forEdit && merged!.Resolved is YamlMapping or YamlSequence) {
                        string at = ContainerPath(mapping);
                        next = mapping.Set(tokens[i], Copy(merged, mapping._style == YamlStyle.Flow, comments: false, 0));
                        _added.Add(Display(YamlPath.AppendKey(at, tokens[i])));
                    }
                }
            } else if (container is YamlSequence sequence && JsonPointer.TryGetIndex(tokens[i], sequence._count, sequence._count, allowEnd: false, out int index)) {
                next = sequence.ItemAt(index);
            }
            if (next is null) return null;
            node = next;
        }
        return node;
    }

    // ------------------------------------------------------------------ applying a JSON Merge Patch

    /// <summary>
    /// Applies RFC 7386: a patch object merges into the root mapping (a root of another kind becomes an empty mapping
    /// first), nulls remove keys and other values replace; a patch of any other kind replaces the root.
    /// </summary>
    private void MergePatch(YamlNode patch) {
        YamlNode? root = _document.Documents.Count == 0 ? null : _document.Documents[0].Root;
        if (patch.Resolved is not YamlMapping patchMapping) {
            int budget = CompareBudget;
            if (root is not null && JsonEqual(root, patch, ref budget, 0)) return;
            YamlNode copy = Copy(patch, flow: false, comments: true, 0);
            if (root is null) _document.AddDocument(copy);
            else _document.Documents[0].Root = copy;
            (root is null ? _added : _replaced).Add(Display(YamlPath.Format(copy)));
            return;
        }
        if (root?.Resolved is not YamlMapping target) {
            target = new YamlMapping();
            if (root is null) _document.AddDocument(target);
            else _document.Documents[0].Root = target;
            (root is null ? _added : _replaced).Add(Display(YamlPath.Format(target)));
        }
        // The patch is read as a JSON value: a block-layout copy of its typed view, so JSON text adds block mappings. The
        // copy is written for this document's schema and held in a document read under it, so the merge reads every key
        // and value (010 or yes under YAML 1.1) as this document will.
        var overlay = (YamlMapping)Copy(patchMapping, flow: false, comments: true, 0);
        YamlMergeReport merged;
        using (YamlDocument holder = YamlDocument.Parse("{}\n", new YamlLoadOptions { Schema = Schema })) {
            holder.Documents[0].Root = overlay;
            merged = YamlMerge.Merge(target, overlay, new YamlMergeOptions { NullRemovesKey = true, MergePatch = true });
        }
        if (target._count == 0) target._style = YamlStyle.Flow;
        _added.AddRange(merged.Added);
        _replaced.AddRange(merged.Replaced);
        _removed.AddRange(merged.Removed);
    }

    // ------------------------------------------------------------------ values

    private YamlSchema Schema => _document.Documents.Count > 0 ? _document.Documents[0].Schema : YamlSchema.Core;

    /// <summary>
    /// Copies a value as its typed view, ready to join this document: aliases expanded, merges flattened, no anchors.
    /// Collections are block, or flow when <paramref name="flow"/> or empty.
    /// </summary>
    private YamlNode Copy(YamlNode node, bool flow, bool comments, int depth) {
        if (depth > MaxDepth) throw new InvalidOperationException("The patch value is nested deeper than " + MaxDepth + " levels.");
        YamlNode value = node.Resolved;
        YamlNode copy;
        switch (value) {
            case YamlScalar scalar:
                copy = YamlMerge.CopyScalar(scalar, Schema);
                break;
            case YamlMapping mapping: {
                List<YamlMerge.Entry> entries = YamlMerge.Effective(mapping);
                var target = new YamlMapping(flow || entries.Count == 0 ? YamlStyle.Flow : YamlStyle.Block);
                bool inner = target._style == YamlStyle.Flow;
                foreach (YamlMerge.Entry entry in entries) {
                    YamlNode key = entry.Key.Resolved is YamlScalar keyScalar ? YamlMerge.CopyScalar(keyScalar, Schema) : Copy(entry.Key, flow: true, comments: false, depth + 1);
                    if (comments && IsForeign(entry.Key)) YamlEdit.CopyComments(entry.Key, key);
                    YamlEdit.InsertEntry(target, target._count, key, Copy(entry.Value, inner, comments, depth + 1));
                }
                copy = target;
                break;
            }
            default: {
                var sequence = (YamlSequence)value;
                var target = new YamlSequence(flow || sequence._count == 0 ? YamlStyle.Flow : YamlStyle.Block);
                bool inner = target._style == YamlStyle.Flow;
                for (int i = 0; i < sequence._count; i++) YamlEdit.InsertItem(target, target._count, Copy(sequence.ItemAt(i), inner, comments, depth + 1));
                copy = target;
                break;
            }
        }
        if (value.Tag is { } tag && copy is not YamlScalar) (copy._properties ??= new YamlNodeProperties()).Tag = tag;
        if (comments && IsForeign(node)) YamlEdit.CopyComments(node, copy);
        return copy;
    }

    /// <summary>Whether a node's comments travel with its copy: those of patch documents do, a <c>copy</c> within this document's do not.</summary>
    private bool IsForeign(YamlNode node) => node._part is null || !ReferenceEquals(node._part._document, _document);

    /// <summary>
    /// Compares two values as JSON does: mappings by their typed view regardless of order, sequences item by item,
    /// numbers by value whether integer or float, and other scalars by kind and value.
    /// </summary>
    private static bool JsonEqual(YamlNode left, YamlNode right, ref int budget, int depth) {
        if (--budget < 0 || depth > MaxDepth) throw new InvalidOperationException("The values are too large to compare.");
        left = left.Resolved;
        right = right.Resolved;
        switch (left) {
            case YamlScalar scalar when right is YamlScalar other:
                YamlMerge.Data(scalar, out YamlScalarKind kind, out string text, out string? tag);
                YamlMerge.Data(other, out YamlScalarKind otherKind, out string otherText, out string? otherTag);
                if (kind is YamlScalarKind.Integer or YamlScalarKind.Float && otherKind is YamlScalarKind.Integer or YamlScalarKind.Float && kind != otherKind) {
                    // Integer text is canonical decimal (or the written digits beyond 64 bits); the float is compared exactly.
                    string integerText = kind == YamlScalarKind.Integer ? text : otherText, floatText = kind == YamlScalarKind.Integer ? otherText : text;
                    return System.Numerics.BigInteger.TryParse(integerText, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out System.Numerics.BigInteger integer)
                        && YamlSchemaResolver.TryParseDouble(floatText.AsSpan(), YamlSchema.Core, out double number)
                        && JsonNumbers.IntegerEquals(integer, number);
                }
                return kind == otherKind && text == otherText && tag == otherTag;
            case YamlSequence sequence when right is YamlSequence other:
                if (sequence._count != other._count) return false;
                for (int i = 0; i < sequence._count; i++) {
                    if (!JsonEqual(sequence.ItemAt(i), other.ItemAt(i), ref budget, depth + 1)) return false;
                }
                return true;
            case YamlMapping mapping when right is YamlMapping other:
                List<YamlMerge.Entry> entries = YamlMerge.Effective(mapping), otherEntries = YamlMerge.Effective(other);
                if (entries.Count != otherEntries.Count) return false;
                var lookup = new Dictionary<string, YamlNode>(StringComparer.Ordinal);
                foreach (YamlMerge.Entry entry in otherEntries) {
                    if (entry.Text is null) return false;
                    lookup[entry.Text] = entry.Value;
                }
                foreach (YamlMerge.Entry entry in entries) {
                    if (entry.Text is null || !lookup.TryGetValue(entry.Text, out YamlNode? value) || !JsonEqual(entry.Value, value, ref budget, depth + 1)) return false;
                }
                return true;
            default:
                return false;
        }
    }

    // ------------------------------------------------------------------ reporting

    private static string ContainerPath(YamlNode container) {
        string path = YamlPath.Format(container);
        return path == "$" ? string.Empty : path;
    }

    private static string Display(string path) => path.Length == 0 ? "$" : path;

    private static string Format(string[] tokens) {
        var pointer = new StringBuilder();
        foreach (string token in tokens) pointer.Append('/').Append(token.Replace("~", "~0").Replace("/", "~1"));
        return pointer.ToString();
    }

    private static YamlPatchException Fail(Operation operation, string message, Exception? inner = null) =>
        inner as YamlPatchException ?? new YamlPatchException(operation.Index, operation.Path, message, inner);
}
