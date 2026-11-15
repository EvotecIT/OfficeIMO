using System.Collections.Specialized;
using OfficeIMO.Patching;

namespace OfficeIMO.Toml;

/// <summary>
/// RFC 6902 JSON Patch and RFC 7386 JSON Merge Patch through the editing operations, so a patched document saves as a
/// minimal diff. Pointers (RFC 6901) address tables by key and arrays and arrays of tables by index.
/// </summary>
/// <remarks>
/// <para>A patch is all or nothing: it runs first on a copy loaded from the document's saved text, and only when every
/// operation succeeds there does it run on the document.</para>
/// <para>TOML has no null, so a JSON null anywhere in a value fails the patch. Objects become <c>[header]</c> tables
/// and arrays of objects <c>[[header]]</c> arrays of tables outside inline tables, and inline tables inside them and in
/// arrays. <c>move</c> and <c>copy</c> carry the value, with date-times kept as date-times; <c>test</c> compares a
/// date-time with a string by its RFC 3339 text.</para>
/// </remarks>
internal sealed class TomlPatch {
    private const int MaxDepth = 256;
    private const int CompareBudget = 1_000_000;

    private readonly TomlDocument _document;
    private readonly List<string> _added = new List<string>();
    private readonly List<string> _replaced = new List<string>();
    private readonly List<string> _removed = new List<string>();
    private readonly List<OfficeDiagnostic> _diagnostics = new List<OfficeDiagnostic>();

    private TomlPatch(TomlDocument document) => _document = document;

    // ------------------------------------------------------------------ entry points

    public static TomlPatchReport ApplyJsonPatch(TomlDocument document, string patch) {
        if (patch is null) throw new ArgumentNullException(nameof(patch));
        List<Operation> operations = ReadOperations(TomlPatchJson.Parse(patch));
        return Atomically(document, target => {
            var run = new TomlPatch(target);
            foreach (Operation operation in operations) run.Apply(operation);
            return run.Report();
        });
    }

    public static TomlPatchReport ApplyMergePatch(TomlDocument document, string patch) {
        if (patch is null) throw new ArgumentNullException(nameof(patch));
        object? value = TomlPatchJson.Parse(patch);
        if (value is not OrderedDictionary members) {
            throw new TomlPatchException(-1, null, "A merge patch for TOML is an object, since the document root is a table; this patch is " + Describe(value) + ".", null);
        }
        return Atomically(document, target => {
            var run = new TomlPatch(target);
            try {
                run.MergePatch(members);
            } catch (Exception exception) when (exception is InvalidOperationException or ArgumentException) {
                throw new TomlPatchException(-1, null, "The merge patch cannot be applied: " + exception.Message, exception);
            }
            return run.Report();
        });
    }

    /// <summary>Runs a patch on a copy of the document, then on the document, so a failure changes nothing.</summary>
    private static TomlPatchReport Atomically(TomlDocument document, Func<TomlDocument, TomlPatchReport> apply) {
        var options = new TomlLoadOptions { MaxDepth = int.MaxValue, MaxNodes = int.MaxValue, MaxStringLength = int.MaxValue };
        using (TomlDocument copy = TomlDocument.Parse(document.ToString(), options)) apply(copy);
        return apply(document);
    }

    private TomlPatchReport Report() => new TomlPatchReport(_added, _replaced, _removed, _diagnostics);

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
        public object? Value { get; set; }
    }

    private static List<Operation> ReadOperations(object? patch) {
        if (patch is not List<object?> list) throw new TomlPatchException(-1, null, "A JSON Patch is an array of operations.", null);
        var operations = new List<Operation>(list.Count);
        for (int i = 0; i < list.Count; i++) {
            if (list[i] is not OrderedDictionary item) throw new TomlPatchException(i, null, "The operation is not an object.", null);
            string name = Member(item, "op", i, null);
            string path = Member(item, "path", i, null);
            if (!JsonPatchOperations.TryParse(name, out JsonPatchOperationKind kind)) {
                throw new TomlPatchException(i, path, "'" + name + "' is not an operation; use add, remove, replace, move, copy or test.", null);
            }
            var operation = new Operation(i, kind, path, Pointer(path, i, path));
            if (JsonPatchOperations.NeedsFrom(kind)) operation.From = Pointer(Member(item, "from", i, path), i, path);
            if (JsonPatchOperations.NeedsValue(kind)) {
                if (!item.Contains("value")) throw new TomlPatchException(i, path, "The '" + name + "' operation has no 'value' member.", null);
                operation.Value = item["value"];
            }
            operations.Add(operation);
        }
        return operations;
    }

    private static string Member(OrderedDictionary item, string name, int index, string? path) {
        if (!item.Contains(name)) throw new TomlPatchException(index, path, "The operation has no '" + name + "' member.", null);
        return item[name] as string ?? throw new TomlPatchException(index, path, "The operation's '" + name + "' member is not a string.", null);
    }

    private static string[] Pointer(string pointer, int index, string path) =>
        JsonPointer.Parse(pointer, out string? error) ?? throw new TomlPatchException(index, path, "'" + pointer + "' is not a JSON Pointer: " + error + ".", null);

    // ------------------------------------------------------------------ applying a JSON Patch

    private void Apply(Operation operation) {
        try {
            switch (operation.Kind) {
                case JsonPatchOperationKind.Add:
                    Add(operation, operation.Tokens, operation.Value);
                    break;
                case JsonPatchOperationKind.Remove:
                    Remove(operation, operation.Tokens);
                    break;
                case JsonPatchOperationKind.Replace:
                    if (Find(operation.Tokens, operation.Tokens.Length) is null) throw Fail(operation, "There is no value to replace at '" + operation.Path + "'.");
                    Add(operation, operation.Tokens, operation.Value);
                    break;
                case JsonPatchOperationKind.Move: {
                    string[] from = operation.From!;
                    TomlNode source = Find(from, from.Length) ?? throw Fail(operation, "There is no value to move at the 'from' location.");
                    if (JsonPointer.AreEqual(from, operation.Tokens)) break;
                    if (JsonPointer.IsProperPrefix(from, operation.Tokens)) throw Fail(operation, "A value cannot be moved into itself.");
                    object value = ToValue(source, 0);
                    Remove(operation, from);
                    Add(operation, operation.Tokens, value);
                    break;
                }
                case JsonPatchOperationKind.Copy: {
                    TomlNode source = Find(operation.From!, operation.From!.Length) ?? throw Fail(operation, "There is no value to copy at the 'from' location.");
                    Add(operation, operation.Tokens, ToValue(source, 0));
                    break;
                }
                default: {
                    TomlNode actual = Find(operation.Tokens, operation.Tokens.Length) ?? throw Fail(operation, "There is no value to test at '" + operation.Path + "'.");
                    int budget = CompareBudget;
                    if (!JsonEqual(ToValue(actual, 0), operation.Value, ref budget, 0)) throw Fail(operation, "The value at '" + operation.Path + "' is not the tested value.");
                    break;
                }
            }
        } catch (Exception exception) when (exception is InvalidOperationException or ArgumentException) {
            throw Fail(operation, exception.Message, exception);
        }
    }

    /// <summary>
    /// Puts a value at a pointer: a table key is set, an array index inserts and <c>-</c> appends, an array of tables
    /// takes a table at its end, and the empty pointer replaces the root table's contents.
    /// </summary>
    private void Add(Operation operation, string[] tokens, object? value) {
        if (tokens.Length == 0) {
            if (value is not OrderedDictionary members) throw Fail(operation, "The root of a TOML document is a table; " + Describe(value) + " cannot replace it.");
            Refill(_document.Root, members);
            _replaced.Add("$");
            return;
        }
        TomlNode parent = Find(tokens, tokens.Length - 1) ?? throw Fail(operation, "The location that would hold '" + operation.Path + "' does not exist.");
        string last = tokens[tokens.Length - 1];
        string at = PathOf(tokens, tokens.Length - 1);
        bool replace = operation.Kind == JsonPatchOperationKind.Replace;
        switch (parent) {
            case TomlTable table: {
                bool existed = table.ContainsKey(last);
                Put(table, last, value);
                (existed ? _replaced : _added).Add(TomlPath.AppendKey(at, last));
                return;
            }
            case TomlArray array: {
                int index = Index(operation, last, array.Count, replace);
                if (replace) array.Set(index, Required(value));
                else array.Insert(index, Required(value));
                (replace ? _replaced : _added).Add(TomlPath.AppendIndex(at, index));
                return;
            }
            case TomlTableArray tables: {
                int index = Index(operation, last, tables.Count, replace);
                if (value is not OrderedDictionary members) throw Fail(operation, "An array of tables holds only tables; " + Describe(value) + " cannot be added to it.");
                if (replace) {
                    Refill(tables[index], members);
                } else {
                    if (index != tables.Count) throw Fail(operation, "A table can only be appended to an array of tables; use '-' or " + tables.Count.ToString(CultureInfo.InvariantCulture) + ".");
                    Refill(tables.AddTable(), members);
                }
                (replace ? _replaced : _added).Add(TomlPath.AppendIndex(at, index));
                return;
            }
            default:
                throw Fail(operation, "The location that would hold '" + operation.Path + "' is a value.");
        }
    }

    private int Index(Operation operation, string token, int count, bool replace) {
        if (JsonPointer.TryGetIndex(token, count, replace ? count : count + 1, allowEnd: !replace, out int index)) return index;
        throw Fail(operation, "'" + token + "' is not an index " + (replace ? "of" : "for") + " the array at '" + operation.Path + "', which has " + count.ToString(CultureInfo.InvariantCulture) + " items.");
    }

    private void Remove(Operation operation, string[] tokens) {
        if (tokens.Length == 0) throw Fail(operation, "The root table cannot be removed; replace it instead.");
        TomlNode parent = Find(tokens, tokens.Length - 1) ?? throw Fail(operation, "There is no value to remove at '" + Format(tokens) + "'.");
        string last = tokens[tokens.Length - 1];
        string at = PathOf(tokens, tokens.Length - 1);
        int index;
        switch (parent) {
            case TomlTable table when RemoveKey(table, last):
                _removed.Add(TomlPath.AppendKey(at, last));
                return;
            case TomlArray array when JsonPointer.TryGetIndex(last, array.Count, array.Count, allowEnd: false, out index):
                array.RemoveAt(index);
                _removed.Add(TomlPath.AppendIndex(at, index));
                return;
            case TomlTableArray tables when JsonPointer.TryGetIndex(last, tables.Count, tables.Count, allowEnd: false, out index):
                tables.RemoveAt(index);
                // An array of tables with no tables has no text, so the key would vanish on save: it becomes [].
                if (tables.Count == 0 && tables.Parent is TomlTable owner && tables.Key is { } key) owner.PutNode(key, new TomlArray(_document));
                _removed.Add(TomlPath.AppendIndex(at, index));
                return;
            default:
                throw Fail(operation, "There is no value to remove at '" + Format(tokens) + "'.");
        }
    }

    /// <summary>
    /// Removes a key and keeps the tables around it. <see cref="TomlTable.Remove"/> drops a table made only by dotted keys
    /// or deeper headers once it is empty, since it has no text of its own; a patch removes only the member it names,
    /// so such a table that lost its last key is added back as an empty <c>[header]</c> table.
    /// </summary>
    private static bool RemoveKey(TomlTable table, string key) {
        // The tables between the root and this one, innermost first.
        var chain = new List<TomlTable>();
        for (TomlTable current = table; current.Parent is TomlTable parent; current = parent) chain.Add(current);
        if (!table.Remove(key)) return false;
        int top = -1;
        for (int i = 0; i < chain.Count; i++) {
            if (chain[i].Parent is null) top = i;
        }
        if (top < 0) return true;
        // The outermost table the removal dropped still knows its parent, which stayed; rebuild the chain below it.
        var restored = (TomlTable)chain[top]._parent!;
        for (int i = top; i >= 0; i--) restored = restored.AddTable(chain[i].Key!);
        return true;
    }

    private TomlNode? Find(string[] tokens, int count) {
        TomlNode node = _document.Root;
        for (int i = 0; i < count; i++) {
            TomlNode? next = null;
            int index;
            switch (node) {
                case TomlTable table:
                    table.TryGetValue(tokens[i], out next);
                    break;
                case TomlArray array when JsonPointer.TryGetIndex(tokens[i], array.Count, array.Count, allowEnd: false, out index):
                    next = array[index];
                    break;
                case TomlTableArray tables when JsonPointer.TryGetIndex(tokens[i], tables.Count, tables.Count, allowEnd: false, out index):
                    next = tables[index];
                    break;
            }
            if (next is null) return null;
            node = next;
        }
        return node;
    }

    /// <summary>The path of the node <paramref name="count"/> tokens address, in the grammar <see cref="TomlNode.Select"/> reads.</summary>
    private string PathOf(string[] tokens, int count) {
        string path = string.Empty;
        TomlNode node = _document.Root;
        for (int i = 0; i < count; i++) {
            if (node is TomlTable table) {
                path = TomlPath.AppendKey(path, tokens[i]);
                node = table[tokens[i]];
            } else {
                int index = int.Parse(tokens[i], NumberStyles.None, CultureInfo.InvariantCulture);
                path = TomlPath.AppendIndex(path, index);
                node = node is TomlArray array ? array[index] : ((TomlTableArray)node)[index];
            }
        }
        return path;
    }

    // ------------------------------------------------------------------ applying a JSON Merge Patch

    /// <summary>
    /// Applies RFC 7386 to the root table: nulls remove keys, objects over tables merge, and everything else replaces.
    /// The removals run here; the rest, with its nulls left out, merges as an overlay through <see cref="TomlMerge"/>, so
    /// equal values keep their spelling and replaced arrays change only the items that differ.
    /// </summary>
    private void MergePatch(OrderedDictionary patch) {
        RemoveNulls(_document.Root, patch, string.Empty, 0);
        using TomlDocument overlay = TomlDocument.Create();
        foreach (DictionaryEntry member in patch) {
            if (member.Value is not null) Build(overlay.Root, (string)member.Key, WithoutNulls(member.Value, 0));
        }
        TomlMergeReport merged = TomlMerge.Merge(_document, overlay, new TomlMergeOptions { MergePatch = true });
        _added.AddRange(merged.Added);
        _replaced.AddRange(merged.Replaced);
        _diagnostics.AddRange(merged.Diagnostics);
    }

    private void RemoveNulls(TomlTable table, OrderedDictionary patch, string path, int depth) {
        if (depth > MaxDepth) throw new InvalidOperationException("The patch nests more than " + MaxDepth + " levels deep.");
        foreach (DictionaryEntry member in patch) {
            string key = (string)member.Key;
            if (member.Value is null) {
                if (RemoveKey(table, key)) _removed.Add(TomlPath.AppendKey(path, key));
            } else if (member.Value is OrderedDictionary inner && table.TryGetValue(key, out TomlNode? node) && node is TomlTable child) {
                RemoveNulls(child, inner, TomlPath.AppendKey(path, key), depth + 1);
            }
        }
    }

    /// <summary>Leaves out null members of objects, and of the objects they hold; an array keeps its items as they are.</summary>
    private static object WithoutNulls(object value, int depth) {
        if (value is not OrderedDictionary members) return value;
        if (depth > MaxDepth) throw new InvalidOperationException("The patch nests more than " + MaxDepth + " levels deep.");
        var copy = new OrderedDictionary(StringComparer.Ordinal);
        foreach (DictionaryEntry member in members) {
            if (member.Value is not null) copy.Add(member.Key, WithoutNulls(member.Value, depth + 1));
        }
        return copy;
    }

    // ------------------------------------------------------------------ values

    /// <summary>
    /// Sets a key to a JSON value. Objects, and arrays of objects outside inline tables, are built in their TOML syntax
    /// and copied in; other values go through <see cref="TomlTable.Set"/>, which keeps an equal value's spelling.
    /// </summary>
    private void Put(TomlTable table, string key, object? value) {
        value = Required(value);
        if (value is OrderedDictionary || (IsTableList(value) && !table.IsInlineScope())) {
            using TomlDocument scratch = TomlDocument.Create();
            Build(scratch.Root, key, value);
            table.PutNode(key, TomlMerge.CopyNode(_document, scratch.Root[key], table.IsInlineScope()));
            return;
        }
        table.Set(key, value);
    }

    /// <summary>Replaces a table's entries with an object's members.</summary>
    private void Refill(TomlTable table, OrderedDictionary members) {
        var keys = new List<string>(table.Keys);
        foreach (string key in keys) {
            if (!members.Contains(key)) table.Remove(key);
        }
        foreach (DictionaryEntry member in members) Put(table, (string)member.Key, member.Value);
    }

    /// <summary>Adds a JSON value under a key of a table being built: objects as tables, arrays of objects as arrays of tables.</summary>
    private static void Build(TomlTable table, string key, object? value) {
        value = Required(value);
        if (value is OrderedDictionary members) {
            TomlTable child = table.AddTable(key);
            foreach (DictionaryEntry member in members) Build(child, (string)member.Key, member.Value);
        } else if (IsTableList(value) && !table.IsInlineScope()) {
            TomlTableArray tables = table.AddTableArray(key);
            foreach (object? item in (List<object?>)value) {
                TomlTable child = tables.AddTable();
                foreach (DictionaryEntry member in (OrderedDictionary)item!) Build(child, (string)member.Key, member.Value);
            }
        } else {
            table.Add(key, value);
        }
    }

    private static bool IsTableList(object value) {
        if (value is not List<object?> items || items.Count == 0) return false;
        foreach (object? item in items) {
            if (item is not OrderedDictionary) return false;
        }
        return true;
    }

    private static object Required(object? value) =>
        value ?? throw new InvalidOperationException("TOML has no null value, so a null cannot be written; remove the key instead.");

    /// <summary>Reads a node as a JSON-like value: dictionaries for tables, lists for arrays, date-times kept as they are.</summary>
    private static object ToValue(TomlNode node, int depth) {
        if (depth > MaxDepth) throw new InvalidOperationException("The value nests more than " + MaxDepth + " levels deep.");
        switch (node) {
            case TomlValue value:
                return value.Value;
            case TomlTable table: {
                var members = new OrderedDictionary(StringComparer.Ordinal);
                foreach (KeyValuePair<string, TomlNode> entry in table) members.Add(entry.Key, ToValue(entry.Value, depth + 1));
                return members;
            }
            case TomlArray array: {
                var items = new List<object?>(array.Count);
                foreach (TomlNode item in array) items.Add(ToValue(item, depth + 1));
                return items;
            }
            default: {
                var tables = (TomlTableArray)node;
                var items = new List<object?>(tables.Count);
                foreach (TomlTable item in tables) items.Add(ToValue(item, depth + 1));
                return items;
            }
        }
    }

    /// <summary>
    /// Compares two values as JSON does: objects regardless of member order, arrays item by item, numbers by value, and a
    /// date-time with a string by its RFC 3339 text.
    /// </summary>
    private static bool JsonEqual(object? left, object? right, ref int budget, int depth) {
        if (--budget < 0 || depth > MaxDepth) throw new InvalidOperationException("The values are too large to compare.");
        switch (left) {
            case null:
                return right is null;
            case long number when right is long otherNumber:
                return number == otherNumber;
            case double number when right is double otherNumber:
                return number.Equals(otherNumber);
            case long number when right is double otherNumber:
                return JsonNumbers.IntegerEquals(number, otherNumber);
            case double number when right is long otherNumber:
                return JsonNumbers.IntegerEquals(otherNumber, number);
            case string text:
                return right is string otherText ? string.Equals(text, otherText, StringComparison.Ordinal) : right is TomlDateTime rightDate && text == rightDate.ToString();
            case TomlDateTime date:
                return right is TomlDateTime otherDate ? date == otherDate : right is string rightText && rightText == date.ToString();
            case bool flag:
                return right is bool otherFlag && flag == otherFlag;
            case OrderedDictionary members when right is OrderedDictionary otherMembers:
                if (members.Count != otherMembers.Count) return false;
                foreach (DictionaryEntry member in members) {
                    if (!otherMembers.Contains(member.Key) || !JsonEqual(member.Value, otherMembers[member.Key], ref budget, depth + 1)) return false;
                }
                return true;
            case List<object?> items when right is List<object?> otherItems:
                if (items.Count != otherItems.Count) return false;
                for (int i = 0; i < items.Count; i++) {
                    if (!JsonEqual(items[i], otherItems[i], ref budget, depth + 1)) return false;
                }
                return true;
            default:
                return false;
        }
    }

    private static string Describe(object? value) => value switch {
        null => "null",
        OrderedDictionary => "an object",
        List<object?> => "an array",
        string => "a string",
        bool => "a boolean",
        _ => "a number"
    };

    private static string Format(string[] tokens) {
        var pointer = new StringBuilder();
        foreach (string token in tokens) pointer.Append('/').Append(token.Replace("~", "~0").Replace("/", "~1"));
        return pointer.ToString();
    }

    private static TomlPatchException Fail(Operation operation, string message, Exception? inner = null) =>
        inner as TomlPatchException ?? new TomlPatchException(operation.Index, operation.Path, message, inner);
}
