namespace OfficeIMO.Toml;

/// <summary>
/// Deep merge of an overlay document into a base document through the editing operations, so the base saves as a
/// minimal diff. Tables merge key by key whatever their syntax, values are assigned (an equal value keeps the base
/// text), arrays and arrays of tables follow <see cref="TomlMergeOptions.Arrays"/>, and new entries are copies of the
/// overlay's nodes. The overlay is only read.
/// </summary>
/// <remarks>
/// <para>Copies take the syntax their new place allows: inside an inline table or array everything is inline and an
/// array of tables becomes an array of inline tables; elsewhere inline tables stay inline, tables made by dotted keys
/// stay dotted keys when they hold only key/value pairs, and other tables become <c>[header]</c> sections. The editing
/// layer places each of them, so the saved text always reloads to the merged data.</para>
/// <para>An array value and an array of tables are both arrays: their items merge as arrays, and a base array of tables
/// that would receive an item other than a table is replaced by the overlay's array.</para>
/// </remarks>
internal sealed class TomlMerge {
    private readonly TomlMergeOptions _options;
    private readonly TomlDocument _document;
    private readonly List<string> _added = new List<string>();
    private readonly List<string> _replaced = new List<string>();
    private readonly List<OfficeDiagnostic> _diagnostics = new List<OfficeDiagnostic>();
    /// <summary>Copied nodes and the overlay comment at the end of their line, set once they are in place.</summary>
    private readonly List<KeyValuePair<TomlNode, string>> _comments = new List<KeyValuePair<TomlNode, string>>();

    private TomlMerge(TomlDocument document, TomlMergeOptions? options) {
        _document = document;
        _options = options ?? TomlMergeOptions.Default;
    }

    public static TomlMergeReport Merge(TomlDocument target, TomlDocument overlay, TomlMergeOptions? options) {
        if (overlay is null) throw new ArgumentNullException(nameof(overlay));
        if (ReferenceEquals(target, overlay)) throw new ArgumentException("A document cannot be merged into itself.", nameof(overlay));
        var merge = new TomlMerge(target, options);
        merge.MergeTables(target._root, overlay._root, string.Empty);
        foreach (KeyValuePair<TomlNode, string> pending in merge._comments) {
            TomlNode node = pending.Key;
            if ((node._state & TomlNodeState.Detached) == 0 && node.CanHaveComment()) node.SetComment(pending.Value);
        }
        return new TomlMergeReport(merge._added, merge._replaced, merge._diagnostics);
    }

    // ------------------------------------------------------------------ deep merge

    private void MergeTables(TomlTable target, TomlTable overlay, string path) {
        for (int i = 0; i < overlay._count; i++) {
            TomlNode incoming = overlay._nodes[i];
            string key = incoming._key!;
            string childPath = TomlPath.AppendKey(path, key);
            int index = target.IndexOf(key);
            if (index < 0) {
                Add(target, key, incoming);
                _added.Add(childPath);
                continue;
            }
            TomlNode current = target._nodes[index];
            if (current is TomlTable table && incoming is TomlTable overlayTable) {
                MergeTables(table, overlayTable, childPath);
            } else if (current is TomlValue value && incoming is TomlValue overlayValue) {
                if (Assign(value, overlayValue)) _replaced.Add(childPath);
            } else if (IsArray(current) && IsArray(incoming)) {
                MergeArrays(target, key, current, incoming, childPath);
            } else {
                Conflict(childPath, current, incoming);
                target.PutNode(key, Copy(incoming, target.IsInlineScope()));
                _replaced.Add(childPath);
            }
        }
    }

    private void MergeArrays(TomlTable owner, string key, TomlNode current, TomlNode incoming, string path) {
        TomlArrayMerge mode = _options.Arrays;
        // An empty overlay array has nothing to match, so it replaces the base array as any other unkeyed array does.
        if (mode.Mode == TomlArrayMergeMode.MergeByKey && ItemCount(incoming) > 0 && AllKeyed(incoming, mode.Key!)) {
            int count = ItemCount(incoming);
            var matched = new HashSet<int>();
            for (int i = 0; i < count; i++) {
                var item = (TomlTable)ItemAt(incoming, i);
                int match = FindItem(current, mode.Key!, (TomlValue)item[mode.Key!], matched);
                if (match >= 0) {
                    matched.Add(match);
                    MergeTables((TomlTable)ItemAt(current, match), item, TomlPath.AppendIndex(path, match));
                } else {
                    matched.Add(ItemCount(current));
                    AppendItem(current, item, path);
                }
            }
            return;
        }
        if (mode.Mode == TomlArrayMergeMode.Append) {
            if (current is TomlTableArray && !AllTables(incoming)) {
                Conflict(path, current, incoming);
                owner.PutNode(key, Copy(incoming, owner.IsInlineScope()));
                _replaced.Add(path);
                return;
            }
            int count = ItemCount(incoming);
            for (int i = 0; i < count; i++) AppendItem(current, ItemAt(incoming, i), path);
            return;
        }
        if (AssignArray(owner, key, current, incoming)) _replaced.Add(path);
    }

    private void AppendItem(TomlNode array, TomlNode item, string path) {
        int index = ItemCount(array);
        if (array is TomlTableArray tables) AddTable(tables, (TomlTable)item);
        else ((TomlArray)array).InsertNode(index, Copy(item, inline: true));
        _added.Add(TomlPath.AppendIndex(path, index));
    }

    /// <summary>Adds a copy of an overlay entry under a new key, with the overlay's comment at the end of its line.</summary>
    private void Add(TomlTable owner, string key, TomlNode incoming) {
        bool inline = owner.IsInlineScope();
        TomlNode copy = Copy(incoming, inline);
        owner.PutNode(key, copy);
        if (!inline) KeepComment(incoming, copy);
    }

    private void AddTable(TomlTableArray tables, TomlTable item) {
        TomlTable copy = CopyTable(item, TomlTableKind.ArrayItem, inline: false);
        tables.AddItem(copy);
        KeepComment(item, copy);
    }

    private static bool AllKeyed(TomlNode array, string key) {
        int count = ItemCount(array);
        for (int i = 0; i < count; i++) {
            if (ItemAt(array, i) is not TomlTable table || !table.TryGetValue(key, out TomlNode? id) || id is not TomlValue) return false;
        }
        return true;
    }

    /// <summary>
    /// Returns the first base table with an equal value under the key that no earlier overlay table matched, so tables
    /// that repeat a value, as a lock file lists two versions of one package, pair up in order.
    /// </summary>
    private static int FindItem(TomlNode array, string key, TomlValue id, HashSet<int> matched) {
        int count = ItemCount(array);
        for (int i = 0; i < count; i++) {
            if (!matched.Contains(i) && ItemAt(array, i) is TomlTable table && table.TryGetValue(key, out TomlNode? value) && value is TomlValue candidate && SameValue(candidate, id)) return i;
        }
        return -1;
    }

    // ------------------------------------------------------------------ replacing by assignment

    /// <summary>Makes a base value equal to an overlay value, keeping its spelling when equal. Returns whether it changed.</summary>
    private static bool Assign(TomlValue target, TomlValue incoming) {
        if (SameValue(target, incoming)) return false;
        target.SetValue(incoming.Value);
        return true;
    }

    /// <summary>
    /// Makes the base node under a key equal to an overlay node with the fewest edits: equal values keep their text,
    /// tables keep the entries the overlay restates, and arrays change only the items that differ.
    /// </summary>
    private bool AssignChild(TomlTable owner, string key, TomlNode current, TomlNode incoming) {
        if (current is TomlValue value && incoming is TomlValue overlayValue) return Assign(value, overlayValue);
        if (current is TomlTable table && incoming is TomlTable overlayTable) {
            // A table made only by dotted keys or deeper headers has no text of its own once empty; an empty overlay
            // table is written as one.
            if (overlayTable._count > 0 || table._count == 0 || table._tableKind is not (TomlTableKind.Dotted or TomlTableKind.Implicit)) {
                return AssignTable(table, overlayTable);
            }
        } else if (IsArray(current) && IsArray(incoming)) {
            return AssignArray(owner, key, current, incoming);
        }
        owner.PutNode(key, Copy(incoming, owner.IsInlineScope()));
        return true;
    }

    private bool AssignTable(TomlTable target, TomlTable overlay) {
        bool changed = false;
        for (int i = 0; i < overlay._count; i++) {
            TomlNode incoming = overlay._nodes[i];
            string key = incoming._key!;
            int index = target.IndexOf(key);
            if (index >= 0) {
                if (AssignChild(target, key, target._nodes[index], incoming)) changed = true;
            } else {
                Add(target, key, incoming);
                changed = true;
            }
        }
        // Removed after the additions, so a dotted table never empties and vanishes before the overlay's keys arrive.
        for (int i = target._count - 1; i >= 0; i--) {
            if (overlay.IndexOf(target._nodes[i]._key!) >= 0) continue;
            target.Remove(target._nodes[i]._key!);
            changed = true;
        }
        return changed;
    }

    private bool AssignArray(TomlTable owner, string key, TomlNode current, TomlNode incoming) {
        // An array of tables cannot hold other items, and one without tables has no text, so neither can stay one.
        if (current is TomlTableArray && (!AllTables(incoming) || ItemCount(incoming) == 0)) {
            owner.PutNode(key, Copy(incoming, owner.IsInlineScope()));
            return true;
        }
        bool changed = false;
        int overlayCount = ItemCount(incoming);
        int common = Math.Min(ItemCount(current), overlayCount);
        for (int i = 0; i < common; i++) {
            if (AssignItem(current, i, ItemAt(current, i), ItemAt(incoming, i))) changed = true;
        }
        while (ItemCount(current) > overlayCount) {
            if (current is TomlArray array) array.RemoveAt(array._count - 1);
            else ((TomlTableArray)current).RemoveAt(((TomlTableArray)current).Count - 1);
            changed = true;
        }
        for (int i = ItemCount(current); i < overlayCount; i++) {
            TomlNode item = ItemAt(incoming, i);
            if (current is TomlTableArray tables) AddTable(tables, (TomlTable)item);
            else ((TomlArray)current).InsertNode(i, Copy(item, inline: true));
            changed = true;
        }
        return changed;
    }

    private bool AssignItem(TomlNode array, int index, TomlNode current, TomlNode incoming) {
        if (current is TomlValue value && incoming is TomlValue overlayValue) return Assign(value, overlayValue);
        if (current is TomlTable table && incoming is TomlTable overlayTable) return AssignTable(table, overlayTable);
        if (current is TomlArray nested && incoming is TomlArray overlayNested) return AssignNested(nested, overlayNested);
        // Only array values hold items of mixed kinds; items of an array of tables are all tables.
        ((TomlArray)array).PutAt(index, Copy(incoming, inline: true));
        return true;
    }

    private bool AssignNested(TomlArray target, TomlArray overlay) {
        bool changed = false;
        int common = Math.Min(target._count, overlay._count);
        for (int i = 0; i < common; i++) {
            if (AssignItem(target, i, target._items[i], overlay._items[i])) changed = true;
        }
        while (target._count > overlay._count) {
            target.RemoveAt(target._count - 1);
            changed = true;
        }
        for (int i = target._count; i < overlay._count; i++) {
            target.InsertNode(i, Copy(overlay._items[i], inline: true));
            changed = true;
        }
        return changed;
    }

    // ------------------------------------------------------------------ copying overlay nodes

    /// <summary>
    /// Copies an overlay node into new nodes of the base document, in the syntax its new place allows. Values keep the
    /// overlay's string quoting and integer base. Comments are copied by the caller, only for added entries.
    /// </summary>
    private TomlNode Copy(TomlNode node, bool inline) {
        switch (node) {
            case TomlValue value:
                return new TomlValue(_document, value._kind) { _value = value.Value, _style = value._style, _base = value._base };
            case TomlArray array: {
                var copy = new TomlArray(_document);
                for (int i = 0; i < array._count; i++) copy.AddLoaded(Copy(array._items[i], inline: true));
                return copy;
            }
            case TomlTableArray tables when inline: {
                var copy = new TomlArray(_document);
                foreach (TomlTable item in tables._items) copy.AddLoaded(CopyTable(item, TomlTableKind.Inline, inline: true));
                return copy;
            }
            case TomlTableArray tables: {
                var copy = new TomlTableArray(_document);
                foreach (TomlTable item in tables._items) {
                    TomlTable itemCopy = CopyTable(item, TomlTableKind.ArrayItem, inline: false);
                    copy.AddLoaded(itemCopy);
                    KeepComment(item, itemCopy);
                }
                return copy;
            }
            default: {
                var table = (TomlTable)node;
                TomlTableKind kind = inline || table._tableKind == TomlTableKind.Inline ? TomlTableKind.Inline
                    : table._tableKind == TomlTableKind.Dotted && OnlyKeyValues(table) ? TomlTableKind.Dotted
                    : TomlTableKind.Header;
                return CopyTable(table, kind, inline);
            }
        }
    }

    private TomlTable CopyTable(TomlTable table, TomlTableKind kind, bool inline) {
        var copy = new TomlTable(_document, kind);
        bool inner = inline || kind == TomlTableKind.Inline;
        for (int i = 0; i < table._count; i++) {
            TomlNode child = table._nodes[i];
            TomlNode childCopy = Copy(child, inner);
            copy.AddEntry(child._key!, childCopy);
            if (!inner) KeepComment(child, childCopy);
        }
        return copy;
    }

    private void KeepComment(TomlNode source, TomlNode copy) {
        if (copy is TomlTable { _tableKind: TomlTableKind.Dotted }) return;
        string? comment = source.TrailingComment;
        if (comment != null) _comments.Add(new KeyValuePair<TomlNode, string>(copy, comment));
    }

    /// <summary>Returns whether a dotted table holds only key/value pairs, so its copy can stay dotted keys.</summary>
    private static bool OnlyKeyValues(TomlTable table) {
        for (int i = 0; i < table._count; i++) {
            TomlNode node = table._nodes[i];
            if (node is TomlTableArray) return false;
            if (node is TomlTable child && child._tableKind != TomlTableKind.Inline && (child._tableKind != TomlTableKind.Dotted || !OnlyKeyValues(child))) return false;
        }
        return true;
    }

    // ------------------------------------------------------------------ helpers

    private static bool IsArray(TomlNode node) => node is TomlArray || node is TomlTableArray;

    private static int ItemCount(TomlNode array) => array is TomlArray values ? values._count : ((TomlTableArray)array)._items.Count;

    private static TomlNode ItemAt(TomlNode array, int index) => array is TomlArray values ? values._items[index] : ((TomlTableArray)array)._items[index];

    private static bool AllTables(TomlNode array) {
        int count = ItemCount(array);
        for (int i = 0; i < count; i++) {
            if (ItemAt(array, i) is not TomlTable) return false;
        }
        return true;
    }

    private static bool SameValue(TomlValue a, TomlValue b) => a._kind == b._kind && TomlValueConverter.ScalarEquals(a.Value, b.Value);

    private void Conflict(string path, TomlNode current, TomlNode incoming) {
        string? location = current.IsSource ? current.Start.ToString() : null;
        _diagnostics.Add(new OfficeDiagnostic(OfficeDiagnosticCodes.TomlMergeKindConflict, OfficeDiagnosticSeverity.Warning,
            "The overlay replaces " + TomlNode.Describe(current.Kind) + " at '" + path + "' with " + TomlNode.Describe(incoming.Kind) + ".",
            location, "TOML"));
    }
}
