namespace OfficeIMO.Yaml;

/// <summary>
/// Deep merge of an overlay into a base document through the editing operations, so the base saves as a minimal diff.
/// Mappings merge key by key, scalars are assigned (an equal value keeps the base text), sequences follow
/// <see cref="YamlMergeOptions.Sequences"/>, and new entries are copies of the overlay's nodes with their comments,
/// anchors and aliases. The overlay is only read.
/// </summary>
/// <remarks>
/// <para>A mapping in the overlay is read through its typed view: its own entries, then the entries its <c>&lt;&lt;</c>
/// merge keys provide. In the base, a key that only a merge provides gets an entry of its own, as
/// <see cref="YamlMapping.Set"/> does; merged mappings are never edited or flattened. A copied overlay subtree keeps its
/// merge keys as written.</para>
/// <para>An overlay alias whose anchored node was copied becomes an alias to the copy. An alias whose anchored node merged
/// into an existing base node, or that would precede its anchor in the base, is written as a copy of the anchored node,
/// because the merged node no longer holds the overlay's value.</para>
/// </remarks>
internal sealed class YamlMerge {
    private const int CompareBudget = 1_000_000;
    private const int MaxMergeDepth = 1000;

    private readonly YamlMergeOptions _options;
    private readonly YamlDocument? _document;
    private readonly int _diagnosticsBefore;
    private readonly List<OfficeDiagnostic>? _local;
    private readonly List<string> _added = new List<string>();
    private readonly List<string> _replaced = new List<string>();
    private readonly List<string> _removed = new List<string>();
    /// <summary>Overlay anchored nodes and their copies in the current scope.</summary>
    private readonly Dictionary<YamlNode, YamlNode> _copies = new Dictionary<YamlNode, YamlNode>();
    /// <summary>Copies and the overlay nodes they came from.</summary>
    private readonly Dictionary<YamlNode, YamlNode> _originals = new Dictionary<YamlNode, YamlNode>();
    private readonly List<YamlAlias> _aliases = new List<YamlAlias>();
    private YamlDocumentPart? _part;
    private YamlNode? _scope;
    private HashSet<string>? _anchors;
    private bool _expandAll;
    private int _changes;

    private YamlMerge(YamlDocument? document, YamlMergeOptions? options) {
        _options = options ?? YamlMergeOptions.Default;
        _document = document;
        if (document is null) _local = new List<OfficeDiagnostic>();
        else _diagnosticsBefore = document.Diagnostics.Count;
    }

    public static YamlMergeReport Merge(YamlDocument target, YamlDocument overlay, YamlMergeOptions? options) {
        if (overlay is null) throw new ArgumentNullException(nameof(overlay));
        if (ReferenceEquals(target, overlay)) throw new ArgumentException("A document cannot be merged into itself.", nameof(overlay));
        var merge = new YamlMerge(target, options);
        int existing = target.Documents.Count;
        int count = Math.Max(existing, overlay.Documents.Count);
        for (int i = 0; i < overlay.Documents.Count; i++) {
            YamlNode overlayRoot = overlay.Documents[i].Root;
            // An empty overlay document (an empty values file) changes nothing.
            if (overlayRoot is null || overlayRoot.IsNull) continue;
            string prefix = count > 1 ? "$[" + i.ToString(CultureInfo.InvariantCulture) + "]" : string.Empty;
            if (i < target.Documents.Count) {
                YamlDocumentPart part = target.Documents[i];
                merge.Begin(part.Root, part);
                merge.MergeValue(new Slot(part), part.Root, overlayRoot, prefix, 0);
                merge.FixAliasOrder();
            } else {
                merge.Begin(null, null);
                YamlNode copy = merge.Copy(overlayRoot, keepAnchor: true, comments: true);
                // Documents are added in order, so an empty overlay document between them leaves no gap to fill.
                while (target.Documents.Count < i) target.AddDocument(new YamlScalar((string?)null));
                target.AddDocument(copy);
                merge.Added(prefix);
            }
        }
        return merge.Report();
    }

    public static YamlMergeReport Merge(YamlMapping target, YamlMapping overlay, YamlMergeOptions? options) {
        if (overlay is null) throw new ArgumentNullException(nameof(overlay));
        if (Contains(target, overlay) || Contains(overlay, target)) {
            throw new ArgumentException("The overlay and the target overlap; merge a mapping that is neither inside nor around the target.", nameof(overlay));
        }
        var merge = new YamlMerge(target.Document, options);
        YamlNode scope = target;
        while (scope._parent is not null) scope = scope._parent;
        merge.Begin(scope, target._part);
        merge.MergeMappings(target, overlay, PathOf(target), 0);
        merge.FixAliasOrder();
        return merge.Report();
    }

    private static bool Contains(YamlNode outer, YamlNode node) {
        for (YamlNode? current = node; current is not null; current = current._parent) {
            if (ReferenceEquals(current, outer)) return true;
        }
        return false;
    }

    private static string PathOf(YamlMapping mapping) {
        if (mapping._parent is null && mapping._part is null) return string.Empty;
        try {
            string path = YamlPath.Format(mapping);
            return path == "$" ? string.Empty : path;
        } catch (InvalidOperationException) {
            return string.Empty;
        }
    }

    private void Begin(YamlNode? scope, YamlDocumentPart? part) {
        _scope = scope;
        _part = part;
        _anchors = null;
        _copies.Clear();
        _originals.Clear();
        _aliases.Clear();
    }

    private YamlMergeReport Report() {
        IReadOnlyList<OfficeDiagnostic> diagnostics;
        if (_document is null) {
            diagnostics = _local!;
        } else {
            var recorded = new List<OfficeDiagnostic>();
            IReadOnlyList<OfficeDiagnostic> all = _document.Diagnostics;
            for (int i = _diagnosticsBefore; i < all.Count; i++) recorded.Add(all[i]);
            diagnostics = recorded;
        }
        return new YamlMergeReport(_added, _replaced, _removed, diagnostics);
    }

    // ------------------------------------------------------------------ merging

    /// <summary>Merges an overlay value into the base value at a slot.</summary>
    private void MergeValue(Slot slot, YamlNode current, YamlNode overlay, string path, int depth) {
        if (depth > MaxMergeDepth) throw new InvalidOperationException("The merge is nested deeper than " + MaxMergeDepth + " levels.");
        YamlNode value = current.Resolved, incoming = overlay.Resolved;
        if (IsPlaceholder(current, incoming)) {
            // An empty '{}' or '[]' in block context is a placeholder: it takes the overlay's block layout.
            YamlNode block = value is YamlMapping ? new YamlMapping() : new YamlSequence();
            if (Put(slot, block, current)) value = block;
        }
        if (value is YamlMapping mapping && incoming is YamlMapping overlayMapping) {
            MergeMappings(mapping, overlayMapping, path, depth + 1);
        } else if (value is YamlSequence sequence && incoming is YamlSequence overlaySequence) {
            MergeSequences(sequence, overlaySequence, path, depth + 1);
        } else if (value is YamlScalar scalar && incoming is YamlScalar overlayScalar) {
            if (AssignScalar(slot, scalar, overlay, overlayScalar)) Replaced(path);
        } else {
            // A null on either side is a value being set or cleared, not a conflict between structures.
            ReplaceNode(slot, current, overlay, path, conflict: !value.IsNull && !incoming.IsNull);
        }
    }

    private static bool IsPlaceholder(YamlNode current, YamlNode incoming) {
        if (current is YamlAlias || current._style != YamlStyle.Flow || current._properties is { Anchor: not null } or { Tag: not null }) return false;
        if (current._parent is { _style: YamlStyle.Flow } || incoming._style != YamlStyle.Block) return false;
        return (current is YamlMapping { _count: 0 } && incoming is YamlMapping { _count: > 0 })
            || (current is YamlSequence { _count: 0 } && incoming is YamlSequence { _count: > 0 });
    }

    private void MergeMappings(YamlMapping target, YamlMapping overlay, string path, int depth) {
        if (depth > MaxMergeDepth) throw new InvalidOperationException("The merge is nested deeper than " + MaxMergeDepth + " levels.");
        List<Entry> entries = Effective(overlay);
        for (int e = 0; e < entries.Count; e++) {
            Entry entry = entries[e];
            if (entry.Text is null) {
                // Collection keys have no text to match; the overlay's entry is added as written.
                YamlEdit.InsertEntry(target, target._count, Copy(entry.Key, keepAnchor: true, comments: true), Copy(entry.Value, keepAnchor: true, comments: true));
                Added(YamlPath.AppendKey(path, entry.Key.ToString()));
                continue;
            }
            string childPath = YamlPath.AppendKey(path, entry.Text);
            bool remove = _options.NullRemovesKey && entry.Value.IsNull;
            int index = target.IndexOf(entry.Text);
            if (index >= 0) {
                if (remove) Remove(target, index, childPath);
                else MergeValue(new Slot(target, index), target._entries[index].Value, entry.Value, childPath, depth + 1);
                continue;
            }
            if (target.TryGetValue(entry.Text, out YamlNode? merged)) {
                if (remove) {
                    Note(OfficeDiagnosticCodes.YamlMergeKeyNotRemoved, OfficeDiagnosticSeverity.Warning,
                        "The overlay null for '" + Display(childPath) + "' cannot remove the key, because a << merge key provides it and merged mappings are never edited.", target);
                } else {
                    OverrideMerged(target, entry, merged!, childPath, depth + 1);
                }
                continue;
            }
            if (remove) continue;
            YamlEdit.InsertEntry(target, target._count, CopyKey(entry.Key), Copy(entry.Value, keepAnchor: true, comments: true));
            Added(childPath);
        }
    }

    /// <summary>
    /// Gives a key that only a <c>&lt;&lt;</c> merge provides an entry of its own: the merged mapping's content with the
    /// overlay merged into it, or the overlay's value. Nothing is added when the overlay restates the merged value.
    /// </summary>
    private void OverrideMerged(YamlMapping target, Entry entry, YamlNode merged, string path, int depth) {
        YamlNode value = merged.Resolved, incoming = entry.Value.Resolved;
        if (value is YamlMapping mergedMapping && incoming is YamlMapping overlayMapping) {
            var own = (YamlMapping)Copy(mergedMapping, keepAnchor: false, comments: false);
            int before = _changes;
            MergeMappings(own, overlayMapping, path, depth);
            if (_changes != before) YamlEdit.InsertEntry(target, target._count, CopyKey(entry.Key), own);
            return;
        }
        int budget = CompareBudget;
        if (Equal(value, incoming, ref budget)) return;
        YamlEdit.InsertEntry(target, target._count, CopyKey(entry.Key), Copy(entry.Value, keepAnchor: true, comments: true));
        if (!SameKind(value, incoming) && !value.IsNull && !incoming.IsNull) NoteConflict(path, value, incoming, Location(target));
        Replaced(path);
    }

    private void MergeSequences(YamlSequence target, YamlSequence overlay, string path, int depth) {
        YamlSequenceMerge mode = _options.Sequences;
        switch (mode.Mode) {
            case YamlSequenceMergeMode.Append:
                for (int i = 0; i < overlay._count; i++) {
                    int index = target._count;
                    YamlEdit.InsertItem(target, index, Copy(overlay._items[i], keepAnchor: true, comments: true));
                    Added(YamlPath.AppendIndex(path, index));
                }
                break;
            case YamlSequenceMergeMode.MergeByKey:
                for (int i = 0; i < overlay._count; i++) {
                    YamlNode item = overlay._items[i];
                    string? id = ItemKey(item, mode.Key!);
                    int match = id is null ? -1 : FindItem(target, mode.Key!, id);
                    if (match >= 0) {
                        MergeMappings((YamlMapping)target._items[match].Resolved, (YamlMapping)item.Resolved, YamlPath.AppendIndex(path, match), depth + 1);
                    } else {
                        int index = target._count;
                        YamlEdit.InsertItem(target, index, Copy(item, keepAnchor: true, comments: true));
                        Added(YamlPath.AppendIndex(path, index));
                    }
                }
                break;
            default:
                if (AssignSequence(target, overlay, depth + 1)) Replaced(path);
                break;
        }
    }

    private static string? ItemKey(YamlNode item, string key) =>
        item.Resolved is YamlMapping mapping && mapping.TryGetValue(key, out YamlNode? value) && value!.Resolved is YamlScalar scalar && !scalar.IsNull
            ? scalar.Value
            : null;

    private static int FindItem(YamlSequence sequence, string key, string id) {
        for (int i = 0; i < sequence._count; i++) {
            if (ItemKey(sequence._items[i], key) == id) return i;
        }
        return -1;
    }

    // ------------------------------------------------------------------ replacing by assignment

    /// <summary>
    /// Makes a base value equal to an overlay value with the fewest edits: equal scalars keep their text, mappings keep
    /// the entries the overlay restates, and sequences change only the items that differ. Returns whether anything changed.
    /// </summary>
    private bool Assign(Slot slot, YamlNode current, YamlNode overlay, int depth) {
        if (depth > MaxMergeDepth) throw new InvalidOperationException("The merge is nested deeper than " + MaxMergeDepth + " levels.");
        YamlNode value = current.Resolved, incoming = overlay.Resolved;
        if (value is YamlScalar scalar && incoming is YamlScalar overlayScalar) return AssignScalar(slot, scalar, overlay, overlayScalar);
        if (value is YamlSequence sequence && incoming is YamlSequence overlaySequence) return AssignSequence(sequence, overlaySequence, depth + 1);
        if (value is YamlMapping mapping && incoming is YamlMapping overlayMapping && !HasMergeKey(mapping)) return AssignMapping(mapping, overlayMapping, depth + 1);
        int budget = CompareBudget;
        if (Equal(value, incoming, ref budget)) return false;
        return Put(slot, Copy(overlay, keepAnchor: true, comments: true), current);
    }

    private bool AssignSequence(YamlSequence target, YamlSequence overlay, int depth) {
        bool changed = false;
        int common = Math.Min(target._count, overlay._count);
        for (int i = 0; i < common; i++) {
            if (Assign(new Slot(target, i), target._items[i], overlay._items[i], depth)) changed = true;
        }
        while (target._count > overlay._count) {
            if (!TryRemove(target, target._count - 1)) break;
            changed = true;
        }
        for (int i = target._count; i < overlay._count; i++) {
            YamlEdit.InsertItem(target, target._count, Copy(overlay._items[i], keepAnchor: true, comments: true));
            changed = true;
        }
        return changed;
    }

    private bool AssignMapping(YamlMapping target, YamlMapping overlay, int depth) {
        List<Entry> entries = Effective(overlay);
        var keep = new HashSet<string>(StringComparer.Ordinal);
        foreach (Entry entry in entries) {
            if (entry.Text is not null) keep.Add(entry.Text);
        }
        bool changed = false;
        for (int i = target._count - 1; i >= 0; i--) {
            string? text = YamlMapping.KeyText(target._entries[i].Key);
            if (text is not null && keep.Contains(text) && target.IndexOf(text) == i) continue;
            if (TryRemove(target, i)) changed = true;
        }
        foreach (Entry entry in entries) {
            int index = entry.Text is null ? -1 : target.IndexOf(entry.Text);
            if (index >= 0) {
                if (Assign(new Slot(target, index), target._entries[index].Value, entry.Value, depth)) changed = true;
            } else {
                YamlEdit.InsertEntry(target, target._count, CopyKey(entry.Key), Copy(entry.Value, keepAnchor: true, comments: true));
                changed = true;
            }
        }
        return changed;
    }

    private static bool HasMergeKey(YamlMapping mapping) {
        for (int i = 0; i < mapping._count; i++) {
            if (mapping._entries[i].Key is YamlScalar key && key.IsMergeKey()) return true;
        }
        return false;
    }

    /// <summary>
    /// Assigns an overlay scalar's value to a base scalar through the editing rule, so an equal value keeps the base text.
    /// A base reached through an alias is the anchored scalar, which every alias shares. Returns whether it changed.
    /// </summary>
    private bool AssignScalar(Slot slot, YamlScalar target, YamlNode overlay, YamlScalar incoming) {
        Data(target, out YamlScalarKind kind, out string text, out string? tag);
        Data(incoming, out YamlScalarKind newKind, out string newText, out string? newTag);
        if (kind == newKind && text == newText && tag == newTag) return false;
        // A custom tag, or an integer beyond 64 bits, is copied as written rather than re-spelled.
        if (newTag is not null || (newKind == YamlScalarKind.Integer && !long.TryParse(newText, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out _))) {
            return Put(slot, Copy(overlay, keepAnchor: true, comments: false), slot.Current);
        }
        if (tag is not null) YamlEdit.SetTag(target, null);
        YamlScalarKind assigned = newKind is YamlScalarKind.Null or YamlScalarKind.Boolean or YamlScalarKind.Integer or YamlScalarKind.Float ? newKind : YamlScalarKind.String;
        YamlEdit.AssignScalar(target, newText, assigned, YamlStyle.None);
        return true;
    }

    // ------------------------------------------------------------------ structural edits

    private void ReplaceNode(Slot slot, YamlNode current, YamlNode overlay, string path, bool conflict) {
        YamlNode value = current.Resolved, incoming = overlay.Resolved;
        string? location = Location(current);
        if (!Put(slot, Copy(overlay, keepAnchor: true, comments: false), current)) return;
        if (conflict) NoteConflict(path, value, incoming, location);
        Replaced(path);
    }

    /// <summary>Puts a node at a slot, or records why it cannot replace the base value there.</summary>
    private bool Put(Slot slot, YamlNode node, YamlNode current) {
        try {
            if (slot.Part is not null) YamlEdit.ReplaceRoot(slot.Part, node);
            else YamlEdit.ReplaceValue(slot.Collection!, slot.Index, node);
            return true;
        } catch (InvalidOperationException exception) {
            Note(OfficeDiagnosticCodes.YamlMergeKindConflict, OfficeDiagnosticSeverity.Warning,
                "The overlay value cannot replace the base value" + current.Where() + ", which stays: " + exception.Message, current);
            return false;
        }
    }

    private void Remove(YamlMapping target, int index, string path) {
        YamlNode value = target._entries[index].Value;
        try {
            YamlEdit.RemoveChild(target, index);
            Removed(path);
        } catch (InvalidOperationException exception) {
            Note(OfficeDiagnosticCodes.YamlMergeKeyNotRemoved, OfficeDiagnosticSeverity.Warning,
                "The overlay null for '" + Display(path) + "' cannot remove the key: " + exception.Message, value);
        }
    }

    private bool TryRemove(YamlNode collection, int index) {
        YamlNode child = collection is YamlMapping mapping ? mapping._entries[index].Value : ((YamlSequence)collection)._items[index];
        try {
            YamlEdit.RemoveChild(collection, index);
            return true;
        } catch (InvalidOperationException exception) {
            Note(OfficeDiagnosticCodes.YamlMergeKeyNotRemoved, OfficeDiagnosticSeverity.Warning,
                "The base value" + child.Where() + " that the overlay replaces cannot be removed: " + exception.Message, child);
            return false;
        }
    }

    // ------------------------------------------------------------------ copying overlay nodes

    private YamlNode CopyKey(YamlNode key) => key is YamlAlias alias
        ? Copy(alias.Target, keepAnchor: false, comments: true)
        : Copy(key, keepAnchor: true, comments: true);

    /// <summary>
    /// Copies a node into new nodes that can join the base: values, styles, tags and, when asked, comments, anchors and
    /// aliases. Strings take their quoting from the base when they join it; block scalars and non-string plain scalars
    /// keep their spelling unless the base schema would read it as something else.
    /// </summary>
    private YamlNode Copy(YamlNode node, bool keepAnchor, bool comments) {
        YamlNode copy;
        switch (node) {
            case YamlAlias alias:
                return CopyAlias(alias);
            case YamlScalar scalar:
                copy = CopyScalar(scalar);
                break;
            case YamlMapping mapping: {
                var target = new YamlMapping(mapping._style);
                for (int i = 0; i < mapping._count; i++) {
                    YamlMappingEntry entry = mapping._entries[i];
                    // Of equal keys a lenient load admitted, only the one lookups see is copied.
                    string? text = YamlMapping.KeyText(entry.Key);
                    if (text is not null && mapping.IndexOf(text) != i) continue;
                    YamlNode key = entry.Key is YamlAlias keyAlias ? Copy(keyAlias.Target, keepAnchor: false, comments) : Copy(entry.Key, keepAnchor, comments);
                    YamlEdit.InsertEntry(target, target._count, key, Copy(entry.Value, keepAnchor, comments));
                }
                copy = target;
                break;
            }
            default: {
                var sequence = (YamlSequence)node;
                var target = new YamlSequence(sequence._style);
                for (int i = 0; i < sequence._count; i++) YamlEdit.InsertItem(target, target._count, Copy(sequence._items[i], keepAnchor, comments));
                copy = target;
                break;
            }
        }
        if (node.Tag is { } tag && copy is not YamlScalar) (copy._properties ??= new YamlNodeProperties()).Tag = tag;
        if (keepAnchor && node.Anchor is { } anchor) {
            (copy._properties ??= new YamlNodeProperties()).Anchor = UniqueAnchor(anchor, node);
            _copies[node] = copy;
            _originals[copy] = node;
        }
        if (comments) YamlEdit.CopyComments(node, copy);
        return copy;
    }

    private YamlNode CopyAlias(YamlAlias alias) {
        YamlNode target = alias.Target;
        if (!_expandAll) {
            YamlAlias? copy = null;
            if (_copies.TryGetValue(target, out YamlNode? copied)) copy = new YamlAlias(copied);
            else if (_part is not null && ReferenceEquals(target._part, _part) && target.Anchor is not null) copy = new YamlAlias(target);
            if (copy is not null) {
                _aliases.Add(copy);
                return copy;
            }
        }
        Note(OfficeDiagnosticCodes.YamlMergeAliasExpanded, OfficeDiagnosticSeverity.Information,
            "The overlay alias *" + alias.Name + " is written as a copy of its anchored node, because that node merged into an existing base node or would follow the alias.", null);
        return Copy(target, keepAnchor: false, comments: false);
    }

    private YamlScalar CopyScalar(YamlScalar scalar) {
        string value = scalar.Value;
        if (scalar.Tag is { } tag) {
            YamlStyle tagged = scalar._style == YamlStyle.None ? YamlStyle.Plain : scalar._style;
            var copy = new YamlScalar(value, tagged);
            (copy._properties = new YamlNodeProperties()).Tag = tag;
            return copy;
        }
        switch (scalar._style) {
            case YamlStyle.Literal:
            case YamlStyle.Folded:
                return new YamlScalar(value, scalar._style);
            case YamlStyle.Plain:
                YamlScalarKind kind = YamlSchemaResolver.Resolve(scalar.ValueSpan, scalar.Schema);
                if (kind == YamlScalarKind.String) return new YamlScalar(value);
                if (kind == YamlScalarKind.Null && value.Length == 0) return new YamlScalar((string?)null);
                if (kind is YamlScalarKind.Null or YamlScalarKind.Boolean or YamlScalarKind.Integer or YamlScalarKind.Float
                    && YamlSchemaResolver.Resolve(value.AsSpan(), BaseSchema) != kind) {
                    // The base schema reads the spelling differently (yes under YAML 1.1 is a string under 1.2): write the value.
                    Data(scalar, out _, out string canonical, out _);
                    return new YamlScalar(canonical, YamlStyle.Plain);
                }
                return new YamlScalar(value, YamlStyle.Plain);
            default:
                // Quoted strings follow the quoting convention of the base.
                return new YamlScalar(value);
        }
    }

    private YamlSchema BaseSchema => _part?.Schema ?? YamlSchema.Core;

    private string UniqueAnchor(string name, YamlNode node) {
        HashSet<string> anchors = _anchors ??= CollectAnchors();
        if (anchors.Add(name)) return name;
        for (int n = 2; ; n++) {
            string candidate = name + "-" + n.ToString(CultureInfo.InvariantCulture);
            if (!anchors.Add(candidate)) continue;
            Note(OfficeDiagnosticCodes.YamlMergeAnchorRenamed, OfficeDiagnosticSeverity.Information,
                "The overlay anchor &" + name + " is copied as &" + candidate + ", because the base already has an anchor named '" + name + "'.", null);
            return candidate;
        }
    }

    private HashSet<string> CollectAnchors() {
        var anchors = new HashSet<string>(StringComparer.Ordinal);
        if (_scope is not null) {
            foreach (YamlNode node in _scope.Descendants()) {
                if (node.Anchor is { } name) anchors.Add(name);
            }
        }
        return anchors;
    }

    /// <summary>
    /// Writes an alias the merge created as a copy of its anchored node when the alias comes before that node in the
    /// base, where a reader could not resolve it.
    /// </summary>
    private void FixAliasOrder() {
        if (_aliases.Count == 0 || _scope is null) return;
        YamlNode scope = _part?._root ?? _scope;
        var created = new HashSet<YamlAlias>(_aliases);
        var anchored = new HashSet<YamlNode>();
        var late = new List<YamlAlias>();
        foreach (YamlNode node in scope.Descendants()) {
            if (node is YamlAlias alias) {
                if (created.Contains(alias) && !anchored.Contains(alias.Target)) late.Add(alias);
            } else if (node.Anchor is not null) {
                anchored.Add(node);
            }
        }
        _expandAll = true;
        foreach (YamlAlias alias in late) {
            YamlNode source = _originals.TryGetValue(alias.Target, out YamlNode? original) ? original : alias.Target;
            YamlNode copy = Copy(source, keepAnchor: false, comments: false);
            if (alias._parent is YamlMapping mapping) YamlEdit.ReplaceValue(mapping, mapping.IndexOfNode(alias, out _), copy);
            else if (alias._parent is YamlSequence sequence) YamlEdit.ReplaceValue(sequence, sequence.IndexOf(alias), copy);
            else if (_part is not null) YamlEdit.ReplaceRoot(_part, copy);
            Note(OfficeDiagnosticCodes.YamlMergeAliasExpanded, OfficeDiagnosticSeverity.Information,
                "The overlay alias *" + alias.Name + " is written as a copy of its anchored node, because the alias would come before the anchor in the merged document.", null);
        }
        _expandAll = false;
    }

    // ------------------------------------------------------------------ values

    /// <summary>One entry of a mapping's typed view: own entries, then entries that merge keys provide.</summary>
    private readonly struct Entry {
        public Entry(YamlNode key, string? text, YamlNode value) {
            Key = key;
            Text = text;
            Value = value;
        }

        public YamlNode Key { get; }
        public string? Text { get; }
        public YamlNode Value { get; }
    }

    /// <summary>
    /// Lists a mapping's entries as its typed view sees them: its own entries (the one lookups pick of equal keys), then
    /// the entries of its <c>&lt;&lt;</c> merges in order, each key once.
    /// </summary>
    private static List<Entry> Effective(YamlMapping mapping) {
        var entries = new List<Entry>(mapping._count);
        Collect(mapping, entries, new HashSet<string>(StringComparer.Ordinal), 0);
        return entries;
    }

    private static void Collect(YamlMapping mapping, List<Entry> entries, HashSet<string> seen, int depth) {
        if (depth > MaxMergeDepth) throw new InvalidOperationException("Merge keys are nested deeper than " + MaxMergeDepth + " levels.");
        bool merges = false;
        for (int i = 0; i < mapping._count; i++) {
            YamlMappingEntry entry = mapping._entries[i];
            if (entry.Key is YamlScalar key && key.IsMergeKey()) {
                merges = true;
                continue;
            }
            string? text = YamlMapping.KeyText(entry.Key);
            if (text is null) {
                if (depth == 0) entries.Add(new Entry(entry.Key, null, entry.Value));
                continue;
            }
            if (mapping.IndexOf(text) != i || !seen.Add(text)) continue;
            entries.Add(new Entry(entry.Key, text, entry.Value));
        }
        if (!merges) return;
        for (int i = 0; i < mapping._count; i++) {
            YamlMappingEntry entry = mapping._entries[i];
            if (entry.Key is not YamlScalar key || !key.IsMergeKey()) continue;
            YamlNode source = entry.Value.Resolved;
            if (source is YamlMapping merged) {
                Collect(merged, entries, seen, depth + 1);
            } else if (source is YamlSequence list) {
                for (int s = 0; s < list._count; s++) {
                    if (list._items[s].Resolved is YamlMapping item) Collect(item, entries, seen, depth + 1);
                }
            }
        }
    }

    /// <summary>
    /// Reads a scalar's value under its own schema: its kind, a canonical text for nulls, booleans and numbers, the decoded
    /// text for others, and the tag when no schema resolves it.
    /// </summary>
    private static void Data(YamlScalar scalar, out YamlScalarKind kind, out string text, out string? customTag) {
        customTag = null;
        if (!scalar.TryGetKind(out kind)) {
            kind = YamlScalarKind.Custom;
            customTag = scalar.Tag;
            text = scalar.Value;
            return;
        }
        ReadOnlySpan<char> value = scalar.ValueSpan;
        YamlSchema schema = scalar.NumericSchema;
        switch (kind) {
            case YamlScalarKind.Null:
                text = "null";
                return;
            case YamlScalarKind.Boolean:
                text = YamlSchemaResolver.TryParseBoolean(value, schema, out bool flag) ? YamlScalarFormat.Format(flag) : scalar.Value;
                return;
            case YamlScalarKind.Integer:
                text = YamlSchemaResolver.TryParseInt64(value, schema, out long integer) ? integer.ToString(CultureInfo.InvariantCulture) : scalar.Value;
                return;
            case YamlScalarKind.Float:
                text = YamlSchemaResolver.TryParseDouble(value, schema, out double number) ? YamlScalarFormat.Format(number) : scalar.Value;
                return;
            default:
                // Timestamps and 1.1 merge or value keys are text to every other schema.
                kind = YamlScalarKind.String;
                text = scalar.Value;
                return;
        }
    }

    private static bool Equal(YamlNode left, YamlNode right, ref int budget) {
        left = left.Resolved;
        right = right.Resolved;
        if (--budget < 0) return false;
        switch (left) {
            case YamlScalar scalar when right is YamlScalar other:
                Data(scalar, out YamlScalarKind kind, out string text, out string? tag);
                Data(other, out YamlScalarKind otherKind, out string otherText, out string? otherTag);
                return kind == otherKind && text == otherText && tag == otherTag;
            case YamlSequence sequence when right is YamlSequence other:
                if (sequence._count != other._count) return false;
                for (int i = 0; i < sequence._count; i++) {
                    if (!Equal(sequence._items[i], other._items[i], ref budget)) return false;
                }
                return true;
            case YamlMapping mapping when right is YamlMapping other:
                List<Entry> entries = Effective(mapping), otherEntries = Effective(other);
                if (entries.Count != otherEntries.Count) return false;
                var lookup = new Dictionary<string, YamlNode>(StringComparer.Ordinal);
                foreach (Entry entry in otherEntries) {
                    if (entry.Text is null) return false;
                    lookup[entry.Text] = entry.Value;
                }
                foreach (Entry entry in entries) {
                    if (entry.Text is null || !lookup.TryGetValue(entry.Text, out YamlNode? value) || !Equal(entry.Value, value, ref budget)) return false;
                }
                return true;
            default:
                return false;
        }
    }

    private static bool SameKind(YamlNode left, YamlNode right) =>
        (left is YamlMapping && right is YamlMapping) || (left is YamlSequence && right is YamlSequence) || (left is YamlScalar && right is YamlScalar);

    // ------------------------------------------------------------------ reporting

    private void NoteConflict(string path, YamlNode value, YamlNode incoming, string? location) {
        Add(OfficeDiagnosticCodes.YamlMergeKindConflict, OfficeDiagnosticSeverity.Warning,
            "The overlay replaces " + Describe(value) + " with " + Describe(incoming) + " at '" + Display(path) + "'.", location);
    }

    private static string Describe(YamlNode node) => node switch {
        YamlMapping => "a mapping",
        YamlSequence => "a sequence",
        _ => "a scalar"
    };

    private static string? Location(YamlNode? node) => node is not null && node.HasSourceSpan ? node.Start.ToString() : null;

    private void Note(string code, OfficeDiagnosticSeverity severity, string message, YamlNode? at) => Add(code, severity, message, Location(at));

    private void Add(string code, OfficeDiagnosticSeverity severity, string message, string? location) {
        var diagnostic = new OfficeDiagnostic(code, severity, message, location, "YAML");
        if (_document is not null) _document.AddDiagnostic(diagnostic);
        else _local!.Add(diagnostic);
    }

    private static string Display(string path) => path.Length == 0 ? "$" : path;

    private void Added(string path) {
        _added.Add(Display(path));
        _changes++;
    }

    private void Replaced(string path) {
        _replaced.Add(Display(path));
        _changes++;
    }

    private void Removed(string path) {
        _removed.Add(Display(path));
        _changes++;
    }

    /// <summary>A place that holds a base value: a mapping value or sequence item, or a document root.</summary>
    private readonly struct Slot {
        public Slot(YamlNode collection, int index) {
            Collection = collection;
            Index = index;
            Part = null;
        }

        public Slot(YamlDocumentPart part) {
            Collection = null;
            Index = 0;
            Part = part;
        }

        public YamlNode? Collection { get; }
        public int Index { get; }
        public YamlDocumentPart? Part { get; }

        public YamlNode Current => Part is not null ? Part.Root
            : Collection is YamlMapping mapping ? mapping._entries[Index].Value
            : ((YamlSequence)Collection!)._items[Index];
    }
}
