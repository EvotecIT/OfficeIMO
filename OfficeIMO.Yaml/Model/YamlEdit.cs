namespace OfficeIMO.Yaml;

/// <summary>
/// The editing operations behind the public node methods. Edits never re-parse: they change the tree, mark the touched
/// node (<see cref="YamlNodeState.Dirty"/>, <see cref="YamlNodeState.PropertiesDirty"/>,
/// <see cref="YamlNodeState.CommentsDirty"/>, <see cref="YamlNodeState.Restructured"/>) and its ancestors
/// (<see cref="YamlNodeState.ContainsDirty"/>), and leave the source text to saving, which copies everything else.
/// </summary>
internal static class YamlEdit {
    public static YamlStyle CheckScalarStyle(YamlStyle style) => style switch {
        YamlStyle.None or YamlStyle.Plain or YamlStyle.SingleQuoted or YamlStyle.DoubleQuoted or YamlStyle.Literal or YamlStyle.Folded => style,
        _ => throw new ArgumentOutOfRangeException(nameof(style), style, "A scalar style is None, Plain, SingleQuoted, DoubleQuoted, Literal or Folded.")
    };

    public static YamlStyle CheckCollectionStyle(YamlStyle style) => style switch {
        YamlStyle.None or YamlStyle.Block => YamlStyle.Block,
        YamlStyle.Flow => YamlStyle.Flow,
        _ => throw new ArgumentOutOfRangeException(nameof(style), style, "A collection style is Block or Flow.")
    };

    // ------------------------------------------------------------------ scalars

    /// <summary>
    /// Gives a scalar a new value. An equal value under the document's schema keeps the source text. Otherwise the
    /// scalar keeps its style when it can represent the value at its position, and takes the most conservative one
    /// that can when it cannot.
    /// </summary>
    public static void AssignScalar(YamlScalar scalar, string text, YamlScalarKind kind, YamlStyle requested) {
        if (scalar._part is null) {
            // Not in a document yet: the style is decided when it joins one.
            scalar._value = text;
            scalar._style = kind == YamlScalarKind.String ? requested : YamlStyle.Plain;
            return;
        }
        bool restyle = kind == YamlScalarKind.String && requested != YamlStyle.None && requested != scalar._style;
        if (!restyle && Equal(scalar, text, kind)) return;
        bool key = IsKey(scalar, out YamlMapping? keyOf);
        if (key && kind == YamlScalarKind.String) {
            int existing = keyOf!.IndexOf(text);
            if (existing >= 0 && !ReferenceEquals(keyOf.Entries[existing].Key, scalar)) throw new ArgumentException("The mapping already has the key '" + text + "'.", nameof(text));
        }
        bool flow = InFlow(scalar);
        YamlStyle style;
        if (kind != YamlScalarKind.String) {
            style = YamlStyle.Plain;
        } else if (requested != YamlStyle.None) {
            style = YamlScalarFormat.Choose(text, requested, flow, key, QuoteFor(scalar, key));
        } else if (CanRepresent(scalar._style, text, flow, key)) {
            style = scalar._style;
        } else {
            style = YamlScalarFormat.Choose(text, YamlStyle.None, flow, key, QuoteFor(scalar, key));
        }
        // A standard tag that names another type would contradict the new value.
        if (scalar._properties?.Tag is { } tag && YamlSchemaResolver.TryGetTagKind(tag, out YamlScalarKind tagKind) && tagKind != kind) {
            scalar._properties.Tag = null;
        }
        // The comments it owns stay readable once its text is no longer the source's.
        if (scalar.HasSourceSpan) CaptureComments(scalar);
        // A block scalar keeps its content indentation; others get one when they become block scalars.
        if (scalar._style is not (YamlStyle.Literal or YamlStyle.Folded)) scalar._indent = 0;
        scalar._value = text;
        scalar._style = style;
        MarkChanged(scalar, YamlNodeState.New | YamlNodeState.Dirty);
        if (key) keyOf!.ResetIndex();
        NoteShared(scalar);
    }

    /// <summary>Whether a style can hold the text at the scalar's position without changing what readers see.</summary>
    private static bool CanRepresent(YamlStyle style, string text, bool flow, bool key) => style switch {
        YamlStyle.Plain => YamlScalarFormat.CanBePlain(text.AsSpan(), flow, key) && !YamlScalarFormat.IsSchemaAmbiguous(text.AsSpan()),
        YamlStyle.SingleQuoted => YamlScalarFormat.CanBeSingleQuoted(text.AsSpan()),
        YamlStyle.DoubleQuoted => true,
        YamlStyle.Literal or YamlStyle.Folded => !flow && !key && YamlScalarFormat.CanBeBlock(text.AsSpan()),
        _ => false
    };

    private static bool Equal(YamlScalar scalar, string text, YamlScalarKind kind) {
        if (!scalar.TryGetKind(out YamlScalarKind current)) return kind == YamlScalarKind.String && scalar.Value == text;
        ReadOnlySpan<char> value = scalar.ValueSpan;
        YamlSchema schema = scalar.NumericSchema;
        switch (kind) {
            case YamlScalarKind.Null:
                return scalar.ResolvesToNull();
            case YamlScalarKind.String:
                return current == YamlScalarKind.String && value.SequenceEqual(text.AsSpan());
            case YamlScalarKind.Integer:
                return current == YamlScalarKind.Integer && YamlSchemaResolver.TryParseInt64(value, schema, out long integer)
                    && integer == long.Parse(text, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
            case YamlScalarKind.Float:
                return current == YamlScalarKind.Float && YamlSchemaResolver.TryParseDouble(value, schema, out double number)
                    && YamlSchemaResolver.TryParseDouble(text.AsSpan(), YamlSchema.Core, out double expected) && number.Equals(expected);
            case YamlScalarKind.Boolean:
                return current == YamlScalarKind.Boolean && YamlSchemaResolver.TryParseBoolean(value, schema, out bool flag) && flag == (text == "true");
            default:
                return false;
        }
    }

    /// <summary>The kind a new scalar stands for: plain text resolves through the core schema, other styles are strings.</summary>
    private static YamlScalarKind KindOf(YamlScalar scalar) {
        if (scalar._style != YamlStyle.Plain) return YamlScalarKind.String;
        YamlScalarKind kind = YamlSchemaResolver.Resolve(scalar._value.AsSpan(), YamlSchema.Core);
        return kind is YamlScalarKind.Null or YamlScalarKind.Boolean or YamlScalarKind.Integer or YamlScalarKind.Float ? kind : YamlScalarKind.String;
    }

    // ------------------------------------------------------------------ collections

    /// <summary>Puts a node at a collection position, updating a scalar in place when a plain new scalar replaces one.</summary>
    public static YamlNode ReplaceValue(YamlNode collection, int index, YamlNode node) {
        YamlNode old = collection is YamlMapping mapping ? mapping.Entries[index].Value : ((YamlSequence)collection).Items[index];
        if (ReferenceEquals(old, node)) return old;
        if (old is YamlScalar oldScalar && node is YamlScalar newScalar && newScalar._part is null && newScalar._parent is null
            && newScalar._properties is null && (newScalar._state & YamlNodeState.New) != 0) {
            YamlScalarKind kind = KindOf(newScalar);
            AssignScalar(oldScalar, newScalar._value!, kind, kind == YamlScalarKind.String ? newScalar._style : YamlStyle.Plain);
            return oldScalar;
        }
        RequireDetached(node);
        CheckRemovable(old);
        bool hadRegion = old._start >= 0 && collection.HasSourceSpan;
        bool replacesBlock = hadRegion && old is YamlMapping or YamlSequence && old._style == YamlStyle.Block;
        int start = old._start, end = old._end;
        if (old.HasSourceSpan) Materialize(old);
        // The replacement takes over the old value's line, and with it the comment at its end.
        // A block collection only takes it where it moves to the key's line; elsewhere the replaced value's comment goes with it.
        bool takesComment = node is not (YamlMapping or YamlSequence) || node._style != YamlStyle.Block || collection is YamlMapping { _style: YamlStyle.Block };
        if (takesComment && old._properties is { CommentSet: true } oldEdits && node._properties is not { CommentSet: true }) {
            YamlNodeProperties properties = node._properties ??= new YamlNodeProperties();
            properties.CommentSet = true;
            properties.Comment = oldEdits.Comment;
        }
        old._parent = null;
        if (collection is YamlMapping map) map.Entries[index] = new YamlMappingEntry(map.Entries[index].Key, node);
        else ((YamlSequence)collection).Items[index] = node;
        node._index = index;
        Adopt(node, collection);
        if (hadRegion) {
            node._start = start;
            node._end = end;
            MarkChanged(node, replacesBlock ? YamlNodeState.Dirty | YamlNodeState.ReplacedBlock : YamlNodeState.Dirty);
        } else {
            MarkChanged(collection, YamlNodeState.None);
        }
        NoteShared(collection);
        return node;
    }

    public static YamlNode InsertEntry(YamlMapping mapping, int index, YamlNode key, YamlNode value) {
        RequireDetached(key);
        RequireDetached(value);
        if (ReferenceEquals(key, value)) throw new ArgumentException("The key and the value must be different nodes.", nameof(value));
        if (key.Resolved is YamlScalar keyScalar && keyScalar._value is { } keyText && mapping.IndexOf(keyText) >= 0) {
            throw new ArgumentException("The mapping already has the key '" + keyText + "'.", nameof(key));
        }
        Snapshot(mapping);
        if (mapping.Entries.Length == mapping._count) Array.Resize(ref mapping._entries, Math.Max(4, mapping._count * 2));
        Array.Copy(mapping.Entries, index, mapping.Entries, index + 1, mapping._count - index);
        mapping.Entries[index] = new YamlMappingEntry(key, value);
        mapping._count++;
        mapping.Renumber(index);
        mapping.ResetIndex();
        Adopt(key, mapping);
        Adopt(value, mapping);
        MoveValueComments(key, value);
        MarkChanged(mapping, YamlNodeState.Restructured);
        NoteShared(mapping);
        return value;
    }

    public static YamlNode InsertItem(YamlSequence sequence, int index, YamlNode item) {
        RequireDetached(item);
        Snapshot(sequence);
        if (sequence.Items.Length == sequence._count) Array.Resize(ref sequence._items, Math.Max(4, sequence._count * 2));
        Array.Copy(sequence.Items, index, sequence.Items, index + 1, sequence._count - index);
        sequence.Items[index] = item;
        sequence._count++;
        sequence.Renumber(index);
        Adopt(item, sequence);
        MarkChanged(sequence, YamlNodeState.Restructured);
        NoteShared(sequence);
        return item;
    }

    public static void RemoveChild(YamlNode collection, int index) {
        if (collection is YamlMapping mapping) {
            YamlMappingEntry entry = mapping.Entries[index];
            CheckRemovable(entry.Key);
            CheckRemovable(entry.Value);
            Snapshot(mapping);
            // Removed nodes keep their values and comments, so they can be inspected or added elsewhere.
            if (entry.Value.HasSourceSpan) Materialize(entry.Value);
            if (entry.Key.HasSourceSpan) Materialize(entry.Key);
            Array.Copy(mapping.Entries, index + 1, mapping.Entries, index, mapping._count - index - 1);
            mapping.Entries[--mapping._count] = default;
            mapping.Renumber(index);
            mapping.ResetIndex();
            entry.Key._parent = null;
            entry.Value._parent = null;
        } else {
            var sequence = (YamlSequence)collection;
            YamlNode item = sequence.Items[index];
            CheckRemovable(item);
            Snapshot(sequence);
            if (item.HasSourceSpan) Materialize(item);
            Array.Copy(sequence.Items, index + 1, sequence.Items, index, sequence._count - index - 1);
            sequence.Items[--sequence._count] = null!;
            sequence.Renumber(index);
            item._parent = null;
        }
        MarkChanged(collection, YamlNodeState.Restructured);
        if (ChildCount(collection) == 0 && collection._style == YamlStyle.Block) BecomeEmptyFlow(collection);
        NoteShared(collection);
    }

    private static int ChildCount(YamlNode collection) => collection is YamlMapping m ? m._count : ((YamlSequence)collection)._count;

    /// <summary>
    /// An empty block collection cannot be written as a block, so it becomes <c>[]</c> or <c>{}</c>. As a mapping value
    /// it moves onto its key's line, taking the comment on that line with it.
    /// </summary>
    private static void BecomeEmptyFlow(YamlNode collection) {
        if (collection.HasSourceSpan && collection._parent is YamlMapping { _style: YamlStyle.Block } mapping) {
            int index = mapping.IndexOfNode(collection, out bool isKey);
            if (!isKey) {
                string source = collection.Document!._source;
                int colon = YamlTrivia.ColonEnd(source, mapping.Entries[index].Key._end);
                int lineEnd = YamlTrivia.LineContentEnd(source, colon);
                for (int i = colon; i < lineEnd; i++) {
                    if (source[i] != '#' || (source[i - 1] != ' ' && source[i - 1] != '\t')) continue;
                    YamlNodeProperties properties = collection._properties ??= new YamlNodeProperties();
                    properties.CommentSet = true;
                    properties.Comment = YamlTrivia.CommentText(source, i, lineEnd);
                    break;
                }
            }
        }
        collection._style = YamlStyle.Flow;
    }

    public static void ReplaceRoot(YamlDocumentPart part, YamlNode node) {
        YamlNode? old = part.RootNode;
        if (ReferenceEquals(old, node)) return;
        RequireDetached(node);
        int start = -1, end = -1;
        if (old is not null) {
            CheckRemovable(old);
            start = old._start;
            end = old._end;
            if (old.HasSourceSpan) Materialize(old);
        }
        part.RootNode = node;
        Adopt(node, null, part);
        if (!part._new && start >= 0) {
            node._start = start;
            node._end = end;
            MarkChanged(node, YamlNodeState.Dirty);
        }
        part._document._partsDirty = true;
    }

    /// <summary>Records the loaded children's spans before the first insertion or removal.</summary>
    private static void Snapshot(YamlNode collection) {
        if (!collection.HasSourceSpan || (collection._state & YamlNodeState.Restructured) != 0) return;
        if (collection is YamlMapping mapping) {
            var spans = new int[mapping._count * 2];
            for (int i = 0; i < mapping._count; i++) {
                spans[2 * i] = mapping.Entries[i].Key._start;
                spans[2 * i + 1] = mapping.Entries[i].Value._end;
            }
            mapping._original = spans;
        } else {
            var sequence = (YamlSequence)collection;
            var spans = new int[sequence._count * 2];
            for (int i = 0; i < sequence._count; i++) {
                spans[2 * i] = sequence.Items[i]._start;
                spans[2 * i + 1] = sequence.Items[i]._end;
            }
            sequence._original = spans;
        }
    }

    // ------------------------------------------------------------------ attaching and detaching

    private static void RequireDetached(YamlNode node) {
        if (node._parent is not null || (node._part is not null && ReferenceEquals(node._part.RootNode, node))) {
            throw new InvalidOperationException("The node already belongs to a document; remove it from its collection first.");
        }
        if (node.HasSourceSpan) Materialize(node);
    }

    /// <summary>
    /// Makes a node a child of a collection (or the root of a part): sets its parent, joins it to the document, checks
    /// anchors and aliases against the document, and decides the styles its new scalars left open.
    /// </summary>
    private static void Adopt(YamlNode node, YamlNode? parent, YamlDocumentPart? rootOf = null) {
        node._parent = parent;
        YamlDocumentPart? part = rootOf ?? parent?._part;
        if (part is null) return;
        if (!ReferenceEquals(node._part, part) || HasUnresolvedStyles(node)) Join(node, part);
    }

    private static bool HasUnresolvedStyles(YamlNode node) => node._part is null;

    /// <summary>Joins a new subtree to a document part. Runs once per subtree, so its cost follows the inserted nodes.</summary>
    private static void Join(YamlNode root, YamlDocumentPart part) {
        YamlDocument document = part._document;
        HashSet<string>? anchors = null;
        foreach (YamlNode node in root.Descendants()) {
            if (node is YamlAlias alias && alias.Target._part is not null && !ReferenceEquals(alias.Target._part, part) && alias.Target._part.RootNode is not null) {
                throw new ArgumentException("The alias *" + alias.Name + " refers to an anchor in another document.");
            }
            if (node.Anchor is { } anchor) (anchors ??= new HashSet<string>(StringComparer.Ordinal)).Add(anchor);
        }
        if (anchors is not null) {
            // A repeated anchor name would capture aliases after it on the next load.
            foreach (YamlNode existing in part.RootNode?.Descendants() ?? default) {
                if (existing.Anchor is { } name && anchors.Contains(name) && !IsInside(existing, root)) {
                    throw new ArgumentException("The document already has an anchor named '" + name + "'.");
                }
            }
        }
        foreach (YamlNode node in root.Descendants()) {
            node._part = part;
            if (node is YamlAlias alias) alias.Target._state |= YamlNodeState.Aliased;
        }
        foreach (YamlNode node in root.Descendants()) {
            if (node is YamlScalar { _style: YamlStyle.None } scalar) ResolveStyle(scalar);
            else if (node is YamlScalar { _style: YamlStyle.Plain or YamlStyle.SingleQuoted or YamlStyle.DoubleQuoted or YamlStyle.Literal or YamlStyle.Folded } fixedStyle
                && (fixedStyle._state & YamlNodeState.New) != 0) {
                // A requested style gives way when it cannot represent the value here.
                bool key = IsKey(fixedStyle, out _);
                fixedStyle._style = YamlScalarFormat.Choose(fixedStyle._value!, fixedStyle._style, InFlow(fixedStyle), key, QuoteFor(fixedStyle, key));
            }
            if (node._parent is YamlMapping { _style: YamlStyle.Flow } || node._parent is YamlSequence { _style: YamlStyle.Flow }) {
                if (node is YamlMapping or YamlSequence) node._style = YamlStyle.Flow;
            }
        }
    }

    private static bool IsInside(YamlNode node, YamlNode root) {
        for (YamlNode? current = node; current is not null; current = current._parent) {
            if (ReferenceEquals(current, root)) return true;
        }
        return false;
    }

    /// <summary>Chooses a new string scalar's style from its position and the quoting its siblings use.</summary>
    private static void ResolveStyle(YamlScalar scalar) {
        bool key = IsKey(scalar, out _);
        bool flow = InFlow(scalar);
        string text = scalar._value!;
        YamlStyle quote = QuoteFor(scalar, key, out bool always);
        if (always && YamlScalarFormat.CanBeSingleQuoted(text.AsSpan())) {
            scalar._style = quote;
            return;
        }
        scalar._style = YamlScalarFormat.Choose(text, YamlStyle.None, flow, key, quote);
    }

    private static YamlStyle QuoteFor(YamlScalar scalar, bool key) => QuoteFor(scalar, key, out _);

    /// <summary>
    /// Reads the quoting convention from up to eight siblings on each side: which quote they use, and whether they quote
    /// strings that could be plain. Without siblings it falls back to the document's convention.
    /// </summary>
    private static YamlStyle QuoteFor(YamlScalar scalar, bool key, out bool always) {
        always = false;
        int single = 0, @double = 0, plain = 0, quotedOptional = 0;
        YamlNode? parent = scalar._parent;
        if (parent is YamlMapping mapping) {
            int index = mapping.IndexOfNode(scalar, out _);
            for (int i = Math.Max(0, index - 8); i < Math.Min(mapping._count, index + 9); i++) {
                if (i == index) continue;
                Count(key ? mapping.Entries[i].Key : mapping.Entries[i].Value);
            }
        } else if (parent is YamlSequence sequence) {
            int index = sequence.IndexOf(scalar);
            for (int i = Math.Max(0, index - 8); i < Math.Min(sequence._count, index + 9); i++) {
                if (i != index) Count(sequence.Items[i]);
            }
        }
        if (single + @double + plain == 0 && scalar._part is not null) {
            YamlLayoutProfile profile = scalar._part._document.Layout;
            return profile.Quote;
        }
        always = quotedOptional > 0 && quotedOptional > plain;
        return @double > single ? YamlStyle.DoubleQuoted : YamlStyle.SingleQuoted;

        void Count(YamlNode node) {
            if (node is not YamlScalar sibling || sibling._style == YamlStyle.None || sibling.Tag is not null) return;
            switch (sibling._style) {
                case YamlStyle.SingleQuoted:
                    single++;
                    if (Optional(sibling)) quotedOptional++;
                    break;
                case YamlStyle.DoubleQuoted:
                    @double++;
                    if (Optional(sibling)) quotedOptional++;
                    break;
                case YamlStyle.Plain:
                    if (sibling.TryGetKind(out YamlScalarKind kind) && kind == YamlScalarKind.String && !sibling.IsNull) plain++;
                    break;
            }
        }

        static bool Optional(YamlScalar sibling) {
            ReadOnlySpan<char> value = sibling.ValueSpan;
            return YamlScalarFormat.CanBePlain(value, flow: false) && !YamlScalarFormat.IsSchemaAmbiguous(value);
        }
    }

    /// <summary>
    /// Converts a loaded subtree into nodes that hold their own values, keeping styles, anchors, tags and the comments the
    /// nodes own, so it can move elsewhere. The source text it came from is left to the save path.
    /// </summary>
    public static void Materialize(YamlNode root) {
        foreach (YamlNode node in root.Descendants()) {
            if (!node.HasSourceSpan) continue;
            CaptureComments(node);
            if (node is YamlScalar scalar) {
                scalar._value = scalar.Value;
                scalar._indent = 0;
            }
        }
        foreach (YamlNode node in root.Descendants()) {
            if (!node.HasSourceSpan) continue;
            node._state = (node._state & YamlNodeState.Aliased) | YamlNodeState.New;
            node._start = node._end = -1;
            if (node is YamlMapping mapping) mapping._original = null;
            else if (node is YamlSequence sequence) sequence._original = null;
        }
    }

    private static void CaptureComments(YamlNode node) => CopyComments(node, node);

    /// <summary>
    /// Gives <paramref name="target"/> the comments <paramref name="source"/> owns: the comment lines and blank line above
    /// its entry and the comment at the end of its line. A merge uses it to carry an overlay node's comments to its copy.
    /// </summary>
    public static void CopyComments(YamlNode source, YamlNode target) {
        if (ReferenceEquals(YamlTrivia.Owner(source), source) && !YamlTrivia.InFlow(source)) {
            string? before = YamlTrivia.GetCommentBefore(source);
            bool space = YamlTrivia.GetSpaceBefore(source);
            if (before is not null || space) {
                YamlNodeProperties properties = target._properties ??= new YamlNodeProperties();
                properties.CommentBeforeSet = true;
                properties.CommentBefore = before;
                properties.SpaceBeforeSet = true;
                properties.SpaceBefore = space;
            }
        }
        if (ReferenceEquals(YamlTrivia.TrailingTarget(source), source) && YamlTrivia.GetComment(source) is { } comment) {
            YamlNodeProperties properties = target._properties ??= new YamlNodeProperties();
            properties.CommentSet = true;
            properties.Comment = comment;
        }
    }

    /// <summary>Moves comments set on a new value to its key, which owns the entry's lines.</summary>
    private static void MoveValueComments(YamlNode key, YamlNode value) {
        if (value._properties is not { } properties) return;
        if (properties.CommentBeforeSet || properties.SpaceBeforeSet) {
            YamlNodeProperties keyProperties = key._properties ??= new YamlNodeProperties();
            if (properties.CommentBeforeSet) {
                keyProperties.CommentBeforeSet = true;
                keyProperties.CommentBefore = properties.CommentBefore;
            }
            if (properties.SpaceBeforeSet) {
                keyProperties.SpaceBeforeSet = true;
                keyProperties.SpaceBefore = properties.SpaceBefore;
            }
            properties.CommentBeforeSet = properties.SpaceBeforeSet = false;
            properties.CommentBefore = null;
        }
        if (properties.CommentSet && value is YamlMapping or YamlSequence && value._style == YamlStyle.Block) {
            YamlNodeProperties keyProperties = key._properties ??= new YamlNodeProperties();
            keyProperties.CommentSet = true;
            keyProperties.Comment = properties.Comment;
            properties.CommentSet = false;
            properties.Comment = null;
        }
    }

    // ------------------------------------------------------------------ properties and comments

    public static void SetAnchor(YamlNode node, string? name) {
        if (node is YamlAlias) throw new InvalidOperationException("An alias cannot have an anchor.");
        if (name is not null && !YamlScalarFormat.IsValidAnchor(name)) {
            throw new ArgumentException("'" + name + "' is not a valid anchor name: use printable characters without spaces or , [ ] { }.", nameof(name));
        }
        string? old = node.Anchor;
        if (old == name) return;
        List<YamlAlias>? aliases = null;
        if (node._part is not null) {
            foreach (YamlNode other in AllNodes(node._part)) {
                if (name is not null && !ReferenceEquals(other, node) && other.Anchor == name) {
                    throw new ArgumentException("The document already has an anchor named '" + name + "'.", nameof(name));
                }
                if (other is YamlAlias alias && ReferenceEquals(alias.Target, node)) (aliases ??= new List<YamlAlias>()).Add(alias);
            }
        }
        if (name is null && aliases is not null) {
            throw new InvalidOperationException("Aliases refer to the anchor '" + old + "'; remove them before removing the anchor.");
        }
        (node._properties ??= new YamlNodeProperties()).Anchor = name;
        MarkProperties(node);
        if (aliases is not null) {
            foreach (YamlAlias alias in aliases) {
                if (alias.HasSourceSpan) MarkChanged(alias, YamlNodeState.New | YamlNodeState.Dirty);
            }
        }
    }

    public static void SetTag(YamlNode node, string? tag) {
        if (node is YamlAlias) throw new InvalidOperationException("An alias cannot have a tag.");
        if (node.Tag == tag) return;
        (node._properties ??= new YamlNodeProperties()).Tag = tag;
        MarkProperties(node);
        NoteShared(node);
    }

    private static void MarkProperties(YamlNode node) {
        if (node.HasSourceSpan) MarkChanged(node, YamlNodeState.PropertiesDirty);
        else MarkChanged(node, YamlNodeState.None);
    }

    public static void SetCommentBefore(YamlNode node, string? text) {
        YamlNode owner = YamlTrivia.Owner(node);
        if (YamlTrivia.InFlow(owner)) throw new InvalidOperationException("Comment lines cannot be placed inside a flow collection.");
        RequireOwnLine(owner);
        YamlNodeProperties properties = owner._properties ??= new YamlNodeProperties();
        properties.CommentBeforeSet = true;
        properties.CommentBefore = text;
        MarkComments(owner);
    }

    public static void SetSpaceBefore(YamlNode node, bool value) {
        YamlNode owner = YamlTrivia.Owner(node);
        if (YamlTrivia.InFlow(owner)) throw new InvalidOperationException("Blank lines cannot be placed inside a flow collection.");
        RequireOwnLine(owner);
        if (YamlTrivia.GetSpaceBefore(owner) == value) return;
        YamlNodeProperties properties = owner._properties ??= new YamlNodeProperties();
        properties.SpaceBeforeSet = true;
        properties.SpaceBefore = value;
        MarkComments(owner);
    }

    public static void SetComment(YamlNode node, string? text) {
        YamlNode target = YamlTrivia.TrailingTarget(node);
        if (YamlTrivia.InFlow(target)) {
            throw new InvalidOperationException("Comments cannot be placed inside a flow collection.");
        }
        if (target is YamlMapping or YamlSequence && target._style == YamlStyle.Block && (target._parent is not null || target._part is not null)) {
            throw new InvalidOperationException("A block collection that is a sequence item or a document root has no line of its own; set the comment on its first entry.");
        }
        if (target._parent is YamlMapping mapping) {
            int index = mapping.IndexOfNode(target, out bool isKey);
            YamlNode value = mapping.Entries[index].Value;
            if (isKey && !(value is YamlMapping or YamlSequence && value._style == YamlStyle.Block)) {
                throw new InvalidOperationException("The comment after a key whose value is on the same line belongs to the value; set it on the value.");
            }
        }
        YamlNodeProperties properties = target._properties ??= new YamlNodeProperties();
        properties.CommentSet = true;
        properties.Comment = text;
        MarkComments(target);
    }

    /// <summary>Fails for a loaded entry that shares its line with earlier content, where no line above it is its own.</summary>
    private static void RequireOwnLine(YamlNode owner) {
        if (owner._start < 0 || owner._part is null || owner._part._new) return;
        if (!YamlTrivia.TryBeforeRegion(owner, out _, out _, out _)) {
            throw new InvalidOperationException("The node shares its line with earlier content, so no comment line above it would be its own.");
        }
    }

    private static void MarkComments(YamlNode owner) {
        if (owner._part is null) return;
        MarkChanged(owner, owner._start >= 0 ? YamlNodeState.CommentsDirty : YamlNodeState.None);
    }

    // ------------------------------------------------------------------ bookkeeping

    /// <summary>Sets flags on a node and marks its ancestors as containing an edit.</summary>
    public static void MarkChanged(YamlNode node, YamlNodeState flags) {
        node._state |= flags;
        for (YamlNode? parent = node._parent; parent is not null && (parent._state & YamlNodeState.ContainsDirty) == 0; parent = parent._parent) {
            parent._state |= YamlNodeState.ContainsDirty;
        }
        if (node._part is not null) node._part._document._partsDirty = true;
    }

    public static bool IsKey(YamlNode node, out YamlMapping? mapping) {
        mapping = node._parent as YamlMapping;
        if (mapping is null) return false;
        mapping.IndexOfNode(node, out bool isKey);
        return isKey;
    }

    public static bool InFlow(YamlNode node) => node._parent is not null && node._parent._style == YamlStyle.Flow;

    /// <summary>
    /// Fails when aliases outside a subtree refer to anchors inside it, because removing or replacing it would leave them
    /// pointing at nothing.
    /// </summary>
    private static void CheckRemovable(YamlNode root) {
        HashSet<YamlNode>? anchored = null;
        foreach (YamlNode node in root.Descendants()) {
            if ((node._state & YamlNodeState.Aliased) != 0) (anchored ??= new HashSet<YamlNode>()).Add(node);
        }
        if (anchored is null || root._part is null) return;
        foreach (YamlNode node in AllNodes(root._part)) {
            if (node is YamlAlias alias && anchored.Contains(alias.Target) && !IsInside(alias, root)) {
                throw new InvalidOperationException("The alias *" + alias.Name + alias.Where() + " refers to an anchor inside the node; remove the alias first.");
            }
        }
    }

    private static YamlNodeWalk AllNodes(YamlDocumentPart part) => part.RootNode is null ? default : part.RootNode.Descendants();

    /// <summary>
    /// Records a diagnostic, once per anchored node, when an edit changes a node that aliases share: the change shows
    /// through every alias, which is easy to miss when the edit went through one of them.
    /// </summary>
    public static void NoteShared(YamlNode node) {
        for (YamlNode? current = node; current is not null; current = current._parent) {
            if ((current._state & YamlNodeState.Aliased) == 0) continue;
            YamlDocument? document = current.Document;
            if (document is null || !document.NoteSharedEdit(current)) return;
            string message = "The edit changes the node anchored as &" + current.Anchor + current.Where()
                + ", which aliases share; every alias sees the change.";
            string? location = current.HasSourceSpan ? current.Start.ToString() : null;
            document.AddDiagnostic(new OfficeDiagnostic(OfficeDiagnosticCodes.YamlAliasedNodeEdited, OfficeDiagnosticSeverity.Information, message, location, "YAML"));
            return;
        }
    }
}
