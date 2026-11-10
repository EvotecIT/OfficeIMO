namespace OfficeIMO.Data.Conversions;

/// <summary>
/// A trie of the flattened paths in one table. Pass one grows it; pass two looks cells up in it without allocating, by
/// comparing hashed names. Columns are assigned in first-seen order once the scan completes.
/// </summary>
internal sealed class TableColumnNode {
    private readonly List<TableColumnNode> _named = new();
    private int[] _hashes = new int[4];
    private List<TableColumnNode?>? _indexed;

    public TableColumnNode(string? name, int index, TableColumnNode? parent) {
        Name = name;
        Index = index;
        Parent = parent;
    }

    public string? Name { get; }
    public int Index { get; }
    public TableColumnNode? Parent { get; }

    /// <summary>The column of the scalar at this path, or -1.</summary>
    public int Column = -1;
    public bool HasScalar;
    public bool HasArray;
    public bool HasNonScalarItem;
    public bool HasEmptyCollection;
    public bool JoinsArray;

    // Value evidence gathered by pass one, for column types.
    public DataScalarKind? Kind;
    public bool MixedKinds;
    public bool HasNull;
    public bool HasEmptyString;
    public bool IntegerOverflow;
    public bool FloatNotDecimal;

    public void Observe(DataScalarKind kind, ReadOnlySpan<char> text) {
        HasScalar = true;
        if (kind == DataScalarKind.Null) {
            HasNull = true;
            return;
        }

        if (kind == DataScalarKind.String && text.Length == 0) HasEmptyString = true;
        if (kind == DataScalarKind.Integer && !TableValues.FitsInt64(text)) IntegerOverflow = true;
        if (kind == DataScalarKind.Float && !TableValues.FitsDecimal(text)) FloatNotDecimal = true;
        if (Kind is not { } seen) {
            Kind = kind;
        } else if (seen != kind) {
            if ((seen == DataScalarKind.Integer && kind == DataScalarKind.Float) || (seen == DataScalarKind.Float && kind == DataScalarKind.Integer)) {
                Kind = DataScalarKind.Float;
            } else {
                MixedKinds = true;
            }
        }
    }

    public IReadOnlyList<TableColumnNode> Named => _named;
    public IReadOnlyList<TableColumnNode?> Indexed => (IReadOnlyList<TableColumnNode?>?)_indexed ?? Array.Empty<TableColumnNode?>();

    // Rows usually repeat their keys in the same order, so the child after the last match is tried first.
    private int _cursor;

    public TableColumnNode? FindName(ReadOnlySpan<char> name) {
        if (_named.Count > 0) {
            int next = _cursor < _named.Count ? _cursor : 0;
            if (name.SequenceEqual(_named[next].Name.AsSpan())) {
                _cursor = next + 1;
                return _named[next];
            }
        }

        int hash = Hash(name);
        for (int i = 0; i < _named.Count; i++) {
            if (_hashes[i] == hash && name.SequenceEqual(_named[i].Name.AsSpan())) {
                _cursor = i + 1;
                return _named[i];
            }
        }

        return null;
    }

    public TableColumnNode GetOrAddName(ReadOnlySpan<char> name) {
        TableColumnNode? existing = FindName(name);
        if (existing != null) return existing;
        var node = new TableColumnNode(name.ToString(), -1, this);
        if (_named.Count == _hashes.Length) Array.Resize(ref _hashes, _hashes.Length * 2);
        _hashes[_named.Count] = Hash(name);
        _named.Add(node);
        return node;
    }

    public TableColumnNode? FindIndex(int index) => _indexed != null && index < _indexed.Count ? _indexed[index] : null;

    public TableColumnNode GetOrAddIndex(int index) {
        _indexed ??= new List<TableColumnNode?>();
        while (_indexed.Count <= index) _indexed.Add(null);
        return _indexed[index] ??= new TableColumnNode(null, index, this);
    }

    /// <summary>Makes this trie the given one: used when the caller fixes the columns.</summary>
    public void Merge(TableColumnNode replacement) {
        _named.Clear();
        _indexed = null;
        _hashes = new int[4];
        HasScalar = replacement.HasScalar;
        foreach (TableColumnNode child in replacement._named) {
            TableColumnNode copy = GetOrAddName(child.Name.AsSpan());
            copy.CopyFrom(child);
        }

        if (replacement._indexed != null) {
            for (int i = 0; i < replacement._indexed.Count; i++) {
                if (replacement._indexed[i] is { } child) GetOrAddIndex(i).CopyFrom(child);
            }
        }
    }

    private void CopyFrom(TableColumnNode source) {
        Merge(source);
        Kind = source.Kind;
        MixedKinds = source.MixedKinds;
        HasNull = source.HasNull;
        HasEmptyString = source.HasEmptyString;
        IntegerOverflow = source.IntegerOverflow;
        FloatNotDecimal = source.FloatNotDecimal;
    }

    public string Path {
        get {
            if (Parent == null) return string.Empty;
            string parent = Parent.Path;
            if (Name == null) return parent + "[" + Index.ToString(CultureInfo.InvariantCulture) + "]";
            return parent.Length == 0 ? Name : parent + "." + Name;
        }
    }

    private static int Hash(ReadOnlySpan<char> key) {
        unchecked {
            int hash = (int)2166136261;
            foreach (char c in key) hash = (hash ^ c) * 16777619;
            return hash;
        }
    }
}

/// <summary>One table discovered by pass one.</summary>
internal sealed class TableLayout {
    public TableLayout(string name) {
        Name = name;
    }

    public string Name { get; }
    public TableColumnNode Root { get; } = new TableColumnNode(null, -1, null);
    public List<TableColumnNode> Columns { get; } = new();
    public List<string> Headers { get; } = new();
    public Type[] Types = Array.Empty<Type>();
    public bool Positional;
    public long Rows;

    /// <summary>Assigns columns depth-first in first-seen order and checks the result against the limits.</summary>
    public void Complete(DataConversionContext context) {
        var stack = new Stack<TableColumnNode>();
        stack.Push(Root);
        while (stack.Count > 0) {
            TableColumnNode node = stack.Pop();
            if (context.Options.ArrayColumns == DataArrayColumns.Join && node.Parent != null && node.HasArray && !node.HasNonScalarItem) {
                node.JoinsArray = true;
            }

            if (node.JoinsArray || node.HasScalar || (node.HasEmptyCollection && node.Named.Count == 0 && node.Indexed.Count == 0)) {
                if (node.HasEmptyCollection && !node.HasScalar && !node.JoinsArray) {
                    context.Report(OfficeDiagnosticCodes.EmptyCollectionDropped, OfficeConversionLossKind.Omission,
                        "Empty arrays and objects inside rows have no cells; they became empty cells.", node.Path);
                }

                node.Column = Columns.Count;
                Columns.Add(node);
                Headers.Add(node.Parent == null ? "value" : node.Path);
                if (node.Name != null && !DataPath.IsPlainName(node.Name)) {
                    context.Report(OfficeDiagnosticCodes.HeaderPathAmbiguous, OfficeConversionLossKind.Approximation,
                        "Keys containing '.', '[' or ']' (or empty keys) were written as literal headers, which read back as nested paths.", node.Path);
                }

                if (!node.JoinsArray && (node.Named.Count > 0 || node.Indexed.Count > 0)) {
                    context.Report(OfficeDiagnosticCodes.HeaderPathAmbiguous, OfficeConversionLossKind.Approximation,
                        "A key holds a value in some rows and an object or array in others; both became columns.", node.Path);
                }
            }

            if (node.JoinsArray) continue;
            for (int i = node.Indexed.Count - 1; i >= 0; i--) {
                if (node.Indexed[i] is { } item) stack.Push(item);
            }

            for (int i = node.Named.Count - 1; i >= 0; i--) stack.Push(node.Named[i]);
        }

        // A root holding only positional items is a table without headers.
        Positional = Root.Named.Count == 0 && Root.Indexed.Count > 0 && !Root.HasScalar;
        if (Columns.Count > context.Options.MaxColumns) {
            throw new InvalidDataException("The table '" + Name + "' would have " + Columns.Count.ToString(CultureInfo.InvariantCulture)
                + " columns, more than MaxColumns (" + context.Options.MaxColumns.ToString(CultureInfo.InvariantCulture)
                + "). Choose the rows with Select, or join arrays into cells with ArrayColumns = Join.");
        }

        Types = new Type[Columns.Count];
        for (int i = 0; i < Columns.Count; i++) Types[i] = TableValues.ColumnType(Columns[i]);
    }

}
