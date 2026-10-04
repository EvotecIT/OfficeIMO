namespace OfficeIMO.Bibliography;

internal sealed partial class BibliographyReferenceResolver {
    private readonly BibliographyReferenceOptions _options;
    private readonly CancellationToken _cancellationToken;
    private readonly BibliographyModelCopy _copy;
    private readonly List<BibliographyDiagnostic> _diagnostics = new List<BibliographyDiagnostic>();
    private readonly List<BibliographyItem> _items = new List<BibliographyItem>();
    private readonly List<Reference>[] _references;
    private readonly Dictionary<string, BibliographyFieldProvenance>[] _provenance;
    private int _referenceCount;

    private BibliographyReferenceResolver(BibliographyDocument source, BibliographyReferenceOptions options, CancellationToken cancellationToken) {
        _options = options; _cancellationToken = cancellationToken;
        _copy = new BibliographyModelCopy(options, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        if (source.Items.Count > options.MaximumItems) throw new InvalidOperationException("Bibliography reference resolution exceeds its item limit.");
        _references = new List<Reference>[source.Items.Count];
        _provenance = new Dictionary<string, BibliographyFieldProvenance>[source.Items.Count];
        for (int index = 0; index < source.Items.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            _items.Add(_copy.Item(source.Items[index]));
            _references[index] = new List<Reference>();
            _provenance[index] = new Dictionary<string, BibliographyFieldProvenance>(StringComparer.Ordinal);
        }
    }

    internal static BibliographyReferenceResult Resolve(BibliographyDocument source, BibliographyReferenceOptions options, CancellationToken cancellationToken) {
        var resolver = new BibliographyReferenceResolver(source, options, cancellationToken);
        resolver.IndexReferences();
        resolver.ResolveGraph();
        var entries = new List<BibliographyNativeEntry>();
        foreach (BibliographyNativeEntry entry in source.NativeEntries) {
            cancellationToken.ThrowIfCancellationRequested();
            entries.Add(new BibliographyNativeEntry(entry.Format, resolver._copy.Value(entry.Kind)!,
                resolver._copy.Value(entry.Value)!, resolver._copy.Value(entry.Name)));
        }
        var document = new BibliographyDocument(source, resolver._items, entries, cancellationToken);
        var provenance = new List<BibliographyFieldProvenance>();
        foreach (Dictionary<string, BibliographyFieldProvenance> fields in resolver._provenance) {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (KeyValuePair<string, BibliographyFieldProvenance> field in fields.OrderBy(pair => pair.Key, StringComparer.Ordinal)) {
                cancellationToken.ThrowIfCancellationRequested();
                provenance.Add(field.Value);
            }
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new BibliographyReferenceResult(document, resolver._diagnostics, provenance);
    }

    internal static bool IsDataContainer(BibliographyItem item) => string.Equals(item.NativeType, "xdata", StringComparison.OrdinalIgnoreCase);
    private static bool IsBib(BibliographyFormat format) => format == BibliographyFormat.BibTex || format == BibliographyFormat.BibLatex;

    private void IndexReferences() {
        var keys = new Dictionary<string, int>(StringComparer.Ordinal);
        var ambiguous = new HashSet<string>(StringComparer.Ordinal);
        for (int index = 0; index < _items.Count; index++) {
            _cancellationToken.ThrowIfCancellationRequested();
            string key = _items[index].Key;
            if (string.IsNullOrWhiteSpace(key)) {
                Diagnose("BIBREF001", BibliographyDiagnosticSeverity.Error, "Record has no usable citation key.", index);
            } else if (keys.ContainsKey(key)) {
                ambiguous.Add(key);
                Diagnose("BIBREF002", BibliographyDiagnosticSeverity.Error, "Duplicate citation key '" + key + "'; references to it are ambiguous.", index);
            } else keys.Add(key, index);
        }
        for (int index = 0; index < _items.Count; index++) {
            _cancellationToken.ThrowIfCancellationRequested();
            // Xdata replacement precedes crossref fallback, independent of source field order.
            if (_options.ResolveXData) IndexField(index, "xdata", true, keys, ambiguous);
            if (_options.ResolveCrossref) IndexField(index, "crossref", false, keys, ambiguous);
        }
    }

    private void IndexField(int itemIndex, string relation, bool multiple, IDictionary<string, int> keys, ISet<string> ambiguous) {
        BibliographyNativeField? reference = null;
        bool duplicate = false;
        foreach (BibliographyNativeField field in _items[itemIndex].NativeFields) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (!IsBib(field.Format) || !string.Equals(field.Name, relation, StringComparison.OrdinalIgnoreCase)) continue;
            if (reference != null) duplicate = true;
            reference = field;
        }
        if (reference == null) return;
        if (duplicate) {
            Diagnose("BIBREF006", BibliographyDiagnosticSeverity.Error, "Repeated '" + relation + "' fields have ambiguous precedence and were not resolved.", itemIndex, relation);
            return;
        }
        string value = reference.Value;
        int start = 0;
        for (int position = 0; position <= value.Length; position++) {
            if ((position & 4095) == 0) _cancellationToken.ThrowIfCancellationRequested();
            if (position < value.Length && (!multiple || value[position] != ',')) continue;
            if (_referenceCount >= _options.MaximumReferences) throw new InvalidOperationException("Bibliography reference resolution exceeds its reference limit.");
            _referenceCount++;
            string key = value.Substring(start, position - start).Trim();
            start = position + 1;
            if (key.Length == 0) {
                Diagnose("BIBREF006", BibliographyDiagnosticSeverity.Warning, "Empty '" + relation + "' reference was not resolved.", itemIndex, relation);
            } else if (ambiguous.Contains(key)) {
                Diagnose("BIBREF002", BibliographyDiagnosticSeverity.Error, "Reference to duplicate citation key '" + key + "' was not resolved.", itemIndex, relation);
            } else if (!keys.TryGetValue(key, out int parentIndex)) {
                Diagnose("BIBREF003", BibliographyDiagnosticSeverity.Warning, "Referenced key '" + key + "' was not found.", itemIndex, relation);
            } else if (multiple && !IsDataContainer(_items[parentIndex])) {
                Diagnose("BIBREF007", BibliographyDiagnosticSeverity.Warning, "Xdata reference '" + key + "' does not identify an @xdata container.", itemIndex, relation);
            } else {
                _references[itemIndex].Add(new Reference(itemIndex, parentIndex, relation));
            }
        }
    }

    private void ResolveGraph() {
        // Topological processing avoids recursion and prevents cache order from bypassing chain limits.
        var pending = new int[_items.Count];
        var depth = new int[_items.Count];
        var complete = new bool[_items.Count];
        var failed = new bool[_items.Count];
        var children = new List<int>[_items.Count];
        var ready = new Queue<int>();
        for (int index = 0; index < _items.Count; index++) {
            _cancellationToken.ThrowIfCancellationRequested();
            children[index] = new List<int>();
        }
        for (int index = 0; index < _items.Count; index++) {
            _cancellationToken.ThrowIfCancellationRequested();
            pending[index] = _references[index].Count;
            foreach (Reference reference in _references[index]) {
                _cancellationToken.ThrowIfCancellationRequested();
                children[reference.Parent].Add(index);
            }
            if (pending[index] == 0) ready.Enqueue(index);
        }
        while (ready.Count != 0) {
            _cancellationToken.ThrowIfCancellationRequested();
            int index = ready.Dequeue();
            foreach (Reference reference in _references[index]) {
                _cancellationToken.ThrowIfCancellationRequested();
                depth[index] = Math.Max(depth[index], depth[reference.Parent] + 1);
                if (failed[reference.Parent]) {
                    failed[index] = true;
                    Diagnose("BIBREF008", BibliographyDiagnosticSeverity.Error, "Parent '" + _items[reference.Parent].Key + "' could not be resolved within the configured limits.", index, reference.Relation);
                }
            }
            if (!failed[index] && depth[index] > _options.MaximumDepth) {
                failed[index] = true;
                Diagnose("BIBREF005", BibliographyDiagnosticSeverity.Error, "Reference chain exceeds MaximumDepth; this record retains its original values.", index);
            }
            if (!failed[index]) foreach (Reference reference in _references[index]) Merge(reference);
            complete[index] = true;
            foreach (int child in children[index]) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (--pending[child] == 0) ready.Enqueue(child);
            }
        }
        for (int index = 0; index < _items.Count; index++) {
            _cancellationToken.ThrowIfCancellationRequested();
            if (!complete[index]) Diagnose("BIBREF004", BibliographyDiagnosticSeverity.Error,
                "Record participates in or depends on a reference cycle; it retains its original values.", index);
        }
    }

    private void Diagnose(string code, BibliographyDiagnosticSeverity severity, string message, int itemIndex, string? field = null) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (_diagnostics.Count >= _options.MaximumDiagnostics) throw new InvalidOperationException("Bibliography reference resolution exceeds its diagnostic limit.");
        _diagnostics.Add(new BibliographyDiagnostic(code, severity, message, itemKey: _items[itemIndex].Key, field: field));
    }

    private void Record(Reference reference, string targetField, string sourceField) {
        _cancellationToken.ThrowIfCancellationRequested();
        BibliographyItem child = _items[reference.Child], parent = _items[reference.Parent];
        _provenance[reference.Parent].TryGetValue(sourceField, out BibliographyFieldProvenance? inherited);
        var path = new List<string> { _copy.Value(child.Key)! };
        if (inherited == null) path.Add(_copy.Value(parent.Key)!);
        else foreach (string key in inherited.ReferencePath) path.Add(_copy.Value(key)!);
        _provenance[reference.Child][targetField] = new BibliographyFieldProvenance(reference.Child, child.Key, targetField,
            inherited?.SourceItemIndex ?? reference.Parent, inherited?.SourceItemKey ?? parent.Key,
            inherited?.SourceField ?? sourceField, reference.Relation, Array.AsReadOnly(path.ToArray()));
    }

    private sealed class Reference {
        internal Reference(int child, int parent, string relation) { Child = child; Parent = parent; Relation = relation; }
        internal int Child { get; }
        internal int Parent { get; }
        internal string Relation { get; }
    }
}
