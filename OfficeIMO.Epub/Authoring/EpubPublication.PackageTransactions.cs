namespace OfficeIMO.Epub;

public sealed partial class EpubPublication {
    private PackageJournal? _packageJournal;
    private bool _packageRollback;
    // Upper bound of the serialized package size in UTF-8 bytes, or -1 when unknown. Budget checks
    // serialize the package only when this bound cannot prove the edit fits.
    private long _packageBytesBound = -1;
    // Package growth that Write will add on top of the live package, such as content properties it
    // derives from edited documents. Budget checks during an edit reserve this much room.
    private long _packageGrowthAllowance;
    // Elements loaded as <a></a>. Adding a child replaces that empty string without raising an event, so
    // undoing the add must restore it. Collected on first use inside an edit, then kept current.
    private HashSet<XElement>? _emptyStringElements;

    // Applies an edit to the live package and checks its budgets. A rejected edit replays the recorded
    // changes backwards, so the package, its declaration wrappers and the change state stay as they were.
    // The cost depends on what the edit touches, not on the size of the package.
    private void EditPackage(Action edit, long payloadDelta = 0) {
        if (_packageJournal != null) {
            edit();
            EnsureLivePackageBudget(payloadDelta);
            return;
        }
        var journal = new PackageJournal();
        long boundBefore = _packageBytesBound;
        _packageJournal = journal;
        bool committed = false;
        try {
            edit();
            EnsureLivePackageBudget(payloadDelta);
            committed = true;
        } finally {
            _packageJournal = null;
            _packageGrowthAllowance = 0;
            if (!committed) {
                _packageRollback = true;
                try { journal.Rollback(); } finally { _packageRollback = false; InvalidatePackageIndexes(); }
                _packageBytesBound = boundBefore;
            } else if (journal.HasChanges) MarkChanged();
        }
    }

    private void EnsureLivePackageBudget(long payloadDelta) {
        long available = _maximumRetainedBytes - RetainedPayloadBytes - payloadDelta;
        long maximum = Math.Min(_maximumMetadataBytes, Math.Min(_maximumEntryBytes, available));
        maximum -= _packageGrowthAllowance;
        if (maximum < 1) throw new InvalidDataException("Package XML exceeds the publication's retained-byte limits.");
        if (_packageBytesBound >= 0 && _packageBytesBound <= maximum) return;
        _packageBytesBound = SerializeXml(_package, maximum).LongLength;
    }

    private void OnPackageChanging(object? sender, XObjectChangeEventArgs args) {
        if (_packageRollback) return;
        IndexPackageChanging(sender, args.ObjectChange);
        if (_packageJournal == null) { _packageBytesBound = -1; return; }
        if (_emptyStringElements == null && args.ObjectChange == XObjectChange.Add)
            _emptyStringElements = new HashSet<XElement>(_package.Descendants().Where(HasEmptyStringContent));
        long growth = _packageJournal.Record(sender as XObject, args.ObjectChange);
        _packageBytesBound = _packageBytesBound < 0 || growth < 0 ? -1 : _packageBytesBound + growth;
    }

    private void OnPackageChanged(object? sender, XObjectChangeEventArgs args) {
        if (_packageRollback) return;
        IndexPackageChanged(sender, args.ObjectChange);
        TrackEmptyStringContent(sender, args.ObjectChange);
        if (_packageJournal == null) { MarkChanged(); return; }
        long growth = PackageJournal.Completed(sender as XObject, args.ObjectChange);
        _packageBytesBound = _packageBytesBound < 0 || growth < 0 ? -1 : _packageBytesBound + growth;
    }

    private static bool HasEmptyStringContent(XElement element) =>
        !element.IsEmpty && !element.HasElements && element.Value.Length == 0 && element.LastNode == null;

    private void TrackEmptyStringContent(object? sender, XObjectChange change) {
        if (_emptyStringElements == null) return;
        if (sender is XElement element && change == XObjectChange.Value) {
            if (HasEmptyStringContent(element)) _emptyStringElements.Add(element); else _emptyStringElements.Remove(element);
        } else if (sender is XNode node && change == XObjectChange.Add && node.Parent is XElement parent &&
            _emptyStringElements.Remove(parent) && _packageJournal != null) {
            _packageJournal.RestoreEmptyStringOnUndo(node, parent);
        }
    }

    // Records enough state before each XML change to undo it exactly, and estimates how much each change
    // can grow the serialized package. A negative estimate means the growth is unknown.
    private sealed class PackageJournal {
        private readonly List<Action> _undo = new();

        internal bool HasChanges => _undo.Count != 0;

        // Called right after the node's add was recorded, so the last undo entry is that add.
        internal void RestoreEmptyStringOnUndo(XNode node, XElement parent) =>
            _undo[_undo.Count - 1] = () => { node.Remove(); if (parent.IsEmpty) parent.Add(string.Empty); };

        internal long Record(XObject? sender, XObjectChange change) {
            switch (sender) {
                case XAttribute attribute when change == XObjectChange.Add:
                    _undo.Add(() => attribute.Remove());
                    return 0;
                case XAttribute attribute when change == XObjectChange.Remove: {
                    XElement parent = attribute.Parent!;
                    XAttribute[] attributes = parent.Attributes().ToArray();
                    _undo.Add(() => parent.ReplaceAttributes(attributes.Cast<object>().ToArray()));
                    return attribute.IsNamespaceDeclaration ? -1 : 0;
                }
                case XAttribute attribute when change == XObjectChange.Value: {
                    string value = attribute.Value;
                    _undo.Add(() => attribute.Value = value);
                    return 0;
                }
                case XNode node when change == XObjectChange.Add:
                    _undo.Add(() => node.Remove());
                    return 0;
                case XNode node when change == XObjectChange.Remove: {
                    XContainer parent = (XContainer?)node.Parent ?? node.Document!;
                    XNode? next = node.NextNode;
                    _undo.Add(() => { if (next != null) next.AddBeforeSelf(node); else parent.Add(node); });
                    return 0;
                }
                case XElement element when change == XObjectChange.Name: {
                    XName name = element.Name;
                    _undo.Add(() => element.Name = name);
                    return -1;
                }
                case XElement element when change == XObjectChange.Value: {
                    // Raised only when content switches between empty (<a/>) and an empty string (<a></a>).
                    bool wasEmpty = element.IsEmpty;
                    _undo.Add(() => { if (wasEmpty) element.RemoveNodes(); else element.Add(string.Empty); });
                    return 0;
                }
                case XText text when change == XObjectChange.Value: {
                    string value = text.Value;
                    _undo.Add(() => text.Value = value);
                    return 0;
                }
                case XComment comment when change == XObjectChange.Value: {
                    string value = comment.Value;
                    _undo.Add(() => comment.Value = value);
                    return 0;
                }
                case XProcessingInstruction instruction when change == XObjectChange.Value: {
                    string data = instruction.Data;
                    _undo.Add(() => instruction.Data = data);
                    return 0;
                }
                default:
                    throw new NotSupportedException("Package edit changes XML in a way that cannot be rolled back.");
            }
        }

        // Growth is estimated once the change is in place. An attached element serializes with the
        // prefixes it has in the package plus any declarations it inherits, so its text is an upper bound.
        internal static long Completed(XObject? sender, XObjectChange change) {
            if (change == XObjectChange.Add) {
                return sender switch {
                    XAttribute attribute => AttributeBound(attribute),
                    XText text => EscapedBound(text.Value) + EndTagBound(text.Parent),
                    XNode node => Encoding.UTF8.GetByteCount(node.ToString(SaveOptions.DisableFormatting)) + 256 + EndTagBound(node.Parent),
                    _ => -1
                };
            }
            if (change != XObjectChange.Value) return 0;
            return sender switch {
                XAttribute attribute => EscapedBound(attribute.Value),
                XText text => EscapedBound(text.Value),
                XComment comment => EscapedBound(comment.Value),
                XProcessingInstruction instruction => EscapedBound(instruction.Data),
                XElement element => EndTagBound(element),
                _ => -1
            };
        }

        internal void Rollback() {
            for (int index = _undo.Count - 1; index >= 0; index--) _undo[index]();
        }

        // Content added to an empty element turns <a/> into <a>...</a>.
        private static long EndTagBound(XElement? element) {
            if (element == null) return 0;
            string prefix = element.GetPrefixOfNamespace(element.Name.Namespace) ?? new string('p', 16);
            return 2L * Encoding.UTF8.GetByteCount(prefix + ":" + element.Name.LocalName) + 8;
        }

        // A namespaced attribute without a declared prefix makes the writer declare one on its element.
        private static long AttributeBound(XAttribute attribute) {
            long size = Encoding.UTF8.GetByteCount(attribute.Name.LocalName) + 4 + EscapedBound(attribute.Value);
            XNamespace ns = attribute.Name.Namespace;
            if (ns == XNamespace.None || attribute.IsNamespaceDeclaration) return size;
            string? prefix = attribute.Parent?.GetPrefixOfNamespace(ns);
            return size + (prefix != null ? Encoding.UTF8.GetByteCount(prefix) + 1 : 64 + EscapedBound(ns.NamespaceName));
        }

        // Escaping expands one character into at most six (&quot;), each up to three UTF-8 bytes.
        private static long EscapedBound(string value) => (long)value.Length * 18;
    }
}
