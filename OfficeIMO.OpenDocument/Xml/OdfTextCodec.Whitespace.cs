namespace OfficeIMO.OpenDocument;

internal static partial class OdfTextCodec {
    // ODF 1.4 section 6.1.2: only XML space characters collapse. Explicit native
    // whitespace tokens remain barriers; cached/opaque inline text retains its fallback.
    private static bool IsXmlSpace(char value) => value is ' ' or '\t' or '\r' or '\n';

    private static bool IsWhitespaceContainer(XElement element) => element.Name.Namespace == OdfNamespaces.Text &&
        element.Name.LocalName is "span" or "a" or "meta" or "meta-field" or "ruby" or "ruby-base";

    private static XElement? ParagraphOf(XNode node) {
        if (node is XElement element && OdfTextTraversal.IsParagraph(element)) return element;
        int depth = 0;
        foreach (XElement ancestor in node.Ancestors()) {
            if (OdfTextTraversal.IsParagraph(ancestor)) return ancestor;
            // Opaque values and other stories do not acquire an outer paragraph's rules.
            if (!IsWhitespaceContainer(ancestor)) return null;
            if (++depth > OdfTextTraversal.MaximumContainerDepth)
                throw new NotSupportedException("OpenDocument text context exceeds the 128-container safety limit.");
        }
        return null;
    }

    private static bool IsCollapsibleText(XText text, XElement? paragraph) {
        if (paragraph == null) return false;
        foreach (XElement ancestor in text.Ancestors()) {
            if (ancestor == paragraph) return true;
            if (!IsWhitespaceContainer(ancestor)) return false;
        }
        return false;
    }

    private static bool IsWhitespaceToken(XElement element) => element.Name == OdfNamespaces.Text + "s" ||
        element.Name == OdfNamespaces.Text + "tab" || element.Name == OdfNamespaces.Text + "line-break";

    // Fragment getters need paragraph context, but must not decode the complete story
    // for every run. Search only until the adjacent meaningful token is found.
    private static XNode? AdjacentLeaf(XNode anchor, XElement paragraph, bool forward, ref int visited) {
        XNode? cursor = Move(anchor);
        while (cursor != null) {
            if (++visited > OdfTextTraversal.MaximumVisitedElements)
                throw new NotSupportedException("OpenDocument text context exceeds the node safety limit.");
            if (cursor.Ancestors().TakeWhile(a => a != paragraph).Take(OdfTextTraversal.MaximumContainerDepth + 1).Count() > OdfTextTraversal.MaximumContainerDepth)
                throw new NotSupportedException("OpenDocument text context exceeds the 128-container safety limit.");
            if (cursor is XText || cursor is XElement token && IsWhitespaceToken(token)) return cursor;
            if (cursor is XElement container && !IsNonVisibleTextElement(container)) {
                XNode? child = forward ? container.FirstNode : container.LastNode;
                if (child != null) { cursor = child; continue; }
            }
            cursor = Move(cursor);
        }
        return null;

        XNode? Move(XNode node) {
            while (node != paragraph) {
                XNode? sibling = forward ? node.NextNode : node.PreviousNode;
                if (sibling != null) return sibling;
                if (node.Parent == null || node.Parent == paragraph) return null;
                node = node.Parent;
            }
            return null;
        }
    }

    private static TextSnapshot DecodeNodes(IEnumerable<XNode> nodes, int maximumCharacters, Action? onVisit = null,
        Func<XElement, string, string?>? projectField = null) {
        if (nodes == null) throw new ArgumentNullException(nameof(nodes));
        var selected = new List<XNode>();
        foreach (XNode node in nodes) {
            if (selected.Count >= OdfTextTraversal.MaximumVisitedElements)
                throw new NotSupportedException($"OpenDocument text decoding exceeds the {OdfTextTraversal.MaximumVisitedElements}-node safety limit.");
            selected.Add(node);
        }
        var snapshot = new TextSnapshot(maximumCharacters);
        if (selected.Count == 0) return snapshot;
        XElement? paragraph = ParagraphOf(selected[0]);
        if (paragraph != ParagraphOf(selected[selected.Count - 1])) paragraph = null;
        var whitespace = new WhitespaceDecoder(snapshot, paragraph, selected[0], selected[selected.Count - 1]);
        foreach (XNode node in VisibleNodes(selected, onVisit, projectField == null ? null : IsScalarField)) {
            if (node is XText text) {
                snapshot.Register(node, text.Value.Length);
                if (IsCollapsibleText(text, paragraph)) whitespace.Literal(text);
                else whitespace.Explicit(node, text.Value);
            } else if (node is XElement element) {
                if (projectField != null && IsScalarField(element)) {
                    var cache = new StringBuilder();
                    foreach (XText part in element.Nodes().OfType<XText>()) {
                        EnsureCapacity(cache, part.Value.Length, maximumCharacters - snapshot.SourceCharacters);
                        cache.Append(part.Value);
                    }
                    string value = projectField(element, cache.ToString()) ?? cache.ToString();
                    // Shorter field output does not refund the source budget; expansion is
                    // charged as well, before it can enter styled layout.
                    snapshot.Register(node, Math.Max(cache.Length, value.Length));
                    whitespace.Explicit(node, value);
                } else if (element.Name == OdfNamespaces.Text + "s") {
                    int count = ParsePositiveCount((string?)element.Attribute(OdfNamespaces.Text + "c"));
                    snapshot.Register(node, count);
                    whitespace.Explicit(node, new string(' ', count));
                } else if (element.Name == OdfNamespaces.Text + "tab" || element.Name == OdfNamespaces.Text + "line-break") {
                    snapshot.Register(node, 1);
                    whitespace.Explicit(node, element.Name.LocalName == "tab" ? "\t" : "\n");
                }
            }
        }
        whitespace.Finish();
        return snapshot;
    }

    /// <summary>Resolves scalar fields before whitespace normalization without modifying their source XML.</summary>
    internal static TextSnapshot Snapshot(XElement paragraph, int maximumCharacters = MaximumDecodedCharacters, Action? onVisit = null,
        Func<XElement, string, string?>? projectField = null) => DecodeNodes(paragraph.Nodes(), maximumCharacters, onVisit, projectField);

    private static bool IsScalarField(XElement element) => OdfTextField.IsField(element.Name) && !element.HasElements;

    /// <summary>A bounded, immutable text contribution map shared by native views and styled projection.</summary>
    internal sealed class TextSnapshot {
        private readonly int _maximumCharacters;
        private readonly Dictionary<XNode, string> _text = new();
        private readonly Dictionary<XNode, int> _sourceCharacters = new();
        private readonly List<XNode> _order = new();
        internal TextSnapshot(int maximumCharacters) { _maximumCharacters = maximumCharacters; }
        internal int SourceCharacters { get; private set; }
        internal string Text => JoinBounded(_order.Select(node => _text.TryGetValue(node, out string? value) ? value : string.Empty), string.Empty);
        internal void Register(XNode node, int characters) {
            if (characters > _maximumCharacters - SourceCharacters)
                throw new InvalidDataException($"Decoded OpenDocument text exceeds the {_maximumCharacters}-character safety limit.");
            SourceCharacters += characters; _sourceCharacters.Add(node, characters); _order.Add(node);
        }
        internal void Set(XNode node, string value) { _text[node] = value; }
        internal void AddSpace(XNode node) { _text[node] = (_text.TryGetValue(node, out string? value) ? value : string.Empty) + " "; }
        internal string ReadNodes(IEnumerable<XNode> nodes, ref int remainingCharacters) {
            var builder = new StringBuilder(); int consumed = 0;
            foreach (XNode node in VisibleNodes(nodes, atomic: element => IsScalarField(element) && _sourceCharacters.ContainsKey(element))) {
                if (!_sourceCharacters.TryGetValue(node, out int count)) continue;
                if (count > remainingCharacters - consumed)
                    throw new InvalidDataException($"Decoded OpenDocument text exceeds the {remainingCharacters}-character safety limit.");
                consumed += count;
                if (_text.TryGetValue(node, out string? value)) builder.Append(value);
            }
            remainingCharacters -= consumed; return builder.ToString();
        }
        internal string Read(XElement element, ref int remainingCharacters) => IsNonVisibleTextElement(element)
            ? OdfTextCodec.Read(element, ref remainingCharacters) : ReadNodes(new[] { element }, ref remainingCharacters);
        internal string ReadNode(XNode node) {
            int remaining = _maximumCharacters;
            return ReadNodes(new[] { node }, ref remaining);
        }
    }

    private sealed class WhitespaceDecoder {
        private readonly TextSnapshot _snapshot;
        private readonly XElement? _paragraph;
        private readonly XNode _last;
        private bool _hasContent;
        private bool _outsidePending;
        private XText? _pending;
        internal WhitespaceDecoder(TextSnapshot snapshot, XElement? paragraph, XNode first, XNode last) {
            _snapshot = snapshot; _paragraph = paragraph; _last = last;
            if (paragraph == null) return;
            int visited = 0; XNode? before = first;
            while ((before = AdjacentLeaf(before, paragraph, forward: false, ref visited)) != null) {
                if (before is XText text && IsCollapsibleText(text, paragraph)) {
                    for (int i = text.Value.Length - 1; i >= 0; i--) {
                        if (IsXmlSpace(text.Value[i])) { _outsidePending = true; continue; }
                        _hasContent = true; return;
                    }
                } else if (before is not XText raw || raw.Value.Length != 0) { _hasContent = true; return; }
            }
            _outsidePending = false;
        }
        internal void Literal(XText node) {
            var builder = new StringBuilder();
            foreach (char value in node.Value) {
                if (IsXmlSpace(value)) {
                    if (_hasContent && _pending == null && !_outsidePending) _pending = node;
                    continue;
                }
                Flush(builder, node); builder.Append(value); _hasContent = true;
            }
            _snapshot.Set(node, builder.ToString());
        }
        internal void Explicit(XNode node, string value) {
            if (value.Length == 0) { _snapshot.Set(node, value); return; }
            Flush(); _snapshot.Set(node, value); _hasContent = true;
        }
        private void Flush(StringBuilder? current = null, XText? owner = null) {
            if (_pending != null) {
                if (_pending == owner) current!.Append(' ');
                else _snapshot.AddSpace(_pending);
                _pending = null;
            }
            _outsidePending = false;
        }
        internal void Finish() {
            if (_pending == null || _paragraph == null) return;
            int visited = 0; XNode? after = _last;
            while ((after = AdjacentLeaf(after, _paragraph, forward: true, ref visited)) != null) {
                if (after is XText text && IsCollapsibleText(text, _paragraph) && text.Value.All(IsXmlSpace)) continue;
                if (after is XText empty && empty.Value.Length == 0) continue;
                Flush(); return;
            }
        }
    }
}
