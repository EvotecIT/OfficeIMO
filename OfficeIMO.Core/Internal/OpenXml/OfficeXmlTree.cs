namespace OfficeIMO.Core.Internal.OpenXml;

internal enum OfficeXmlTreeNodeKind : byte {
    Element,
    Text
}

/// <summary>
/// An editable XML part kept as arrays of node and attribute records instead of one object per element. A parsed
/// tree remembers where each node came from in the part's bytes, so the writer copies unchanged subtrees, start tags,
/// text, comments, processing instructions and whitespace exactly as they were written; only what an edit touched is
/// serialised again.
/// </summary>
/// <remarks>
/// Nodes are addressed by index. Removing a node detaches it; its record stays until the tree is dropped. Markup
/// compatibility content is kept as written: <c>mc:AlternateContent</c> and elements in ignorable namespaces are
/// ordinary elements here.
/// </remarks>
internal sealed partial class OfficeXmlTree {
    internal const int NoNode = -1;

    private const byte StartTagChanged = 1;
    private const byte ContentChanged = 2;
    private const byte SubtreeChanged = 4;
    private const byte CData = 8;
    private const byte SelfClosing = 16;

    private Node[] _nodes;
    private int _nodeCount;
    // Source byte ranges of the nodes a parse produced; nodes created later have none.
    private Span[] _spans = Array.Empty<Span>();
    private int _sourceNodeCount;
    private Attr[] _attrs;
    private int _attrCount;
    private readonly byte[]? _source;
    private readonly int _sourceLength;
    private readonly int _prologEnd;
    private readonly int _epilogStart;
    // The value of the declaration's encoding of a part transcoded from UTF-16, which is written as UTF-8; -1 otherwise.
    private readonly int _encodingStart;
    private readonly int _encodingEnd;
    private readonly OfficeXmlPullReader? _decoder;

    private OfficeXmlTree(OfficeXmlNamespaces.Table namespaces, int nodeCapacity, int attributeCapacity) {
        Namespaces = namespaces;
        _nodes = new Node[Math.Max(4, nodeCapacity)];
        _attrs = new Attr[Math.Max(4, attributeCapacity)];
        Root = NoNode;
    }

    private OfficeXmlTree(OfficeXmlNamespaces.Table namespaces, OfficeXmlPullReader reader, int nodeCapacity, int attributeCapacity)
        : this(namespaces, nodeCapacity, attributeCapacity) {
        _decoder = reader;
        _source = reader.Data;
        _sourceLength = reader.DataLength;
        _spans = new Span[_nodes.Length];
        Parse(reader, out _prologEnd, out _epilogStart);
        _sourceNodeCount = _nodeCount;
        if (reader.Transcoded) FindDeclaredEncoding(out _encodingStart, out _encodingEnd);
        else _encodingStart = _encodingEnd = -1;
    }

    /// <summary>Namespace ids shared with the other parts of the package.</summary>
    internal OfficeXmlNamespaces.Table Namespaces { get; }

    /// <summary>The root element.</summary>
    internal int Root { get; private set; }

    /// <summary>Whether the tree was parsed from bytes rather than created empty.</summary>
    internal bool IsParsed => _source != null;

    /// <summary>Whether anything changed since the tree was parsed; a created tree always counts as changed.</summary>
    internal bool IsChanged => _source == null || Root != NoNode && (_nodes[Root].Flags & SubtreeChanged) != 0;

    /// <summary>Parses the UTF-8 or UTF-16 bytes of a part. The tree keeps <paramref name="data"/>; don't reuse it while the tree lives.</summary>
    /// <exception cref="System.Xml.XmlException">The part is not well-formed XML.</exception>
    internal static OfficeXmlTree Parse(byte[] data, int length, OfficeXmlNamespaces.Table namespaces) {
        var reader = new OfficeXmlPullReader(data, length, namespaces, processMarkupCompatibility: false);
        // About one element per 40 bytes and one attribute per element in typical PresentationML.
        return new OfficeXmlTree(namespaces, reader, length / 40 + 8, length / 40 + 8);
    }

    /// <summary>Creates a tree with only a root element.</summary>
    internal static OfficeXmlTree Create(OfficeXmlNamespaces.Table namespaces, int namespaceId, string localName, int capacity = 32) {
        var tree = new OfficeXmlTree(namespaces, capacity, capacity);
        tree.Root = tree.CreateElement(namespaceId, localName);
        return tree;
    }

    private void Parse(OfficeXmlPullReader reader, out int prologEnd, out int epilogStart) {
        prologEnd = 0;
        epilogStart = 0;
        int[] open = new int[16];
        int[] lastEnd = new int[16];
        int depth = 0;
        while (reader.Read()) {
            switch (reader.NodeKind) {
                case OfficeXmlNodeKind.StartElement: {
                    int node = NewNode(OfficeXmlTreeNodeKind.Element);
                    ref Node record = ref _nodes[node];
                    record.Name = Namespaces.GetNameId(reader.LocalName, reader.ElementPrefix, reader.NamespaceId, reader.ElementUriId);
                    ref Span span = ref SourceSpan(node);
                    span.StartTagStart = reader.MarkupStart;
                    span.StartTagEnd = reader.MarkupEnd;
                    span.GapStart = -1;
                    if (reader.IsEmptyElement) record.Flags = SelfClosing;
                    int count = reader.AttributeCount;
                    if (count > 0) {
                        record.AttrFirst = ReserveAttributes(count);
                        record.AttrCount = record.AttrCapacity = count;
                        for (int index = 0; index < count; index++) {
                            ref Attr attribute = ref _attrs[record.AttrFirst + index];
                            int attributeNamespace = reader.GetAttributeNamespace(index, out int attributeUri);
                            attribute.Name = Namespaces.GetNameId(reader.GetAttributeLocalName(index), reader.GetAttributePrefix(index), attributeNamespace, attributeUri);
                            reader.GetAttributeRawValue(index, out attribute.ValueStart, out attribute.ValueLength);
                            attribute.Quote = reader.GetAttributeQuote(index);
                        }
                    }
                    if (depth == 0) {
                        Root = node;
                        prologEnd = reader.MarkupStart;
                    } else {
                        span.GapStart = lastEnd[depth - 1];
                        Link(open[depth - 1], node, NoNode);
                    }
                    if (depth == open.Length) {
                        Array.Resize(ref open, depth * 2);
                        Array.Resize(ref lastEnd, depth * 2);
                    }
                    open[depth] = node;
                    lastEnd[depth] = reader.MarkupEnd;
                    depth++;
                    break;
                }
                case OfficeXmlNodeKind.EndElement: {
                    depth--;
                    ref Span span = ref _spans[open[depth]];
                    span.TrailingGapStart = lastEnd[depth];
                    span.EndTagStart = reader.MarkupStart;
                    span.EndTagEnd = reader.MarkupEnd;
                    if (depth > 0) lastEnd[depth - 1] = reader.MarkupEnd;
                    else epilogStart = reader.MarkupEnd;
                    break;
                }
                case OfficeXmlNodeKind.Text: {
                    int node = NewNode(OfficeXmlTreeNodeKind.Text);
                    ref Span span = ref SourceSpan(node);
                    span.StartTagStart = reader.MarkupStart;
                    span.StartTagEnd = reader.MarkupEnd;
                    span.GapStart = lastEnd[depth - 1];
                    if (reader.IsCData) _nodes[node].Flags = CData;
                    Link(open[depth - 1], node, NoNode);
                    lastEnd[depth - 1] = reader.MarkupEnd;
                    break;
                }
            }
        }
    }

    // ---- Navigation -------------------------------------------------------------------------------------------

    internal OfficeXmlTreeNodeKind KindOf(int node) => _nodes[node].Kind;
    internal int ParentOf(int node) => _nodes[node].Parent;
    internal int FirstChildOf(int node) => _nodes[node].FirstChild;
    internal int LastChildOf(int node) => _nodes[node].LastChild;
    internal int NextSiblingOf(int node) => _nodes[node].Next;
    internal int PreviousSiblingOf(int node) => _nodes[node].Previous;
    internal string LocalNameOf(int node) => Namespaces.GetName(_nodes[node].Name).LocalName;
    internal int NamespaceOf(int node) => Namespaces.GetName(_nodes[node].Name).NamespaceId;

    internal bool IsElement(int node, int namespaceId, string localName) {
        ref Node record = ref _nodes[node];
        if (record.Kind != OfficeXmlTreeNodeKind.Element) return false;
        ref readonly OfficeXmlName name = ref Namespaces.GetName(record.Name);
        return name.NamespaceId == namespaceId && string.Equals(name.LocalName, localName, StringComparison.Ordinal);
    }

    /// <summary>The first child element with this name, or <see cref="NoNode"/>.</summary>
    internal int FindChild(int parent, int namespaceId, string localName) {
        for (int child = _nodes[parent].FirstChild; child != NoNode; child = _nodes[child].Next) {
            if (IsElement(child, namespaceId, localName)) return child;
        }
        return NoNode;
    }

    /// <summary>The next sibling element after <paramref name="node"/> with this name, or <see cref="NoNode"/>.</summary>
    internal int FindNextSibling(int node, int namespaceId, string localName) {
        for (int sibling = _nodes[node].Next; sibling != NoNode; sibling = _nodes[sibling].Next) {
            if (IsElement(sibling, namespaceId, localName)) return sibling;
        }
        return NoNode;
    }

    /// <summary>The decoded value of a text node.</summary>
    internal string GetText(int node) {
        ref Node record = ref _nodes[node];
        if (record.Text != null) return record.Text;
        bool cdata = (record.Flags & CData) != 0;
        ref Span span = ref _spans[node];
        int start = span.StartTagStart + (cdata ? 9 : 0);
        int length = span.StartTagEnd - span.StartTagStart - (cdata ? 12 : 0);
        return _decoder!.DecodeTextValue(start, length, cdata);
    }

    /// <summary>The text of the element's text children, joined, or an empty string.</summary>
    internal string GetInnerText(int element) {
        string? single = null;
        System.Text.StringBuilder? builder = null;
        for (int child = _nodes[element].FirstChild; child != NoNode; child = _nodes[child].Next) {
            if (_nodes[child].Kind != OfficeXmlTreeNodeKind.Text) continue;
            string text = GetText(child);
            if (single == null) single = text;
            else (builder ??= new System.Text.StringBuilder(single)).Append(text);
        }
        return builder?.ToString() ?? single ?? string.Empty;
    }

    // ---- Attributes ------------------------------------------------------------------------------------------

    internal int AttributeCountOf(int node) => _nodes[node].AttrCount;

    internal bool TryGetAttribute(int node, int namespaceId, string localName, out string value) {
        int index = FindAttribute(node, namespaceId, localName);
        if (index < 0) {
            value = string.Empty;
            return false;
        }
        value = AttributeValue(ref _attrs[index]);
        return true;
    }

    internal string? GetAttribute(int node, int namespaceId, string localName) =>
        TryGetAttribute(node, namespaceId, localName, out string value) ? value : null;

    /// <summary>Sets an attribute, keeping its position when it exists and appending it otherwise.</summary>
    internal void SetAttribute(int node, int namespaceId, string localName, string value) {
        ArgumentNullException.ThrowIfNull(value);
        int index = FindAttribute(node, namespaceId, localName);
        if (index < 0) {
            index = AppendAttribute(node);
            ref Attr added = ref _attrs[index];
            added.Name = Namespaces.GetNameId(localName, namespaceId == OfficeXmlNamespaces.None ? string.Empty : null, namespaceId,
                OfficeXmlNamespaces.TransitionalUriId(namespaceId));
            added.ValueStart = -1;
        } else if (string.Equals(AttributeValue(ref _attrs[index]), value, StringComparison.Ordinal)) {
            return;
        }
        _attrs[index].Value = value;
        _attrs[index].Changed = true;
        MarkChanged(node, StartTagChanged);
    }

    internal bool RemoveAttribute(int node, int namespaceId, string localName) {
        int index = FindAttribute(node, namespaceId, localName);
        if (index < 0) return false;
        ref Node record = ref _nodes[node];
        int end = record.AttrFirst + record.AttrCount;
        Array.Copy(_attrs, index + 1, _attrs, index, end - index - 1);
        _attrs[end - 1] = default;
        record.AttrCount--;
        MarkChanged(node, StartTagChanged);
        return true;
    }

    /// <summary>Declares <paramref name="prefix"/> (empty for the default namespace) on an element.</summary>
    internal void DeclareNamespace(int node, string prefix, string uri) {
        int namespaceId = Namespaces.GetId(uri, out int uriId);
        int index = AppendAttribute(node);
        ref Attr declaration = ref _attrs[index];
        declaration.Name = Namespaces.GetNameId(prefix.Length == 0 ? "xmlns" : prefix, prefix.Length == 0 ? string.Empty : "xmlns",
            OfficeXmlNamespaces.Xmlns, OfficeXmlNamespaces.Xmlns);
        declaration.Value = uri;
        declaration.DeclaredNamespace = namespaceId;
        declaration.DeclaredUri = uriId;
        declaration.ValueStart = -1;
        declaration.Changed = true;
        MarkChanged(node, StartTagChanged);
    }

    private int FindAttribute(int node, int namespaceId, string localName) {
        ref Node record = ref _nodes[node];
        int end = record.AttrFirst + record.AttrCount;
        for (int index = record.AttrFirst; index < end; index++) {
            ref Attr attribute = ref _attrs[index];
            ref readonly OfficeXmlName name = ref Namespaces.GetName(attribute.Name);
            if (name.NamespaceId == namespaceId && string.Equals(name.LocalName, localName, StringComparison.Ordinal)) return index;
        }
        return -1;
    }

    private string AttributeValue(ref Attr attribute) =>
        attribute.Value ??= _decoder!.DecodeAttributeValue(attribute.ValueStart, attribute.ValueLength);

    private int AppendAttribute(int node) {
        ref Node record = ref _nodes[node];
        if (record.AttrCount == record.AttrCapacity) {
            if (record.AttrCount == 0 || record.AttrFirst + record.AttrCapacity == _attrCount) {
                // The block is the last one in the array (an element being built), so it grows in place.
                int first = ReserveAttributes(1);
                record = ref _nodes[node];
                if (record.AttrCount == 0) record.AttrFirst = first;
                record.AttrCapacity++;
            } else {
                // The block moves to the end of the array with room to grow; its old slots are left unused.
                int capacity = record.AttrCapacity * 2;
                int first = ReserveAttributes(capacity);
                record = ref _nodes[node];
                Array.Copy(_attrs, record.AttrFirst, _attrs, first, record.AttrCount);
                record.AttrFirst = first;
                record.AttrCapacity = capacity;
            }
        }
        int index = record.AttrFirst + record.AttrCount++;
        _attrs[index] = default;
        return index;
    }

    private int ReserveAttributes(int count) {
        if (_attrCount + count > _attrs.Length) Array.Resize(ref _attrs, Math.Max(_attrs.Length * 2, _attrCount + count));
        int first = _attrCount;
        _attrCount += count;
        return first;
    }

    // ---- Structure -------------------------------------------------------------------------------------------

    /// <summary>Creates a detached element; insert it with <see cref="Append"/> or <see cref="InsertBefore"/>.</summary>
    /// <remarks>Its prefix is chosen when it is written: the innermost prefix in scope for its namespace, else the
    /// namespace's usual prefix, declared on the element.</remarks>
    internal int CreateElement(int namespaceId, string localName) {
        int node = NewNode(OfficeXmlTreeNodeKind.Element);
        ref Node record = ref _nodes[node];
        record.Name = Namespaces.GetNameId(localName, null, namespaceId, OfficeXmlNamespaces.TransitionalUriId(namespaceId));
        record.Flags = StartTagChanged | ContentChanged | SubtreeChanged;
        return node;
    }

    /// <summary>Creates a detached text node.</summary>
    internal int CreateText(string text) {
        int node = NewNode(OfficeXmlTreeNodeKind.Text);
        ref Node record = ref _nodes[node];
        record.Text = text;
        record.Flags = ContentChanged | SubtreeChanged;
        return node;
    }

    /// <summary>Replaces the value of a text node.</summary>
    internal void SetText(int node, string text) {
        ArgumentNullException.ThrowIfNull(text);
        ref Node record = ref _nodes[node];
        if (record.Kind != OfficeXmlTreeNodeKind.Text) throw new InvalidOperationException("Only a text node has a value.");
        if (string.Equals(GetText(node), text, StringComparison.Ordinal)) return;
        _nodes[node].Text = text;
        _nodes[node].Flags &= unchecked((byte)~CData);
        MarkChanged(node, ContentChanged);
    }

    /// <summary>Replaces the children of an element with one text node; an empty string leaves no children.</summary>
    internal void SetInnerText(int element, string text) {
        int first = _nodes[element].FirstChild;
        if (first != NoNode && _nodes[first].Kind == OfficeXmlTreeNodeKind.Text && _nodes[first].Next == NoNode) {
            if (text.Length > 0) {
                SetText(first, text);
                return;
            }
        }
        while (_nodes[element].FirstChild != NoNode) Remove(_nodes[element].FirstChild);
        if (text.Length > 0) Append(element, CreateText(text));
    }

    internal void Append(int parent, int child) => InsertBefore(parent, child, NoNode);

    /// <summary>Inserts a detached node before <paramref name="reference"/>, or last when it is <see cref="NoNode"/>.</summary>
    internal void InsertBefore(int parent, int child, int reference) {
        if (_nodes[child].Parent != NoNode || child == Root) throw new InvalidOperationException("The node is already in the tree.");
        if (_nodes[parent].Kind != OfficeXmlTreeNodeKind.Element) throw new InvalidOperationException("Only an element has children.");
        Link(parent, child, reference);
        MarkChanged(parent, ContentChanged);
        MarkSubtree(child);
        // A parsed node that moves may rely on namespace declarations of its old ancestors, so every start tag in it is
        // written again, declaring what its new place lacks.
        if (child < _sourceNodeCount) MarkMoved(child);
    }

    private void MarkMoved(int node) {
        _nodes[node].Flags |= StartTagChanged | SubtreeChanged;
        for (int child = _nodes[node].FirstChild; child != NoNode; child = _nodes[child].Next) {
            if (_nodes[child].Kind == OfficeXmlTreeNodeKind.Element) MarkMoved(child);
        }
    }

    private void FindDeclaredEncoding(out int start, out int end) {
        start = end = -1;
        byte[] source = _source!;
        if (_prologEnd < 6 || source[0] != (byte)'<' || source[1] != (byte)'?' || source[2] != (byte)'x' || source[3] != (byte)'m' || source[4] != (byte)'l') return;
        int close = 5;
        while (close + 1 < _prologEnd && !(source[close] == (byte)'?' && source[close + 1] == (byte)'>')) close++;
        for (int index = 5; index + 9 < close; index++) {
            if (source[index] != (byte)'e' || !Matches(source, index, "encoding")) continue;
            int position = index + 8;
            while (position < close && (source[position] == (byte)' ' || source[position] == (byte)'\t' || source[position] == (byte)'\r' || source[position] == (byte)'\n')) position++;
            if (position >= close || source[position] != (byte)'=') continue;
            position++;
            while (position < close && (source[position] == (byte)' ' || source[position] == (byte)'\t' || source[position] == (byte)'\r' || source[position] == (byte)'\n')) position++;
            if (position >= close || source[position] != (byte)'"' && source[position] != (byte)'\'') return;
            byte quote = source[position];
            int valueEnd = Array.IndexOf(source, quote, position + 1, close - position - 1);
            if (valueEnd < 0) return;
            start = position + 1;
            end = valueEnd;
            return;
        }

        static bool Matches(byte[] data, int at, string text) {
            for (int offset = 0; offset < text.Length; offset++) {
                if (data[at + offset] != (byte)text[offset]) return false;
            }
            return true;
        }
    }

    /// <summary>Detaches a node and its subtree.</summary>
    internal void Remove(int node) {
        ref Node record = ref _nodes[node];
        int parent = record.Parent;
        if (parent == NoNode) throw new InvalidOperationException("The root cannot be removed.");
        if (record.Previous != NoNode) _nodes[record.Previous].Next = record.Next;
        else _nodes[parent].FirstChild = record.Next;
        if (record.Next != NoNode) _nodes[record.Next].Previous = record.Previous;
        else _nodes[parent].LastChild = record.Previous;
        record.Parent = record.Next = record.Previous = NoNode;
        // Comments and instructions written before the node leave with it rather than reappear where it is inserted.
        if (node < _sourceNodeCount) _spans[node].GapStart = -1;
        MarkChanged(parent, ContentChanged);
    }

    /// <summary>
    /// Inserts a detached element among the children of <paramref name="parent"/> in schema order:
    /// before the first child element whose rank in <paramref name="order"/> is higher than its own.
    /// </summary>
    /// <param name="parent">The element that receives the child.</param>
    /// <param name="child">A detached element.</param>
    /// <param name="order">The child element names of the parent's schema type, in sequence order.</param>
    internal void InsertOrdered(int parent, int child, OfficeXmlChildOrder order) {
        int rank = RankOf(order, child);
        int reference = NoNode;
        if (rank >= 0) {
            for (int sibling = _nodes[parent].FirstChild; sibling != NoNode; sibling = _nodes[sibling].Next) {
                if (_nodes[sibling].Kind != OfficeXmlTreeNodeKind.Element) continue;
                int siblingRank = RankOf(order, sibling);
                if (siblingRank > rank) {
                    reference = sibling;
                    break;
                }
            }
        }
        InsertBefore(parent, child, reference);
    }

    /// <summary>Returns the child element with this name, creating it in schema order when it is missing.</summary>
    internal int GetOrInsertChild(int parent, int namespaceId, string localName, OfficeXmlChildOrder order) {
        int child = FindChild(parent, namespaceId, localName);
        if (child != NoNode) return child;
        child = CreateElement(namespaceId, localName);
        InsertOrdered(parent, child, order);
        return child;
    }

    private int RankOf(OfficeXmlChildOrder order, int node) {
        ref readonly OfficeXmlName name = ref Namespaces.GetName(_nodes[node].Name);
        return order.RankOf(name.NamespaceId, name.LocalName);
    }

    private void Link(int parent, int child, int reference) {
        ref Node record = ref _nodes[child];
        record.Parent = parent;
        if (reference == NoNode) {
            int last = _nodes[parent].LastChild;
            record.Previous = last;
            record.Next = NoNode;
            if (last != NoNode) _nodes[last].Next = child;
            else _nodes[parent].FirstChild = child;
            _nodes[parent].LastChild = child;
        } else {
            int previous = _nodes[reference].Previous;
            record.Previous = previous;
            record.Next = reference;
            _nodes[reference].Previous = child;
            if (previous != NoNode) _nodes[previous].Next = child;
            else _nodes[parent].FirstChild = child;
        }
    }

    private void MarkChanged(int node, byte flag) {
        _nodes[node].Flags |= flag;
        MarkSubtree(node);
    }

    // A changed node makes every ancestor's subtree changed, so the writer stops copying them verbatim.
    private void MarkSubtree(int node) {
        for (int current = node; current != NoNode; current = _nodes[current].Parent) {
            if ((_nodes[current].Flags & SubtreeChanged) != 0 && current != node) return;
            _nodes[current].Flags |= SubtreeChanged;
        }
    }

    private int NewNode(OfficeXmlTreeNodeKind kind) {
        if (_nodeCount == _nodes.Length) Array.Resize(ref _nodes, _nodes.Length * 2);
        int node = _nodeCount++;
        _nodes[node] = new Node {
            Kind = kind,
            Parent = NoNode,
            FirstChild = NoNode,
            LastChild = NoNode,
            Next = NoNode,
            Previous = NoNode
        };
        return node;
    }

    private struct Node {
        internal OfficeXmlTreeNodeKind Kind;
        internal byte Flags;
        internal int Parent;
        internal int FirstChild;
        internal int LastChild;
        internal int Next;
        internal int Previous;
        // Name id in the package's namespace table; a created element's name has no prefix until it is written.
        internal int Name;
        internal int AttrFirst;
        internal int AttrCount;
        internal int AttrCapacity;
        internal string? Text;
    }

    private ref Span SourceSpan(int node) {
        if (node >= _spans.Length) Array.Resize(ref _spans, _nodes.Length);
        return ref _spans[node];
    }

    // Where a parsed node was written. For text, StartTagStart..StartTagEnd is the text or CDATA section as written.
    private struct Span {
        internal int StartTagStart;
        internal int StartTagEnd;
        internal int EndTagStart;
        internal int EndTagEnd;
        // Bytes between the previous sibling (or the parent's start tag) and this node: comments, instructions.
        internal int GapStart;
        // Bytes between the last child and the end tag.
        internal int TrailingGapStart;
    }

    private struct Attr {
        // Name id in the package's namespace table; an attribute added in a namespace has no prefix until it is written.
        internal int Name;
        internal int ValueStart;
        internal int ValueLength;
        internal byte Quote;
        // The decoded or assigned value; Changed says whether it differs from the source.
        internal string? Value;
        internal bool Changed;
        // For a declaration added by DeclareNamespace: the namespace it binds.
        internal int DeclaredNamespace;
        internal int DeclaredUri;
    }
}

/// <summary>The child element names of a schema type in sequence order, for inserting children where the schema expects them.</summary>
/// <remarks>Alternatives of one choice share a rank: they may appear in any order among themselves.</remarks>
internal sealed class OfficeXmlChildOrder {
    private readonly int[] _namespaces;
    private readonly string[] _names;
    private readonly int[] _ranks;

    internal OfficeXmlChildOrder(int[] namespaces, string[] names, int[]? ranks = null) {
        if (namespaces.Length != names.Length || ranks != null && ranks.Length != names.Length) throw new ArgumentException("Each name needs a namespace and a rank.", nameof(names));
        _namespaces = namespaces;
        _names = names;
        _ranks = ranks ?? CreateSequentialRanks(names.Length);
    }

    /// <summary>Rank of a child name in the sequence, or -1 when the type does not list it.</summary>
    internal int RankOf(int namespaceId, string localName) {
        for (int index = 0; index < _names.Length; index++) {
            if (_namespaces[index] == namespaceId && string.Equals(_names[index], localName, StringComparison.Ordinal)) return _ranks[index];
        }
        return -1;
    }

    private static int[] CreateSequentialRanks(int count) {
        var ranks = new int[count];
        for (int index = 0; index < count; index++) ranks[index] = index;
        return ranks;
    }
}
