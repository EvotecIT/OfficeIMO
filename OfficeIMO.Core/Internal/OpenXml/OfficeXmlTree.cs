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
    // Bindings a moved or re-scoped element keeps from where it was written, so its names and any QName values in its
    // attributes or text mean what they meant there. The writer declares those its new scope binds differently.
    private Dictionary<int, List<CarriedBinding>>? _carried;

    /// <summary>Counts edits; a cache built from the tree is current while this has not changed.</summary>
    internal int Version { get; private set; }

    // Set when the node, span and attribute arrays come from the shared array pool; ReleasePooledStorage returns them.
    private readonly bool _pooled;

    private OfficeXmlTree(OfficeXmlNamespaces.Table namespaces, int nodeCapacity, int attributeCapacity, bool pooled = false) {
        Namespaces = namespaces;
        _pooled = pooled;
        _nodes = Allocate<Node>(Math.Max(4, nodeCapacity));
        _attrs = Allocate<Attr>(Math.Max(4, attributeCapacity));
        Root = NoNode;
    }

    private OfficeXmlTree(OfficeXmlNamespaces.Table namespaces, OfficeXmlPullReader reader, int nodeCapacity, int attributeCapacity, bool pooled = false)
        : this(namespaces, nodeCapacity, attributeCapacity, pooled) {
        _decoder = reader;
        _source = reader.Data;
        _sourceLength = reader.DataLength;
        _spans = Allocate<Span>(_nodes.Length);
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

    /// <summary>
    /// Parses a part into storage rented from the shared array pool, sized from the markup so it doesn't grow while it
    /// parses. The owner calls <see cref="ReleasePooledStorage"/> when the tree is no longer used; the tree must not be
    /// read after that.
    /// </summary>
    internal static OfficeXmlTree ParsePooled(byte[] data, int length, OfficeXmlNamespaces.Table namespaces) {
        var reader = new OfficeXmlPullReader(data, length, namespaces, processMarkupCompatibility: false);
        // Every element and text node starts at or after a '<' of its own or its neighbour, and every attribute has an '='.
        // Text may hold '=' too, so the attribute estimate is capped at a few per tag and grows past that when needed.
        int nodes = Count(data, length, (byte)'<') + 8;
        int attributes = (int)Math.Min(Count(data, length, (byte)'=') + 8L, nodes * 2L);
        return new OfficeXmlTree(namespaces, reader, nodes, attributes, pooled: true);
    }

    private static int Count(byte[] data, int length, byte value) {
#if NET8_0_OR_GREATER
        return System.MemoryExtensions.Count(new ReadOnlySpan<byte>(data, 0, length), value);
#else
        int count = 0;
        for (int index = 0; index < length; index++) {
            if (data[index] == value) count++;
        }
        return count;
#endif
    }

#if NET8_0_OR_GREATER
    private T[] Allocate<T>(int length) => _pooled ? System.Buffers.ArrayPool<T>.Shared.Rent(length) : new T[length];
#else
    // Older targets have no shared array pool here; pooled trees own their arrays like any other.
    private static T[] Allocate<T>(int length) => new T[length];
#endif

    private void Grow<T>(ref T[] array, int length) {
#if NET8_0_OR_GREATER
        if (_pooled) {
            T[] larger = System.Buffers.ArrayPool<T>.Shared.Rent(length);
            Array.Copy(array, larger, array.Length);
            Return(array, array.Length);
            array = larger;
            return;
        }
#endif
        Array.Resize(ref array, length);
    }

    private static void Return<T>(T[] array, int used) {
        if (array.Length == 0) return;
        // Records hold strings, which must not stay reachable from the pool.
        Array.Clear(array, 0, Math.Min(used, array.Length));
#if NET8_0_OR_GREATER
        System.Buffers.ArrayPool<T>.Shared.Return(array);
#endif
    }

    /// <summary>Returns pooled storage to the array pool. Does nothing for a tree that owns its arrays.</summary>
    internal void ReleasePooledStorage() {
        if (!_pooled || _nodes.Length == 0) return;
        Return(_nodes, _nodeCount);
        Return(_spans, _nodeCount);
        Return(_attrs, _attrCount);
        _nodes = Array.Empty<Node>();
        _spans = Array.Empty<Span>();
        _attrs = Array.Empty<Attr>();
        _nodeCount = _attrCount = 0;
        Root = NoNode;
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
        var names = new ParseNameCache();
        while (reader.Read()) {
            switch (reader.NodeKind) {
                case OfficeXmlNodeKind.StartElement: {
                    int node = NewNode(OfficeXmlTreeNodeKind.Element);
                    ref Node record = ref _nodes[node];
                    record.Name = names.GetNameId(Namespaces, reader.LocalName, reader.ElementPrefix, reader.NamespaceId, reader.ElementUriId);
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
                            attribute.Name = names.GetNameId(Namespaces, reader.GetAttributeLocalName(index), reader.GetAttributePrefix(index), attributeNamespace, attributeUri);
                            reader.GetAttributeRawValue(index, out attribute.ValueStart, out attribute.ValueLength);
                            attribute.Quote = reader.GetAttributeQuote(index);
                            if (attributeNamespace == OfficeXmlNamespaces.Xmlns) Declare(ref attribute, reader, index);
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

    /// <summary>
    /// Whether the parsed root element holds a comment or processing instruction. The tree keeps them as written
    /// between nodes rather than as nodes, so a caller that walks siblings as another reader sees them can check first.
    /// </summary>
    internal bool HasCommentsOrInstructions() {
        if (_source == null || Root == NoNode) return false;
        int end = _epilogStart;
        for (int index = Array.IndexOf(_source, (byte)'<', _prologEnd, end - _prologEnd); index >= 0 && index + 1 < end;
             index = Array.IndexOf(_source, (byte)'<', index + 1, end - index - 1)) {
            byte next = _source[index + 1];
            if (next == (byte)'?' || next == (byte)'!' && index + 3 < end && _source[index + 2] == (byte)'-' && _source[index + 3] == (byte)'-') return true;
        }
        return false;
    }

    internal OfficeXmlTreeNodeKind KindOf(int node) => _nodes[node].Kind;
    /// <summary>Whether a text node was written as a CDATA section.</summary>
    internal bool IsCData(int node) => (_nodes[node].Flags & CData) != 0;
    internal int ParentOf(int node) => _nodes[node].Parent;
    internal int FirstChildOf(int node) => _nodes[node].FirstChild;
    internal int LastChildOf(int node) => _nodes[node].LastChild;
    internal int NextSiblingOf(int node) => _nodes[node].Next;
    internal int PreviousSiblingOf(int node) => _nodes[node].Previous;
    internal string LocalNameOf(int node) => Namespaces.GetName(_nodes[node].Name).LocalName;
    /// <summary>The number of nodes, so node indices run from 0 to one less than this.</summary>
    internal int NodeCount => _nodeCount;

    internal int NamespaceOf(int node) => Namespaces.GetName(_nodes[node].Name).NamespaceId;
    /// <summary>The id of the exact namespace URI the element was written with; see <see cref="OfficeXmlNamespaces.TransitionalUriId"/>.</summary>
    internal int UriIdOf(int node) => Namespaces.GetName(_nodes[node].Name).UriId;

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

    /// <summary>The name of the element's attribute at <paramref name="index"/>, counting from zero in written order.</summary>
    internal ref readonly OfficeXmlName AttributeNameAt(int node, int index) =>
        ref Namespaces.GetName(_attrs[_nodes[node].AttrFirst + index].Name);

    /// <summary>The decoded value of the element's attribute at <paramref name="index"/>.</summary>
    internal string AttributeValueAt(int node, int index) => AttributeValue(ref _attrs[_nodes[node].AttrFirst + index]);

    /// <summary>
    /// Whether an element declares a namespace outside the ECMA-376 set. A copy of its descendants placed elsewhere keeps
    /// element and attribute names, which the writer binds again, but a prefix such a namespace lends to a QName value
    /// would be lost.
    /// </summary>
    internal bool DeclaresForeignNamespaces(int element) {
        int first = _nodes[element].AttrFirst, count = _nodes[element].AttrCount;
        for (int index = first; index < first + count; index++) {
            if (Namespaces.GetName(_attrs[index].Name).NamespaceId != OfficeXmlNamespaces.Xmlns) continue;
            if (!OfficeXmlNamespaces.IsEcma376(Namespaces.GetId(AttributeValue(ref _attrs[index])))) return true;
        }
        return false;
    }

    internal bool TryGetAttribute(int node, int namespaceId, string localName, out string value) {
        int index = FindAttribute(node, namespaceId, localName);
        if (index < 0) {
            value = string.Empty;
            return false;
        }
        value = AttributeValue(ref _attrs[index]);
        return true;
    }

    internal bool HasAttribute(int node, int namespaceId, string localName) => FindAttribute(node, namespaceId, localName) >= 0;

    /// <summary>
    /// Whether the attribute exists and its value is <paramref name="expected"/>. A plain ASCII value is compared as
    /// written, without decoding it into a string.
    /// </summary>
    internal bool AttributeEquals(int node, int namespaceId, string localName, string expected) {
        int index = FindAttribute(node, namespaceId, localName);
        if (index < 0) return false;
        ref Attr attribute = ref _attrs[index];
        if (attribute.Value == null && _source != null && attribute.ValueLength == expected.Length) {
            bool plain = true;
            for (int offset = 0; offset < expected.Length; offset++) {
                byte value = _source[attribute.ValueStart + offset];
                if (value == (byte)'&' || value < 0x20 || value >= 0x80) {
                    plain = false;
                    break;
                }
                if (value != expected[offset]) return false;
            }
            if (plain) return true;
        } else if (attribute.Value == null && _source != null
            && Array.IndexOf(_source, (byte)'&', attribute.ValueStart, attribute.ValueLength) < 0) {
            // Without references, an ASCII value of another length can only match after whitespace normalization. A
            // value with other characters has more bytes than characters, so it is decoded and compared.
            bool decode = false;
            for (int offset = 0; offset < attribute.ValueLength; offset++) {
                byte value = _source[attribute.ValueStart + offset];
                if (value < 0x20 || value >= 0x80) {
                    decode = true;
                    break;
                }
            }
            if (!decode) return false;
        }
        return string.Equals(AttributeValue(ref attribute), expected, StringComparison.Ordinal);
    }

    internal string? GetAttribute(int node, int namespaceId, string localName) =>
        TryGetAttribute(node, namespaceId, localName, out string value) ? value : null;

    /// <summary>Sets an attribute, keeping its position when it exists and appending it otherwise.</summary>
    internal void SetAttribute(int node, int namespaceId, string localName, string value) {
        ArgumentNullException.ThrowIfNull(value);
        int index = FindAttribute(node, namespaceId, localName);
        if (namespaceId == OfficeXmlNamespaces.Xmlns) {
            // A declaration: "xmlns" names the default namespace, any other local name a prefix.
            string prefix = localName == "xmlns" ? string.Empty : localName;
            if (index < 0) {
                DeclareNamespace(node, prefix, value);
                return;
            }
            if (string.Equals(AttributeValue(ref _attrs[index]), value, StringComparison.Ordinal)) return;
            CheckDeclaration(node, prefix, value);
            CarryBindingToChildren(node, prefix);
            BindDeclaration(ref _attrs[index], prefix, value);
            _attrs[index].Value = value;
            _attrs[index].Changed = true;
            MarkChanged(node, StartTagChanged);
            return;
        }
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
        if (namespaceId == OfficeXmlNamespaces.Xmlns) CarryBindingToChildren(node, localName == "xmlns" ? string.Empty : localName);
        ref Node record = ref _nodes[node];
        int end = record.AttrFirst + record.AttrCount;
        Array.Copy(_attrs, index + 1, _attrs, index, end - index - 1);
        _attrs[end - 1] = default;
        record.AttrCount--;
        MarkChanged(node, StartTagChanged);
        return true;
    }

    /// <summary>Declares <paramref name="prefix"/> (empty for the default namespace) on an element.</summary>
    /// <exception cref="InvalidOperationException">The element already binds the prefix to another namespace.</exception>
    internal void DeclareNamespace(int node, string prefix, string uri) {
        ArgumentNullException.ThrowIfNull(uri);
        int existing = FindAttribute(node, OfficeXmlNamespaces.Xmlns, prefix.Length == 0 ? "xmlns" : prefix);
        if (existing >= 0) {
            if (string.Equals(AttributeValue(ref _attrs[existing]), uri, StringComparison.Ordinal)) return;
            throw new InvalidOperationException("The element already declares the prefix '" + prefix + "' for another namespace.");
        }
        CheckDeclaration(node, prefix, uri);
        CarryBindingToChildren(node, prefix);
        int index = AppendAttribute(node);
        ref Attr declaration = ref _attrs[index];
        declaration.Name = Namespaces.GetNameId(prefix.Length == 0 ? "xmlns" : prefix, prefix.Length == 0 ? string.Empty : "xmlns",
            OfficeXmlNamespaces.Xmlns, OfficeXmlNamespaces.Xmlns);
        declaration.Value = uri;
        BindDeclaration(ref declaration, prefix, uri);
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
        if (_attrCount + count > _attrs.Length) Grow(ref _attrs, Math.Max(_attrs.Length * 2, _attrCount + count));
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

    /// <summary>
    /// Creates a detached deep copy of an element: its attributes, namespace declarations, child elements and text. A
    /// copy is written with the prefixes in scope where it is inserted.
    /// </summary>
    internal int CloneElement(int element) {
        if (_nodes[element].Kind != OfficeXmlTreeNodeKind.Element) throw new InvalidOperationException("Only an element can be copied.");
        OfficeXmlName name = Namespaces.GetName(_nodes[element].Name);
        int copy = CreateElement(name.NamespaceId, name.LocalName);
        int first = _nodes[element].AttrFirst, count = _nodes[element].AttrCount;
        for (int index = first; index < first + count; index++) {
            OfficeXmlName attribute = Namespaces.GetName(_attrs[index].Name);
            string value = AttributeValue(ref _attrs[index]);
            SetAttribute(copy, attribute.NamespaceId, attribute.LocalName, value);
        }
        for (int child = _nodes[element].FirstChild; child != NoNode; child = _nodes[child].Next) {
            int cloned = _nodes[child].Kind == OfficeXmlTreeNodeKind.Element ? CloneElement(child) : CreateText(GetText(child));
            Append(copy, cloned);
        }
        return copy;
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
        if (reference != NoNode && _nodes[reference].Parent != parent) throw new InvalidOperationException("The reference node is not a child of the parent.");
        for (int ancestor = parent; ancestor != NoNode; ancestor = _nodes[ancestor].Parent) {
            if (ancestor == child) throw new InvalidOperationException("A node cannot be inserted into itself or its own subtree.");
        }
        Link(parent, child, reference);
        MarkChanged(parent, ContentChanged);
        MarkSubtree(child);
        // A parsed node that moves may rely on namespace declarations of its old ancestors, so every start tag in it is
        // written again, declaring what its new place binds differently.
        if (child < _sourceNodeCount) MarkMoved(child);
    }

    // Records every binding in scope where the element sits now that it does not declare itself, keeping a binding it
    // already carries: the markup was written for the scope it was parsed in. An ancestor's own declarations come
    // before the bindings it carries, as they do when it is written.
    private void CarryScope(int element, int parent) {
        _carried ??= new Dictionary<int, List<CarriedBinding>>();
        if (!_carried.TryGetValue(element, out List<CarriedBinding>? bindings)) _carried.Add(element, bindings = new List<CarriedBinding>());
        var seen = new HashSet<string>(StringComparer.Ordinal);
        CollectDeclaredPrefixes(element, seen);
        foreach (CarriedBinding binding in bindings) seen.Add(binding.Prefix);
        for (int ancestor = parent; ancestor != NoNode; ancestor = _nodes[ancestor].Parent) {
            ref Node record = ref _nodes[ancestor];
            int end = record.AttrFirst + record.AttrCount;
            for (int index = record.AttrFirst; index < end; index++) {
                ref Attr attribute = ref _attrs[index];
                if (!IsDeclaration(ref attribute, out string prefix) || !seen.Add(prefix)) continue;
                bindings.Add(new CarriedBinding(prefix, attribute.DeclaredNamespace, attribute.DeclaredUri));
            }
            if (_carried.TryGetValue(ancestor, out List<CarriedBinding>? inherited)) {
                foreach (CarriedBinding binding in inherited) {
                    if (seen.Add(binding.Prefix)) bindings.Add(binding);
                }
            }
        }
        if (seen.Add(string.Empty)) bindings.Add(new CarriedBinding(string.Empty, OfficeXmlNamespaces.None, OfficeXmlNamespaces.None));
    }

    // A declaration on an element whose own name or attributes use the prefix would need two bindings on one tag.
    private void CheckDeclaration(int element, string prefix, string uri) {
        if (prefix.Length > 0 && uri.Length == 0) throw new ArgumentException("A prefix cannot be bound to an empty namespace.", nameof(uri));
        int uriId = OfficeXmlNamespaces.None;
        if (uri.Length > 0) Namespaces.GetId(uri, out uriId);
        if (UsesPrefix(_nodes[element].Name)) Conflict();
        ref Node record = ref _nodes[element];
        int end = record.AttrFirst + record.AttrCount;
        // An unprefixed attribute is in no namespace, whatever the default namespace is.
        for (int index = record.AttrFirst; prefix.Length > 0 && index < end; index++) {
            if (!IsDeclaration(ref _attrs[index], out _) && UsesPrefix(_attrs[index].Name)) Conflict();
        }

        bool UsesPrefix(int nameId) {
            OfficeXmlName name = Namespaces.GetName(nameId);
            // A created element has no prefix until it is written; in no namespace it needs the default undeclared.
            if (name.Prefix == null) return prefix.Length == 0 && name.NamespaceId == OfficeXmlNamespaces.None && uriId != OfficeXmlNamespaces.None;
            return name.Prefix == prefix && name.UriId != uriId && (prefix.Length > 0 || name.NamespaceId != OfficeXmlNamespaces.None || uriId != OfficeXmlNamespaces.None);
        }

        void Conflict() => throw new InvalidOperationException("The element uses the prefix '" + prefix + "' for another namespace, so it cannot bind it to '" + uri + "'.");
    }

    // Before a declaration on the element changes, its children keep the binding the prefix had there, and their start
    // tags are written again so they can declare it.
    private void CarryBindingToChildren(int element, string prefix) {
        if (!TryFindBinding(element, prefix, out int namespaceId, out int uriId)) return;
        for (int child = _nodes[element].FirstChild; child != NoNode; child = _nodes[child].Next) {
            if (_nodes[child].Kind != OfficeXmlTreeNodeKind.Element) continue;
            if (FindAttribute(child, OfficeXmlNamespaces.Xmlns, prefix.Length == 0 ? "xmlns" : prefix) >= 0) continue;
            _carried ??= new Dictionary<int, List<CarriedBinding>>();
            if (!_carried.TryGetValue(child, out List<CarriedBinding>? bindings)) _carried.Add(child, bindings = new List<CarriedBinding>());
            if (bindings.Exists(binding => binding.Prefix == prefix)) continue;
            bindings.Add(new CarriedBinding(prefix, namespaceId, uriId));
            MarkMoved(child);
        }
    }

    // The binding of a prefix in scope at the element; the default namespace is no namespace when nothing declares it.
    private bool TryFindBinding(int element, string prefix, out int namespaceId, out int uriId) {
        for (int current = element; current != NoNode; current = _nodes[current].Parent) {
            ref Node record = ref _nodes[current];
            int end = record.AttrFirst + record.AttrCount;
            for (int index = record.AttrFirst; index < end; index++) {
                ref Attr attribute = ref _attrs[index];
                if (!IsDeclaration(ref attribute, out string declared) || declared != prefix) continue;
                namespaceId = attribute.DeclaredNamespace;
                uriId = attribute.DeclaredUri;
                return true;
            }
            // A moved element's inherited bindings may live only in what it carries.
            if (CarriedBindingsOf(current) is { } carried) {
                foreach (CarriedBinding binding in carried) {
                    if (binding.Prefix != prefix) continue;
                    namespaceId = binding.NamespaceId;
                    uriId = binding.UriId;
                    return true;
                }
            }
        }
        namespaceId = uriId = OfficeXmlNamespaces.None;
        return prefix.Length == 0;
    }

    private void CollectDeclaredPrefixes(int element, HashSet<string> prefixes) {
        ref Node record = ref _nodes[element];
        int end = record.AttrFirst + record.AttrCount;
        for (int index = record.AttrFirst; index < end; index++) {
            if (IsDeclaration(ref _attrs[index], out string prefix)) prefixes.Add(prefix);
        }
    }

    private bool IsDeclaration(ref Attr attribute, out string prefix) {
        OfficeXmlName name = Namespaces.GetName(attribute.Name);
        if (name.NamespaceId != OfficeXmlNamespaces.Xmlns) {
            prefix = string.Empty;
            return false;
        }
        prefix = name.Prefix!.Length == 0 ? string.Empty : name.LocalName;
        return true;
    }

    /// <summary>The bindings the writer keeps for an element written away from where it was parsed, or null.</summary>
    private List<CarriedBinding>? CarriedBindingsOf(int element) =>
        _carried != null && _carried.TryGetValue(element, out List<CarriedBinding>? bindings) ? bindings : null;

    // An empty default declaration undeclares the default namespace, as the parser reads it.
    private void BindDeclaration(ref Attr declaration, string prefix, string uri) {
        if (prefix.Length == 0 && uri.Length == 0) {
            declaration.DeclaredNamespace = declaration.DeclaredUri = OfficeXmlNamespaces.None;
        } else {
            declaration.DeclaredNamespace = Namespaces.GetId(uri, out declaration.DeclaredUri);
        }
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
        int parent = _nodes[node].Parent;
        if (parent == NoNode) throw new InvalidOperationException("The root cannot be removed.");
        // The scope a parsed element was written in is taken before it leaves, so later edits there do not change it.
        if (node < _sourceNodeCount && _nodes[node].Kind == OfficeXmlTreeNodeKind.Element) CarryScope(node, parent);
        ref Node record = ref _nodes[node];
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
                // An element the order doesn't rank, such as an extension or unknown element, comes after the ranked
                // ones, as the SDK places them.
                if (siblingRank > rank || siblingRank < 0) {
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
        Version++;
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
        if (_nodeCount == _nodes.Length) Grow(ref _nodes, _nodes.Length * 2);
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

    // Binds a parsed declaration to the namespace its decoded value names, as the reader bound it, so the writer
    // matches a value written with character references to the names declared through it.
    private void Declare(ref Attr declaration, OfficeXmlPullReader reader, int index) {
        if (declaration.ValueLength == 0) {
            declaration.DeclaredNamespace = declaration.DeclaredUri = OfficeXmlNamespaces.None;
        } else if (Array.IndexOf(reader.Data, (byte)'&', declaration.ValueStart, declaration.ValueLength) >= 0) {
            declaration.DeclaredNamespace = Namespaces.GetId(reader.GetAttributeValue(index), out declaration.DeclaredUri);
        } else {
            declaration.DeclaredNamespace = Namespaces.GetId(reader.Data, declaration.ValueStart, declaration.ValueLength, out declaration.DeclaredUri);
        }
    }

    // The reader returns one string instance per distinct name, so a parse can find a name's id by comparing those
    // instances, with a hash of a few characters, before the namespace table hashes and compares whole strings.
    private struct ParseNameCache {
        private const int Size = 64;
        private string?[]? _locals;
        private string?[] _prefixes;
        private int[] _namespaces;
        private int[] _uris;
        private int[] _ids;

        internal int GetNameId(OfficeXmlNamespaces.Table table, string localName, string? prefix, int namespaceId, int uriId) {
            if (_locals == null) {
                _locals = new string?[Size];
                _prefixes = new string?[Size];
                _namespaces = new int[Size];
                _uris = new int[Size];
                _ids = new int[Size];
            }
            int length = localName.Length;
            int hash = length == 0 ? 0 : (length * 31 + localName[0]) * 31 + localName[length - 1];
            int slot = (hash * 31 + namespaceId * 7 + (prefix?.Length ?? 0)) & (Size - 1);
            if (ReferenceEquals(_locals[slot], localName) && ReferenceEquals(_prefixes[slot], prefix)
                && _namespaces[slot] == namespaceId && _uris[slot] == uriId) {
                return _ids[slot];
            }
            int id = table.GetNameId(localName, prefix, namespaceId, uriId);
            _locals[slot] = localName;
            _prefixes[slot] = prefix;
            _namespaces[slot] = namespaceId;
            _uris[slot] = uriId;
            _ids[slot] = id;
            return id;
        }
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
        if (node >= _spans.Length) Grow(ref _spans, _nodes.Length);
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

    private readonly struct CarriedBinding(string prefix, int namespaceId, int uriId) {
        internal string Prefix { get; } = prefix;
        internal int NamespaceId { get; } = namespaceId;
        internal int UriId { get; } = uriId;
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
