using System.Collections;
using System.Globalization;
using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.Word.LegacyDoc.Write {
    /// <summary>
    /// An element of a package part the native DOC writer reads, or none. Reads on none give none, null or nothing, so
    /// optional properties chain the way the writer asks for them (<c>paragraph.Child("pPr").Child("jc").Val</c>).
    /// Names without a namespace argument are in the WordprocessingML (<c>w</c>) namespace.
    /// </summary>
    internal readonly struct LegacyDocElement : IEquatable<LegacyDocElement> {
        internal const int W = OfficeXmlNamespaces.WordprocessingML;

        private readonly OfficeXmlTree? _tree;
        private readonly int _node;

        internal LegacyDocElement(OfficeXmlTree tree, int node) {
            _tree = node == OfficeXmlTree.NoNode ? null : tree;
            _node = node;
        }

        internal static LegacyDocElement None => default;

        internal OfficeXmlTree Tree => _tree!;

        internal int Node => _node;

        internal bool Exists => _tree != null;

        internal bool IsNone => _tree == null;

        internal string LocalName => _tree == null ? string.Empty : _tree.LocalNameOf(_node);

        internal int Namespace => _tree == null ? OfficeXmlNamespaces.None : _tree.NamespaceOf(_node);

        /// <summary>The element's namespace URI, in its transitional spelling for an ECMA-376 namespace.</summary>
        internal string NamespaceUri => _tree == null ? string.Empty : _tree.Namespaces.GetUri(_tree.NamespaceOf(_node));

        /// <summary>Whether the element has this namespace URI and local name.</summary>
        internal bool Is(string namespaceUri, string localName) =>
            _tree != null && string.Equals(_tree.LocalNameOf(_node), localName, StringComparison.Ordinal)
            && string.Equals(NamespaceUri, namespaceUri, StringComparison.Ordinal);

        /// <summary>An attribute in the namespace with this URI, or null.</summary>
        internal string? AttrNs(string namespaceUri, string localName) {
            if (_tree == null) return null;
            foreach (LegacyDocAttribute attribute in Attributes) {
                if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal)
                    && string.Equals(_tree.Namespaces.GetUri(attribute.Namespace), namespaceUri, StringComparison.Ordinal)) {
                    return attribute.Value;
                }
            }
            return null;
        }

        /// <summary>The local name when the element is in the <c>w</c> namespace, otherwise null.</summary>
        internal string? WName => _tree != null && _tree.NamespaceOf(_node) == W ? _tree.LocalNameOf(_node) : null;

        internal bool Is(string localName) => _tree != null && _tree.IsElement(_node, W, localName);

        internal bool Is(int namespaceId, string localName) => _tree != null && _tree.IsElement(_node, namespaceId, localName);

        internal LegacyDocElement Parent {
            get {
                if (_tree == null) return None;
                int parent = _tree.ParentOf(_node);
                return parent == OfficeXmlTree.NoNode || _tree.KindOf(parent) != OfficeXmlTreeNodeKind.Element ? None : new LegacyDocElement(_tree, parent);
            }
        }

        /// <summary>The first child element with this <c>w</c> name, or none.</summary>
        internal LegacyDocElement Child(string localName) => _tree == null ? None : new LegacyDocElement(_tree, _tree.FindChild(_node, W, localName));

        internal LegacyDocElement Child(int namespaceId, string localName) =>
            _tree == null ? None : new LegacyDocElement(_tree, _tree.FindChild(_node, namespaceId, localName));

        internal bool HasChild(string localName) => _tree != null && _tree.FindChild(_node, W, localName) != OfficeXmlTree.NoNode;

        internal ElementList Elements() => new(_tree, _node, OfficeXmlNamespaces.None, null);

        internal ElementList Elements(string localName) => new(_tree, _node, W, localName);

        internal ElementList Elements(int namespaceId, string localName) => new(_tree, _node, namespaceId, localName);

        /// <summary>Whether the element has a child element.</summary>
        internal bool HasElements => FirstElement.Exists;

        internal LegacyDocElement FirstElement {
            get {
                if (_tree == null) return None;
                for (int child = _tree.FirstChildOf(_node); child != OfficeXmlTree.NoNode; child = _tree.NextSiblingOf(child)) {
                    if (_tree.KindOf(child) == OfficeXmlTreeNodeKind.Element) return new LegacyDocElement(_tree, child);
                }
                return None;
            }
        }

        internal LegacyDocElement NextElement {
            get {
                if (_tree == null) return None;
                for (int sibling = _tree.NextSiblingOf(_node); sibling != OfficeXmlTree.NoNode; sibling = _tree.NextSiblingOf(sibling)) {
                    if (_tree.KindOf(sibling) == OfficeXmlTreeNodeKind.Element) return new LegacyDocElement(_tree, sibling);
                }
                return None;
            }
        }

        internal LegacyDocElement PreviousElement {
            get {
                if (_tree == null) return None;
                for (int sibling = _tree.PreviousSiblingOf(_node); sibling != OfficeXmlTree.NoNode; sibling = _tree.PreviousSiblingOf(sibling)) {
                    if (_tree.KindOf(sibling) == OfficeXmlTreeNodeKind.Element) return new LegacyDocElement(_tree, sibling);
                }
                return None;
            }
        }

        /// <summary>A <c>w</c> attribute, or null.</summary>
        internal string? Attr(string localName) => _tree?.GetAttribute(_node, W, localName);

        internal string? Attr(int namespaceId, string localName) => _tree?.GetAttribute(_node, namespaceId, localName);

        internal bool HasAttr(string localName) => _tree != null && _tree.HasAttribute(_node, W, localName);

        internal bool HasAttr(int namespaceId, string localName) => _tree != null && _tree.HasAttribute(_node, namespaceId, localName);

        /// <summary>The <c>w:val</c> attribute, or null.</summary>
        internal string? Val => _tree?.GetAttribute(_node, W, "val");

        /// <summary>
        /// The element's attributes as the SDK lists them: namespace declarations and markup-compatibility attributes
        /// are left out.
        /// </summary>
        internal AttributeList Attributes => new(_tree, _node);

        internal bool HasAttributes {
            get {
                foreach (LegacyDocAttribute _ in Attributes) return true;
                return false;
            }
        }

        /// <summary>The descendant elements in document order.</summary>
        internal DescendantList Descendants() => new(_tree, _node, OfficeXmlNamespaces.None, null);

        /// <summary>The descendant elements with this <c>w</c> name in document order.</summary>
        internal DescendantList Descendants(string localName) => new(_tree, _node, W, localName);

        internal DescendantList Descendants(int namespaceId, string localName) => new(_tree, _node, namespaceId, localName);

        internal bool HasDescendant(string localName) => Descendants(localName).FirstOrNone.Exists;

        internal bool HasDescendant(int namespaceId, string localName) => Descendants(namespaceId, localName).FirstOrNone.Exists;

        /// <summary>The ancestor elements, nearest first.</summary>
        internal IEnumerable<LegacyDocElement> Ancestors() {
            for (LegacyDocElement parent = Parent; parent.Exists; parent = parent.Parent) yield return parent;
        }

        /// <summary>The text of a text element such as <c>w:t</c>: its text children joined.</summary>
        internal string Text => _tree == null ? string.Empty : _tree.GetInnerText(_node);

        /// <summary>
        /// The element's text as the SDK's <c>InnerText</c> reads it: a leaf element's text, or its child elements'
        /// texts joined. Whitespace between child elements isn't text.
        /// </summary>
        internal string InnerText {
            get {
                if (_tree == null) return string.Empty;
                if (!HasElements) return _tree.GetInnerText(_node);
                var builder = new System.Text.StringBuilder();
                AppendInnerText(_tree, _node, builder);
                return builder.ToString();
            }
        }

        private static void AppendInnerText(OfficeXmlTree tree, int element, System.Text.StringBuilder builder) {
            bool hasElements = false;
            for (int child = tree.FirstChildOf(element); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
                if (tree.KindOf(child) == OfficeXmlTreeNodeKind.Element) {
                    hasElements = true;
                    break;
                }
            }
            if (!hasElements) {
                builder.Append(tree.GetInnerText(element));
                return;
            }
            for (int child = tree.FirstChildOf(element); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
                if (tree.KindOf(child) == OfficeXmlTreeNodeKind.Element) AppendInnerText(tree, child, builder);
            }
        }

        public bool Equals(LegacyDocElement other) => ReferenceEquals(_tree, other._tree) && (_tree == null || _node == other._node);

        public override bool Equals(object? obj) => obj is LegacyDocElement other && Equals(other);

        public override int GetHashCode() => _tree == null ? 0 : (System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_tree) * 397) ^ _node;

        public static bool operator ==(LegacyDocElement left, LegacyDocElement right) => left.Equals(right);

        public static bool operator !=(LegacyDocElement left, LegacyDocElement right) => !left.Equals(right);

        public override string ToString() => _tree == null ? "(none)" : LocalName;

        /// <summary>Child elements, optionally only those with one name.</summary>
        internal readonly struct ElementList : IEnumerable<LegacyDocElement> {
            private readonly OfficeXmlTree? _tree;
            private readonly int _parent;
            private readonly int _namespaceId;
            private readonly string? _localName;

            internal ElementList(OfficeXmlTree? tree, int parent, int namespaceId, string? localName) {
                _tree = tree;
                _parent = parent;
                _namespaceId = namespaceId;
                _localName = localName;
            }

            internal bool Any() => GetEnumerator().MoveNext();

            internal LegacyDocElement FirstOrNone {
                get {
                    Enumerator enumerator = GetEnumerator();
                    return enumerator.MoveNext() ? enumerator.Current : None;
                }
            }

            internal int Count {
                get {
                    int count = 0;
                    foreach (LegacyDocElement _ in this) count++;
                    return count;
                }
            }

            internal LegacyDocElement[] ToArray() {
                var list = new List<LegacyDocElement>();
                foreach (LegacyDocElement element in this) list.Add(element);
                return list.ToArray();
            }

            internal List<LegacyDocElement> ToList() {
                var list = new List<LegacyDocElement>();
                foreach (LegacyDocElement element in this) list.Add(element);
                return list;
            }

            public Enumerator GetEnumerator() => new(_tree, _parent, _namespaceId, _localName);

            IEnumerator<LegacyDocElement> IEnumerable<LegacyDocElement>.GetEnumerator() => GetEnumerator();

            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

            internal struct Enumerator : IEnumerator<LegacyDocElement> {
                private readonly OfficeXmlTree? _tree;
                private readonly int _parent;
                private readonly int _namespaceId;
                private readonly string? _localName;
                private int _current;
                private bool _started;

                internal Enumerator(OfficeXmlTree? tree, int parent, int namespaceId, string? localName) {
                    _tree = tree;
                    _parent = parent;
                    _namespaceId = namespaceId;
                    _localName = localName;
                    _current = OfficeXmlTree.NoNode;
                    _started = false;
                }

                public LegacyDocElement Current => new(_tree!, _current);

                object IEnumerator.Current => Current;

                public bool MoveNext() {
                    if (_tree == null) return false;
                    int node = _started ? (_current == OfficeXmlTree.NoNode ? OfficeXmlTree.NoNode : _tree.NextSiblingOf(_current)) : _tree.FirstChildOf(_parent);
                    _started = true;
                    for (; node != OfficeXmlTree.NoNode; node = _tree.NextSiblingOf(node)) {
                        if (_localName == null ? _tree.KindOf(node) == OfficeXmlTreeNodeKind.Element : _tree.IsElement(node, _namespaceId, _localName)) {
                            _current = node;
                            return true;
                        }
                    }
                    _current = OfficeXmlTree.NoNode;
                    return false;
                }

                public void Reset() {
                    _current = OfficeXmlTree.NoNode;
                    _started = false;
                }

                public void Dispose() {
                }
            }
        }

        /// <summary>Descendant elements in document order, optionally only those with one name.</summary>
        internal readonly struct DescendantList : IEnumerable<LegacyDocElement> {
            private readonly OfficeXmlTree? _tree;
            private readonly int _root;
            private readonly int _namespaceId;
            private readonly string? _localName;

            internal DescendantList(OfficeXmlTree? tree, int root, int namespaceId, string? localName) {
                _tree = tree;
                _root = root;
                _namespaceId = namespaceId;
                _localName = localName;
            }

            internal bool Any() => GetEnumerator().MoveNext();

            internal LegacyDocElement FirstOrNone {
                get {
                    Enumerator enumerator = GetEnumerator();
                    return enumerator.MoveNext() ? enumerator.Current : None;
                }
            }

            internal List<LegacyDocElement> ToList() {
                var list = new List<LegacyDocElement>();
                foreach (LegacyDocElement element in this) list.Add(element);
                return list;
            }

            public Enumerator GetEnumerator() => new(_tree, _root, _namespaceId, _localName);

            IEnumerator<LegacyDocElement> IEnumerable<LegacyDocElement>.GetEnumerator() => GetEnumerator();

            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

            internal struct Enumerator : IEnumerator<LegacyDocElement> {
                private readonly OfficeXmlTree? _tree;
                private readonly int _root;
                private readonly int _namespaceId;
                private readonly string? _localName;
                private int _current;

                internal Enumerator(OfficeXmlTree? tree, int root, int namespaceId, string? localName) {
                    _tree = tree;
                    _root = root;
                    _namespaceId = namespaceId;
                    _localName = localName;
                    _current = root;
                }

                public LegacyDocElement Current => new(_tree!, _current);

                object IEnumerator.Current => Current;

                public bool MoveNext() {
                    if (_tree == null || _current == OfficeXmlTree.NoNode) return false;
                    while (true) {
                        int next = Advance(_current);
                        _current = next;
                        if (next == OfficeXmlTree.NoNode) return false;
                        if (_localName == null || _tree.IsElement(next, _namespaceId, _localName)) return true;
                    }
                }

                // The next element in document order below the root, descending into the current element first.
                private int Advance(int node) {
                    OfficeXmlTree tree = _tree!;
                    int child = FirstElementChild(tree, node);
                    if (child != OfficeXmlTree.NoNode) return child;
                    while (node != _root) {
                        for (int sibling = tree.NextSiblingOf(node); sibling != OfficeXmlTree.NoNode; sibling = tree.NextSiblingOf(sibling)) {
                            if (tree.KindOf(sibling) == OfficeXmlTreeNodeKind.Element) return sibling;
                        }
                        node = tree.ParentOf(node);
                    }
                    return OfficeXmlTree.NoNode;
                }

                private static int FirstElementChild(OfficeXmlTree tree, int node) {
                    for (int child = tree.FirstChildOf(node); child != OfficeXmlTree.NoNode; child = tree.NextSiblingOf(child)) {
                        if (tree.KindOf(child) == OfficeXmlTreeNodeKind.Element) return child;
                    }
                    return OfficeXmlTree.NoNode;
                }

                public void Reset() => _current = _root;

                public void Dispose() {
                }
            }
        }

        /// <summary>An element's attributes without namespace declarations and markup-compatibility attributes.</summary>
        internal readonly struct AttributeList : IEnumerable<LegacyDocAttribute> {
            private readonly OfficeXmlTree? _tree;
            private readonly int _node;

            internal AttributeList(OfficeXmlTree? tree, int node) {
                _tree = tree;
                _node = node;
            }

            public IEnumerator<LegacyDocAttribute> GetEnumerator() {
                if (_tree == null) yield break;
                int count = _tree.AttributeCountOf(_node);
                for (int index = 0; index < count; index++) {
                    OfficeXmlName name = _tree.AttributeNameAt(_node, index);
                    if (name.NamespaceId is OfficeXmlNamespaces.Xmlns or OfficeXmlNamespaces.MarkupCompatibility) continue;
                    yield return new LegacyDocAttribute(name.NamespaceId, name.LocalName, _tree.AttributeValueAt(_node, index));
                }
            }

            IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        }
    }

    internal readonly struct LegacyDocAttribute {
        internal LegacyDocAttribute(int namespaceId, string localName, string value) {
            Namespace = namespaceId;
            LocalName = localName;
            Value = value;
        }

        internal int Namespace { get; }

        internal string LocalName { get; }

        internal string Value { get; }
    }

    /// <summary>Reads WordprocessingML simple values the way the Open XML SDK parsed them for the writer.</summary>
    internal static class LegacyDocValues {
        /// <summary>An on/off value: true, 1 and on are on; false, 0 and off are off.</summary>
        internal static bool ParseOnOff(string value) => value switch {
            "true" or "1" or "on" => true,
            "false" or "0" or "off" => false,
            _ => throw new FormatException("'" + value + "' is not a valid on/off value.")
        };

        /// <summary>An on/off property element: on when present without a value.</summary>
        internal static bool IsOn(LegacyDocElement element) => element.Val is not { } value || ParseOnOff(value);

        /// <summary>An optional on/off attribute: null when absent.</summary>
        internal static bool? OnOff(string? value) => value == null ? null : ParseOnOff(value);

        internal static int? Int32(string? value) => value == null ? null : XmlInt32(value);

        internal static long? Int64(string? value) => value == null ? null : long.Parse(value.Trim(), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);

        internal static uint? UInt32(string? value) => value == null ? null : uint.Parse(value.Trim(), NumberStyles.None, CultureInfo.InvariantCulture);

        internal static int XmlInt32(string value) => int.Parse(value.Trim(), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);
    }
}
