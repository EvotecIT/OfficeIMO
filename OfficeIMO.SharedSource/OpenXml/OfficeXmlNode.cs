using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.OpenXml.Internal {
    /// <summary>
    /// An element of a part, wherever the part is held: in OfficeIMO's own tree while the document works without the
    /// Open XML SDK, or in the SDK package once the engine's <c>OpenXmlDocument</c> opened it, where edits go to the
    /// SDK's elements in place so references the caller holds stay current. Engine code reads and edits elements through
    /// this type only, so it has one path for both. The SDK side is <c>OfficeSdkXml</c>, which each engine compiles into
    /// its SDK bridge folder.
    /// </summary>
    /// <remarks>
    /// Children are elements only: text is read and written through <see cref="InnerText"/>, and comments and
    /// processing instructions are skipped. Namespaces are the fixed ids of <see cref="OfficeXmlNamespaces"/>; a namespace
    /// outside them is named by its URI.
    /// </remarks>
    internal readonly struct OfficeXmlNode : IEquatable<OfficeXmlNode> {
        private readonly OfficeXmlTree? _tree;
        private readonly int _node;
        private readonly object? _sdk;

        internal OfficeXmlNode(OfficeXmlTree tree, int node) {
            _tree = node == OfficeXmlTree.NoNode ? null : tree;
            _node = node;
            _sdk = null;
        }

        private OfficeXmlNode(object element) {
            _tree = null;
            _node = OfficeXmlTree.NoNode;
            _sdk = element;
        }

        /// <summary>An element of the SDK package, passed as <see cref="object"/> so engine code names no SDK type.</summary>
        internal static OfficeXmlNode FromSdk(object? element) => element == null ? default : new OfficeXmlNode(element);

        /// <summary>No element: what a lookup returns when nothing matches.</summary>
        internal static OfficeXmlNode None => default;

        internal bool IsNone => _tree == null && _sdk == null;
        internal bool Exists => !IsNone;

        /// <summary>The tree that holds the element, or null when it is an SDK element.</summary>
        internal OfficeXmlTree? Tree => _tree;

        /// <summary>The element's node in <see cref="Tree"/>.</summary>
        internal int Node => _tree == null ? OfficeXmlTree.NoNode : _node;

        /// <summary>The SDK element, or null when the element is in a tree.</summary>
        internal object? SdkElement => _sdk;

        internal bool IsSdk => _sdk != null;

        // ---- Name -----------------------------------------------------------------------------------------------

        internal string LocalName => _tree != null ? _tree.LocalNameOf(_node) : _sdk != null ? OfficeSdkXml.LocalName(_sdk) : string.Empty;

        /// <summary>The fixed namespace id, or <see cref="OfficeXmlNamespaces.None"/> for a namespace outside them.</summary>
        internal int NamespaceId {
            get {
                if (_tree != null) {
                    int id = _tree.NamespaceOf(_node);
                    return OfficeXmlNamespaces.IsEcma376(id) ? id : OfficeXmlNamespaces.None;
                }
                return _sdk != null ? OfficeSdkXml.NamespaceId(_sdk) : OfficeXmlNamespaces.None;
            }
        }

        internal string NamespaceUri => _tree != null ? _tree.Namespaces.GetUri(_tree.NamespaceOf(_node)) : _sdk != null ? OfficeSdkXml.NamespaceUri(_sdk) : string.Empty;

        internal bool Is(int namespaceId, string localName) =>
            _tree != null ? _tree.IsElement(_node, namespaceId, localName)
            : _sdk != null && OfficeSdkXml.Is(_sdk, namespaceId, localName);

        internal bool Is(string namespaceUri, string localName) =>
            !IsNone && string.Equals(LocalName, localName, StringComparison.Ordinal) && string.Equals(NamespaceUri, namespaceUri, StringComparison.Ordinal);

        // ---- Navigation -----------------------------------------------------------------------------------------

        internal OfficeXmlNode Parent {
            get {
                if (_tree != null) return new OfficeXmlNode(_tree, _tree.ParentOf(_node));
                return _sdk != null ? FromSdk(OfficeSdkXml.Parent(_sdk)) : default;
            }
        }

        internal OfficeXmlNode FirstChild {
            get {
                if (_tree != null) return new OfficeXmlNode(_tree, SkipToElement(_tree, _tree.FirstChildOf(_node), forward: true));
                return _sdk != null ? FromSdk(OfficeSdkXml.FirstChild(_sdk)) : default;
            }
        }

        internal OfficeXmlNode LastChild {
            get {
                if (_tree != null) return new OfficeXmlNode(_tree, SkipToElement(_tree, _tree.LastChildOf(_node), forward: false));
                return _sdk != null ? FromSdk(OfficeSdkXml.LastChild(_sdk)) : default;
            }
        }

        internal OfficeXmlNode NextSibling {
            get {
                if (_tree != null) return new OfficeXmlNode(_tree, SkipToElement(_tree, _tree.NextSiblingOf(_node), forward: true));
                return _sdk != null ? FromSdk(OfficeSdkXml.NextSibling(_sdk)) : default;
            }
        }

        internal OfficeXmlNode PreviousSibling {
            get {
                if (_tree != null) return new OfficeXmlNode(_tree, SkipToElement(_tree, _tree.PreviousSiblingOf(_node), forward: false));
                return _sdk != null ? FromSdk(OfficeSdkXml.PreviousSibling(_sdk)) : default;
            }
        }

        private static int SkipToElement(OfficeXmlTree tree, int node, bool forward) {
            while (node != OfficeXmlTree.NoNode && tree.KindOf(node) != OfficeXmlTreeNodeKind.Element) {
                node = forward ? tree.NextSiblingOf(node) : tree.PreviousSiblingOf(node);
            }
            return node;
        }

        /// <summary>The first child element with this name, or <see cref="None"/>.</summary>
        internal OfficeXmlNode Child(int namespaceId, string localName) {
            if (_tree != null) return new OfficeXmlNode(_tree, _tree.FindChild(_node, namespaceId, localName));
            if (_sdk == null) return default;
            for (OfficeXmlNode child = FirstChild; !child.IsNone; child = child.NextSibling) {
                if (child.Is(namespaceId, localName)) return child;
            }
            return default;
        }

        internal OfficeXmlNode Child(string namespaceUri, string localName) {
            for (OfficeXmlNode child = FirstChild; !child.IsNone; child = child.NextSibling) {
                if (child.Is(namespaceUri, localName)) return child;
            }
            return default;
        }

        /// <summary>The next sibling element with this name, or <see cref="None"/>.</summary>
        internal OfficeXmlNode NextSiblingNamed(int namespaceId, string localName) {
            if (_tree != null) return new OfficeXmlNode(_tree, _tree.FindNextSibling(_node, namespaceId, localName));
            for (OfficeXmlNode sibling = NextSibling; !sibling.IsNone; sibling = sibling.NextSibling) {
                if (sibling.Is(namespaceId, localName)) return sibling;
            }
            return default;
        }

        /// <summary>The child elements, in document order.</summary>
        internal ChildList Children() => new(this, OfficeXmlNamespaces.None, null);

        /// <summary>The child elements with this name, in document order.</summary>
        internal ChildList Children(int namespaceId, string localName) => new(this, namespaceId, localName);

        /// <summary>Whether this element is <paramref name="ancestor"/> or inside it.</summary>
        internal bool IsInside(OfficeXmlNode ancestor) {
            for (OfficeXmlNode node = this; !node.IsNone; node = node.Parent) {
                if (node.Equals(ancestor)) return true;
            }
            return false;
        }

        /// <summary>The nearest ancestor with this name, this element included, or <see cref="None"/>.</summary>
        internal OfficeXmlNode AncestorOrSelf(int namespaceId, string localName) {
            for (OfficeXmlNode node = this; !node.IsNone; node = node.Parent) {
                if (node.Is(namespaceId, localName)) return node;
            }
            return default;
        }

        /// <summary>Whether the element is attached up to a part's root element.</summary>
        internal bool IsAttached {
            get {
                if (_tree != null) {
                    int node = _node;
                    while (_tree.ParentOf(node) != OfficeXmlTree.NoNode) node = _tree.ParentOf(node);
                    return node == _tree.Root;
                }
                return _sdk != null && OfficeSdkXml.IsAttached(_sdk);
            }
        }

        // ---- Attributes -----------------------------------------------------------------------------------------

        internal string? Attribute(string localName) => Attribute(OfficeXmlNamespaces.None, localName);

        internal string? Attribute(int namespaceId, string localName) =>
            _tree != null ? _tree.GetAttribute(_node, namespaceId, localName)
            : _sdk != null ? OfficeSdkXml.GetAttribute(_sdk, namespaceId, localName) : null;

        internal bool HasAttribute(string localName) => Attribute(OfficeXmlNamespaces.None, localName) != null;

        internal bool HasAttribute(int namespaceId, string localName) => Attribute(namespaceId, localName) != null;

        /// <summary>Sets an attribute; null removes it.</summary>
        internal void SetAttribute(string localName, string? value) => SetAttribute(OfficeXmlNamespaces.None, localName, value);

        /// <summary>Sets an attribute; null removes it.</summary>
        internal void SetAttribute(int namespaceId, string localName, string? value) {
            if (_tree != null) {
                if (value == null) _tree.RemoveAttribute(_node, namespaceId, localName);
                else _tree.SetAttribute(_node, namespaceId, localName, value);
            } else if (_sdk != null) {
                OfficeSdkXml.SetAttribute(_sdk, namespaceId, localName, value);
            } else {
                throw new InvalidOperationException("There is no element to set the attribute on.");
            }
        }

        internal bool RemoveAttribute(string localName) => RemoveAttribute(OfficeXmlNamespaces.None, localName);

        internal bool RemoveAttribute(int namespaceId, string localName) {
            if (Attribute(namespaceId, localName) == null) return false;
            SetAttribute(namespaceId, localName, null);
            return true;
        }

        internal bool? BooleanAttribute(string localName) => OfficeXmlViewValues.ParseBoolean(Attribute(localName));
        internal int? Int32Attribute(string localName) => OfficeXmlViewValues.ParseInt32(Attribute(localName));
        internal uint? UInt32Attribute(string localName) => OfficeXmlViewValues.ParseUInt32(Attribute(localName));
        internal long? Int64Attribute(string localName) => OfficeXmlViewValues.ParseInt64(Attribute(localName));
        internal double? DoubleAttribute(string localName) => OfficeXmlViewValues.ParseDouble(Attribute(localName));

        internal void SetBooleanAttribute(string localName, bool? value) => SetAttribute(localName, OfficeXmlViewValues.Format(value));
        internal void SetInt32Attribute(string localName, int? value) => SetAttribute(localName, OfficeXmlViewValues.Format(value));
        internal void SetUInt32Attribute(string localName, uint? value) => SetAttribute(localName, OfficeXmlViewValues.Format(value));
        internal void SetInt64Attribute(string localName, long? value) => SetAttribute(localName, OfficeXmlViewValues.Format(value));
        internal void SetDoubleAttribute(string localName, double? value) => SetAttribute(localName, OfficeXmlViewValues.Format(value));

        // ---- Text -----------------------------------------------------------------------------------------------

        /// <summary>The text of the element and its descendants.</summary>
        internal string InnerText =>
            _tree != null ? _tree.GetDescendantText(_node) : _sdk != null ? OfficeSdkXml.InnerText(_sdk) : string.Empty;

        /// <summary>Replaces the element's content with text, as for <c>a:t</c>.</summary>
        internal void SetInnerText(string text) {
            ArgumentNullException.ThrowIfNull(text);
            if (_tree != null) _tree.SetInnerText(_node, text);
            else if (_sdk != null) OfficeSdkXml.SetInnerText(_sdk, text);
            else throw new InvalidOperationException("There is no element to set the text of.");
        }

        /// <summary>The text of a child element, or null when there is no such child.</summary>
        internal string? ChildText(int namespaceId, string localName) {
            OfficeXmlNode child = Child(namespaceId, localName);
            return child.IsNone ? null : child.InnerText;
        }

        // ---- Structure ------------------------------------------------------------------------------------------

        /// <summary>
        /// Creates a detached element to fill and then insert under an element of this part. For an SDK element it is the
        /// SDK's typed class this element expects as its child, read again if it goes under another parent; one the SDK
        /// has no class for there is built in a small tree of its own and becomes an SDK element when it is inserted.
        /// </summary>
        internal OfficeXmlNode CreateElement(int namespaceId, string localName) {
            if (_tree != null) return new OfficeXmlNode(_tree, _tree.CreateElement(namespaceId, localName));
            if (_sdk != null && OfficeXmlNamespaces.KnownUri(namespaceId) != null
                && OfficeSdkXml.CreateElement(_sdk, namespaceId, localName) is { } typed) return FromSdk(typed);
            OfficeXmlTree scratch = OfficeXmlTree.Create(new OfficeXmlNamespaces.Table(), namespaceId, localName, 8);
            return new OfficeXmlNode(scratch, scratch.Root);
        }

        /// <summary>Creates a detached element in a namespace outside the fixed ones, written with <paramref name="prefix"/>.</summary>
        internal OfficeXmlNode CreateElement(string namespaceUri, string localName, string prefix) {
            if (_tree != null) return new OfficeXmlNode(_tree, _tree.CreateElement(namespaceUri, localName, prefix));
            var table = new OfficeXmlNamespaces.Table();
            OfficeXmlTree scratch = OfficeXmlTree.Create(table, OfficeXmlNamespaces.None, "scratch", 8);
            int node = scratch.CreateElement(namespaceUri, localName, prefix);
            scratch.Append(scratch.Root, node);
            return new OfficeXmlNode(scratch, node);
        }

        /// <summary>Appends a detached element and returns it as it is now held under this element.</summary>
        internal OfficeXmlNode Append(OfficeXmlNode child) => InsertBefore(child, default);

        /// <summary>
        /// Inserts a detached element before <paramref name="reference"/>, or last when it is <see cref="None"/>, and
        /// returns it as it is now held under this element: the same node in the same tree, or a new element made from
        /// a copy of it otherwise.
        /// </summary>
        internal OfficeXmlNode InsertBefore(OfficeXmlNode child, OfficeXmlNode reference) {
            if (child.IsNone) throw new ArgumentException("There is no element to insert.", nameof(child));
            if (_tree != null) {
                int node;
                if (ReferenceEquals(child._tree, _tree)) {
                    node = child._node;
                } else if (child._tree != null) {
                    node = _tree.ImportElement(child._tree, child._node);
                } else {
                    // An SDK element into a tree: its markup is read into the tree.
                    byte[] markup = OfficeSdkXml.OuterXml(child._sdk!);
                    OfficeXmlTree parsed = OfficeXmlTree.Parse(markup, markup.Length, _tree.Namespaces);
                    node = _tree.ImportElement(parsed, parsed.Root);
                }
                _tree.InsertBefore(_node, node, reference.Node);
                return new OfficeXmlNode(_tree, node);
            }
            if (_sdk == null) throw new InvalidOperationException("There is no element to insert into.");
            return FromSdk(child._sdk != null
                ? OfficeSdkXml.InsertBefore(_sdk, child._sdk, reference._sdk)
                : OfficeSdkXml.InsertMarkupBefore(_sdk, child._tree!.WriteElement(child._node), reference._sdk));
        }

        /// <summary>Inserts a detached element right after this one and returns it as it is now held.</summary>
        internal OfficeXmlNode InsertAfterSelf(OfficeXmlNode sibling) => Parent.InsertBefore(sibling, NextSibling);

        /// <summary>Inserts a detached element among the children in schema order and returns it as it is now held.</summary>
        internal OfficeXmlNode InsertOrdered(OfficeXmlNode child, OfficeXmlChildOrder order) {
            string localName = child.LocalName;
            int namespaceId = child.NamespaceId;
            int rank = order.RankOf(namespaceId, localName);
            OfficeXmlNode reference = default;
            if (rank >= 0) {
                for (OfficeXmlNode sibling = FirstChild; !sibling.IsNone; sibling = sibling.NextSibling) {
                    // An element the order doesn't rank, such as an extension or unknown element, comes after the
                    // ranked ones, as the SDK places them.
                    int siblingRank = order.RankOf(sibling.NamespaceId, sibling.LocalName);
                    if (siblingRank > rank || siblingRank < 0) {
                        reference = sibling;
                        break;
                    }
                }
            }
            return InsertBefore(child, reference);
        }

        /// <summary>Adds a new empty child element in schema order and returns it.</summary>
        internal OfficeXmlNode AddOrdered(int namespaceId, string localName, OfficeXmlChildOrder order) {
            if (_tree != null) {
                int node = _tree.CreateElement(namespaceId, localName);
                _tree.InsertOrdered(_node, node, order);
                return new OfficeXmlNode(_tree, node);
            }
            return InsertOrdered(CreateElement(namespaceId, localName), order);
        }

        /// <summary>The child element with this name, created in schema order when it is missing.</summary>
        internal OfficeXmlNode GetOrAddChild(int namespaceId, string localName, OfficeXmlChildOrder order) {
            if (_tree != null) return new OfficeXmlNode(_tree, _tree.GetOrInsertChild(_node, namespaceId, localName, order));
            OfficeXmlNode child = Child(namespaceId, localName);
            return child.IsNone ? AddOrdered(namespaceId, localName, order) : child;
        }

        /// <summary>The child element with this name, created first among the children when it is missing.</summary>
        internal OfficeXmlNode GetOrPrependChild(int namespaceId, string localName) {
            OfficeXmlNode child = Child(namespaceId, localName);
            return child.IsNone ? InsertBefore(CreateElement(namespaceId, localName), FirstChild) : child;
        }

        /// <summary>Sets the text of a child element, creating it in schema order; null removes the child.</summary>
        internal void SetChildText(int namespaceId, string localName, string? value, OfficeXmlChildOrder order) {
            OfficeXmlNode child = Child(namespaceId, localName);
            if (value == null) {
                if (!child.IsNone) child.Remove();
                return;
            }
            if (child.IsNone) child = AddOrdered(namespaceId, localName, order);
            child.SetInnerText(value);
        }

        /// <summary>Detaches the element and its subtree.</summary>
        internal void Remove() {
            if (_tree != null) _tree.Remove(_node);
            else if (_sdk != null) OfficeSdkXml.Remove(_sdk);
        }

        /// <summary>Removes every child element with this name.</summary>
        internal void RemoveChildren(int namespaceId, string localName) {
            OfficeXmlNode child = Child(namespaceId, localName);
            while (!child.IsNone) {
                OfficeXmlNode next = child.NextSiblingNamed(namespaceId, localName);
                child.Remove();
                child = next;
            }
        }

        /// <summary>A detached deep copy of the element in the same part.</summary>
        internal OfficeXmlNode Clone() {
            if (_tree != null) return new OfficeXmlNode(_tree, _tree.CloneElement(_node));
            return _sdk != null ? FromSdk(OfficeSdkXml.Clone(_sdk)) : default;
        }

        /// <summary>The element's markup as UTF-8, every prefix it uses declared on it.</summary>
        internal byte[] OuterXml() =>
            _tree != null ? _tree.WriteElement(_node) : _sdk != null ? OfficeSdkXml.OuterXml(_sdk) : Array.Empty<byte>();

        // ---- Identity -------------------------------------------------------------------------------------------

        public bool Equals(OfficeXmlNode other) =>
            ReferenceEquals(_tree, other._tree) && ReferenceEquals(_sdk, other._sdk) && (_tree == null || _node == other._node);

        public override bool Equals(object? obj) => obj is OfficeXmlNode other && Equals(other);

        public override int GetHashCode() =>
            _tree != null ? (_tree.GetHashCode() * 397) ^ _node : _sdk != null ? System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(_sdk) : 0;

        public static bool operator ==(OfficeXmlNode left, OfficeXmlNode right) => left.Equals(right);
        public static bool operator !=(OfficeXmlNode left, OfficeXmlNode right) => !left.Equals(right);

        /// <summary>Child elements, all of them or those of one name, enumerated without allocating.</summary>
        internal readonly struct ChildList {
            private readonly OfficeXmlNode _parent;
            private readonly int _namespaceId;
            private readonly string? _localName;

            internal ChildList(OfficeXmlNode parent, int namespaceId, string? localName) {
                _parent = parent;
                _namespaceId = namespaceId;
                _localName = localName;
            }

            public Enumerator GetEnumerator() => new(_parent, _namespaceId, _localName);

            internal int Count {
                get {
                    int count = 0;
                    foreach (OfficeXmlNode _ in this) count++;
                    return count;
                }
            }

            internal OfficeXmlNode First() {
                foreach (OfficeXmlNode child in this) return child;
                return default;
            }

            internal OfficeXmlNode Last() {
                OfficeXmlNode last = default;
                foreach (OfficeXmlNode child in this) last = child;
                return last;
            }

            internal List<OfficeXmlNode> ToList() {
                var list = new List<OfficeXmlNode>();
                foreach (OfficeXmlNode child in this) list.Add(child);
                return list;
            }

            internal struct Enumerator {
                private readonly OfficeXmlNode _parent;
                private readonly int _namespaceId;
                private readonly string? _localName;
                private OfficeXmlNode _current;
                private bool _started;

                internal Enumerator(OfficeXmlNode parent, int namespaceId, string? localName) {
                    _parent = parent;
                    _namespaceId = namespaceId;
                    _localName = localName;
                    _current = default;
                    _started = false;
                }

                public OfficeXmlNode Current => _current;

                public bool MoveNext() {
                    _current = _started ? _current.NextSibling : _parent.FirstChild;
                    _started = true;
                    while (!_current.IsNone && _localName != null && !_current.Is(_namespaceId, _localName)) _current = _current.NextSibling;
                    return !_current.IsNone;
                }
            }
        }
    }
}
