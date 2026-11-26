namespace OfficeIMO.Core.Internal.OpenXml;

internal sealed partial class OfficeXmlTree {
    /// <summary>
    /// Creates a detached deep copy of an element of another tree in this one: its attributes, namespace declarations,
    /// child elements and text. Comments and processing instructions are left out. The copy is written with the
    /// prefixes in scope where it is inserted.
    /// </summary>
    internal int ImportElement(OfficeXmlTree source, int element) {
        ArgumentNullException.ThrowIfNull(source);
        if (ReferenceEquals(source, this)) return CloneElement(element);
        if (source.KindOf(element) != OfficeXmlTreeNodeKind.Element) throw new InvalidOperationException("Only an element can be copied.");
        bool sameTable = ReferenceEquals(source.Namespaces, Namespaces);
        OfficeXmlName name = source.Namespaces.GetName(source._nodes[element].Name);
        int copy = CreateElement(sameTable ? name.NamespaceId : MapNamespace(source, name.NamespaceId), name.LocalName);
        int first = source._nodes[element].AttrFirst, count = source._nodes[element].AttrCount;
        for (int index = first; index < first + count; index++) {
            OfficeXmlName attribute = source.Namespaces.GetName(source._attrs[index].Name);
            string value = source.AttributeValue(ref source._attrs[index]);
            SetAttribute(copy, sameTable ? attribute.NamespaceId : MapNamespace(source, attribute.NamespaceId), attribute.LocalName, value);
        }
        for (int child = source._nodes[element].FirstChild; child != NoNode; child = source._nodes[child].Next) {
            switch (source._nodes[child].Kind) {
                case OfficeXmlTreeNodeKind.Element:
                    Append(copy, ImportElement(source, child));
                    break;
                case OfficeXmlTreeNodeKind.Text:
                    Append(copy, CreateText(source.GetText(child)));
                    break;
            }
        }
        return copy;
    }

    /// <summary>
    /// Creates a detached element in a namespace outside the fixed ones, written with <paramref name="prefix"/> unless
    /// that prefix is bound to another namespace where it is inserted.
    /// </summary>
    internal int CreateElement(string namespaceUri, string localName, string prefix) {
        int namespaceId = Namespaces.GetId(namespaceUri, out int uriId);
        int node = NewNode(OfficeXmlTreeNodeKind.Element);
        ref Node record = ref _nodes[node];
        record.Name = Namespaces.GetNameId(localName, prefix, namespaceId, uriId);
        record.Flags = StartTagChanged | ContentChanged | SubtreeChanged;
        return node;
    }

    /// <summary>
    /// Creates a detached deep copy of an element that keeps the prefixes it was written with, so attribute values that
    /// name prefixes (such as <c>mc:Ignorable</c>) stay valid; a prefix the place it is inserted does not bind is declared.
    /// </summary>
    internal int CloneElementKeepingPrefixes(int element) =>
        ImportOrClone(this, element, Namespaces.GetName(_nodes[element].Name));

    /// <summary>A new tree whose root is a copy of <paramref name="element"/>, for writing the element on its own.</summary>
    internal static OfficeXmlTree CopyElement(OfficeXmlTree source, int element) {
        ArgumentNullException.ThrowIfNull(source);
        OfficeXmlName name = source.Namespaces.GetName(source._nodes[element].Name);
        var tree = new OfficeXmlTree(source.Namespaces, 32, 32);
        tree.Root = tree.ImportOrClone(source, element, name);
        return tree;
    }

    private int ImportOrClone(OfficeXmlTree source, int element, OfficeXmlName name) {
        // Both trees share the table, so the copy keeps every name id, prefixes included; a prefix the scope where the
        // copy is written does not bind is declared there.
        int copy = CreateElement(name.NamespaceId, name.LocalName);
        _nodes[copy].Name = source._nodes[element].Name;
        int first = source._nodes[element].AttrFirst, count = source._nodes[element].AttrCount;
        for (int index = first; index < first + count; index++) {
            OfficeXmlName attribute = source.Namespaces.GetName(source._attrs[index].Name);
            SetAttribute(copy, attribute.NamespaceId, attribute.LocalName, source.AttributeValue(ref source._attrs[index]));
        }
        for (int child = source._nodes[element].FirstChild; child != NoNode; child = source._nodes[child].Next) {
            switch (source._nodes[child].Kind) {
                case OfficeXmlTreeNodeKind.Element:
                    Append(copy, ImportOrClone(source, child, source.Namespaces.GetName(source._nodes[child].Name)));
                    break;
                case OfficeXmlTreeNodeKind.Text:
                    Append(copy, CreateText(source.GetText(child)));
                    break;
            }
        }
        return copy;
    }

    /// <summary>
    /// The element and its subtree as UTF-8 markup without an XML declaration, every prefix it uses declared on it, so
    /// another XML layer can parse it on its own.
    /// </summary>
    internal byte[] WriteElement(int element) {
        OfficeXmlTree tree = element == Root && _source == null ? this : CopyElement(this, element);
        using var stream = new MemoryStream();
        var writer = new OfficeXmlUtf8Writer(4096) { Target = stream };
        writer.ResetScope();
        tree.WriteNode(writer, tree.Root);
        writer.Flush();
        return stream.ToArray();
    }

    /// <summary>The text of every text node under the element, in document order, as XML's string value reads it.</summary>
    internal string GetDescendantText(int element) {
        int first = _nodes[element].FirstChild;
        if (first == NoNode) return string.Empty;
        if (_nodes[first].Next == NoNode && _nodes[first].Kind == OfficeXmlTreeNodeKind.Text) return GetText(first);
        var builder = new System.Text.StringBuilder();
        AppendDescendantText(element, builder);
        return builder.ToString();
    }

    private void AppendDescendantText(int element, System.Text.StringBuilder builder) {
        for (int child = _nodes[element].FirstChild; child != NoNode; child = _nodes[child].Next) {
            switch (_nodes[child].Kind) {
                case OfficeXmlTreeNodeKind.Text: builder.Append(GetText(child)); break;
                case OfficeXmlTreeNodeKind.Element: AppendDescendantText(child, builder); break;
            }
        }
    }

    /// <summary>Whether an attribute of the element in the relationships namespace has the value <paramref name="id"/>.</summary>
    internal bool HasRelationshipAttribute(int element, string id) {
        int first = _nodes[element].AttrFirst, count = _nodes[element].AttrCount;
        for (int index = first; index < first + count; index++) {
            if (Namespaces.GetName(_attrs[index].Name).NamespaceId == OfficeXmlNamespaces.Relationships
                && string.Equals(AttributeValue(ref _attrs[index]), id, StringComparison.Ordinal)) {
                return true;
            }
        }
        return false;
    }

    // The id this tree's table gives the namespace an id of the source table stands for.
    private int MapNamespace(OfficeXmlTree source, int namespaceId) =>
        OfficeXmlNamespaces.IsEcma376(namespaceId) || namespaceId <= OfficeXmlNamespaces.Xmlns
            ? namespaceId
            : Namespaces.GetId(source.Namespaces.GetUri(namespaceId));
}
