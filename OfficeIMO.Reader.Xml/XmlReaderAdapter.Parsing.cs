namespace OfficeIMO.Reader.Xml;

internal static partial class XmlReaderAdapter {
    // Build source-ordered rows directly from the pull parser. Limits apply before model growth;
    // no DOM or recursive traversal is needed, and malformed inputs emit no partial rows.
    private static List<StructuredRow> ReadRows(Stream stream, XmlReadOptions options, CancellationToken token) {
        using var input = new ReaderIncrementalInputStream(stream, null, token);
        using var reader = XmlReader.Create(input, new XmlReaderSettings {
            DtdProcessing = DtdProcessing.Ignore,
            XmlResolver = null
        });
        var rows = new List<StructuredRow>();
        var stack = new Stack<ElementFrame>();
        var buffer = new char[4096];
        int nodes = 0;
        void CountNode() {
            if (++nodes > options.MaxNodes)
                throw new ReaderResourceLimitException(nameof(XmlReadOptions.MaxNodes), options.MaxNodes);
        }
        while (reader.Read()) {
            token.ThrowIfCancellationRequested();
            CountNode();
            if (reader.NodeType == XmlNodeType.Element) {
                if (stack.Count >= options.MaxDepth)
                    throw new ReaderResourceLimitException(nameof(XmlReadOptions.MaxDepth), options.MaxDepth);
                bool empty = reader.IsEmptyElement;
                XName name = XName.Get(reader.LocalName, reader.NamespaceURI);
                ElementFrame? parent = stack.Count == 0 ? null : stack.Peek();
                int sibling = parent == null ? 1 : parent.Siblings.Next(name);
                var attributes = new List<(XName Name, string Value)>();
                List<(string Prefix, string Namespace)>? declarations = null;
                if (reader.MoveToFirstAttribute()) {
                    do {
                        CountNode();
                        string value = ReadScalar(reader, buffer, options.MaxScalarLength, options.MaxScalarLength, token);
                        // LINQ to XML represents the default declaration as the unqualified name "xmlns".
                        XName attributeName = reader.Name == "xmlns" ? XName.Get("xmlns")
                            : XName.Get(reader.LocalName, reader.NamespaceURI);
                        attributes.Add((attributeName, value));
                        if (reader.Prefix == "xmlns") {
                            declarations ??= new List<(string Prefix, string Namespace)>();
                            declarations.Add((reader.LocalName, value));
                        }
                    } while (reader.MoveToNextAttribute());
                    reader.MoveToElement();
                }
                NamespaceScope? namespaces = declarations == null ? parent?.Namespaces
                    : new NamespaceScope(parent?.Namespaces, declarations);
                string path = (parent == null ? string.Empty : parent.Path + "/") + QualifiedName(reader, name, namespaces)
                    + "[" + sibling.ToString(CultureInfo.InvariantCulture) + "]";
                var frame = new ElementFrame(path, rows.Count, namespaces);
                rows.Add(new StructuredRow(path, "element", string.Empty));
                foreach (var attribute in attributes)
                    rows.Add(new StructuredRow(path + "/@" + QualifiedName(reader, attribute.Name, namespaces),
                        "attribute", NormalizeText(attribute.Value)));
                if (!empty) stack.Push(frame);
            } else if (reader.NodeType == XmlNodeType.EndElement) {
                ElementFrame frame = stack.Pop();
                rows[frame.RowIndex] = new StructuredRow(frame.Path, "element", NormalizeText(frame.Text.ToString()));
            } else if (stack.Count > 0 && (reader.NodeType == XmlNodeType.Text || reader.NodeType == XmlNodeType.CDATA
                       || reader.NodeType == XmlNodeType.Whitespace || reader.NodeType == XmlNodeType.SignificantWhitespace)) {
                ElementFrame frame = stack.Peek();
                int remaining = options.MaxScalarLength - frame.Text.Length;
                string value = ReadScalar(reader, buffer, remaining, options.MaxScalarLength, token);
                if (frame.Text.Length > 0 && value.Length > 0) {
                    if (value.Length >= remaining)
                        throw new ReaderResourceLimitException(nameof(XmlReadOptions.MaxScalarLength), options.MaxScalarLength);
                    frame.Text.Append(' ');
                }
                frame.Text.Append(value);
            }
        }
        return rows;
    }

    private static string ReadScalar(XmlReader reader, char[] buffer, int maximum, int reportedMaximum, CancellationToken token) {
        var value = new StringBuilder();
        int count;
        while ((count = reader.ReadValueChunk(buffer, 0, buffer.Length)) > 0) {
            token.ThrowIfCancellationRequested();
            if (count > maximum - value.Length)
                throw new ReaderResourceLimitException(nameof(XmlReadOptions.MaxScalarLength), reportedMaximum);
            value.Append(buffer, 0, count);
        }
        return value.ToString();
    }

    private static string QualifiedName(XmlReader reader, XName name, NamespaceScope? namespaces) {
        if (name.Namespace == XNamespace.None) return name.LocalName;
        if (name.Namespace == XNamespace.Xml) return "xml:" + name.LocalName;
        if (name.Namespace == XNamespace.Xmlns) return "xmlns:" + name.LocalName;
        // Retain the prior path convention: prefer the nearest first-declared non-default
        // prefix for the namespace, excluding declarations shadowed by the current element.
        for (NamespaceScope? scope = namespaces; scope != null; scope = scope.Parent)
            foreach (var declaration in scope.Declarations)
                if (declaration.Namespace == name.NamespaceName && reader.LookupNamespace(declaration.Prefix) == name.NamespaceName)
                    return declaration.Prefix + ":" + name.LocalName;
        return "{" + name.NamespaceName + "}" + name.LocalName;
    }

    private sealed class NamespaceScope {
        internal readonly NamespaceScope? Parent;
        internal readonly List<(string Prefix, string Namespace)> Declarations;
        internal NamespaceScope(NamespaceScope? parent, List<(string Prefix, string Namespace)> declarations) {
            Parent = parent;
            Declarations = declarations;
        }
    }

    private sealed class ElementFrame {
        internal readonly string Path;
        internal readonly int RowIndex;
        internal readonly StringBuilder Text = new StringBuilder();
        internal readonly NamespaceScope? Namespaces;
        internal SiblingNameCounter Siblings;
        internal ElementFrame(string path, int rowIndex, NamespaceScope? namespaces) {
            Path = path;
            RowIndex = rowIndex;
            Namespaces = namespaces;
        }
    }
}
