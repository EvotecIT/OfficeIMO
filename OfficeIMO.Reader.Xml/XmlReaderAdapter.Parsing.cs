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
                string path = (parent == null ? string.Empty : parent.Path + "/") + QualifiedName(reader)
                    + "[" + sibling.ToString(CultureInfo.InvariantCulture) + "]";
                var frame = new ElementFrame(path, rows.Count);
                rows.Add(new StructuredRow(path, "element", string.Empty));
                if (reader.MoveToFirstAttribute()) {
                    do {
                        CountNode();
                        string attributePath = path + "/@" + QualifiedName(reader);
                        string value = ReadScalar(reader, buffer, options.MaxScalarLength, options.MaxScalarLength, token);
                        rows.Add(new StructuredRow(attributePath, "attribute", NormalizeText(value)));
                    } while (reader.MoveToNextAttribute());
                    reader.MoveToElement();
                }
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

    private static string QualifiedName(XmlReader reader) => reader.Prefix.Length > 0 || reader.NamespaceURI.Length == 0
        ? reader.Name : "{" + reader.NamespaceURI + "}" + reader.LocalName;

    private sealed class ElementFrame {
        internal readonly string Path;
        internal readonly int RowIndex;
        internal readonly StringBuilder Text = new StringBuilder();
        internal SiblingNameCounter Siblings;
        internal ElementFrame(string path, int rowIndex) { Path = path; RowIndex = rowIndex; }
    }
}
