using System.Xml;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Reads XML as data with the Newtonsoft-style convention: the root element is the single top-level property,
/// attributes are <c>@name</c>, text beside attributes or elements is <c>#text</c>, and repeated child elements become
/// arrays. <c>xsi:nil="true"</c> is null; DTDs are prohibited.
/// <para>
/// A quick first pass reads only the root's child names. When same-name children are adjacent (the usual
/// <c>&lt;rows&gt;&lt;row/&gt;…</c> shape) they stream one record at a time, each record buffered on its own; only a
/// root whose children interleave (<c>a, b, a</c>) is buffered whole so the repeats can be grouped.
/// </para>
/// </summary>
internal sealed class XmlEventReader : DataEventReader {
    private const string XsiNamespace = "http://www.w3.org/2001/XMLSchema-instance";
    private readonly Func<XmlReader> _open;
    private readonly bool _infer;
    private readonly Queue<Item> _queue = new();
    private DataNodeReplay? _replay;
    private XmlReader? _xml;
    private Phase _phase;
    private string? _arrayName;
    private readonly List<string> _rootText = new();

    private enum Phase {
        Start,
        RootChildren,
        Done
    }

    // One queued output: an event, or a buffered node to replay.
    private readonly struct Item {
        public Item(DataTokenType type, DataScalarKind kind = default, string? text = null) {
            Type = type;
            Kind = kind;
            Text = text;
            Node = null;
        }

        public Item(DataNode node) {
            Type = DataTokenType.None;
            Kind = default;
            Text = null;
            Node = node;
        }

        public DataTokenType Type { get; }
        public DataScalarKind Kind { get; }
        public string? Text { get; }
        public DataNode? Node { get; }
    }

    public XmlEventReader(Func<XmlReader> open, DataConversionContext context) : base(context) {
        _open = open;
        _infer = context.Options.InferTypes ?? false;
        TrackDuplicateKeys = true;
    }

    /// <summary>Reads a seekable XML stream (it is read twice).</summary>
    public static XmlEventReader FromStream(Stream input, DataConversionContext context) {
        if (!input.CanSeek) throw new ArgumentException("XML conversion needs a seekable stream.", nameof(input));
        long start = input.Position;
        XmlReaderSettings settings = Settings(context);
        return new XmlEventReader(() => {
            input.Position = start;
            return XmlReader.Create(input, settings);
        }, context);
    }

    internal static XmlReaderSettings Settings(DataConversionContext context) {
        XmlReaderSettings settings = context.Options.XmlReader?.Clone() ?? new XmlReaderSettings();
        settings.DtdProcessing = DtdProcessing.Prohibit;
        settings.XmlResolver = null;
        settings.IgnoreWhitespace = true;
        settings.CloseInput = false;
        return settings;
    }

    public override string? Location => _xml is IXmlLineInfo info && info.HasLineInfo()
        ? "line " + info.LineNumber.ToString(CultureInfo.InvariantCulture) + ", column " + info.LinePosition.ToString(CultureInfo.InvariantCulture)
        : null;

    protected override bool ReadCore() {
        try {
            while (true) {
                if (_replay != null) {
                    if (_replay.Next(out DataTokenType type, out DataScalarKind kind, out string? text)) {
                        Deliver(type, kind, text);
                        return true;
                    }

                    _replay = null;
                }

                if (_queue.Count > 0) {
                    Item item = _queue.Dequeue();
                    if (item.Node != null) {
                        _replay = new DataNodeReplay(item.Node);
                        continue;
                    }

                    Deliver(item.Type, item.Kind, item.Text);
                    return true;
                }

                if (!Advance()) return false;
            }
        } catch (XmlException error) {
            throw new InvalidDataException("The XML input is invalid: " + error.Message, error);
        }
    }

    private void Deliver(DataTokenType type, DataScalarKind kind, string? text) {
        switch (type) {
            case DataTokenType.PropertyName: SetPropertyName(text.AsSpan()); break;
            case DataTokenType.Scalar: SetScalar(kind, text.AsSpan()); break;
            case DataTokenType.Comment: SetComment(text.AsSpan()); break;
            default: Set(type); break;
        }
    }

    private bool Advance() {
        switch (_phase) {
            case Phase.Start:
                Start();
                return true;
            case Phase.RootChildren:
                RootChild();
                return true;
            default:
                return false;
        }
    }

    private void Start() {
        (int elements, bool interleaved) = ScanRoot();
        _xml = _open();
        if (!MoveToRoot()) throw new InvalidDataException("The XML input has no root element.");
        _queue.Enqueue(new Item(DataTokenType.StartDocument));
        _queue.Enqueue(new Item(DataTokenType.StartObject));
        _queue.Enqueue(new Item(DataTokenType.PropertyName, text: Name(_xml.Name)));
        if (elements == 0 || interleaved || IsNil() || _xml.IsEmptyElement) {
            _queue.Enqueue(new Item(ReadElement()));
            Finish();
            return;
        }

        List<KeyValuePair<string, DataNode>> attributes = ReadAttributes();
        _xml.Read();
        _queue.Enqueue(new Item(DataTokenType.StartObject));
        foreach (KeyValuePair<string, DataNode> attribute in attributes) {
            _queue.Enqueue(new Item(DataTokenType.PropertyName, text: attribute.Key));
            _queue.Enqueue(new Item(attribute.Value));
        }

        _phase = Phase.RootChildren;
    }

    // Reads the root's child element names only, to choose between streaming and grouping.
    private (int Elements, bool Interleaved) ScanRoot() {
        using XmlReader scan = _open();
        if (scan.MoveToContent() != XmlNodeType.Element || scan.IsEmptyElement) return (0, false);
        int depth = scan.Depth;
        var seen = new HashSet<string>(StringComparer.Ordinal);
        string? previous = null;
        int elements = 0;
        bool interleaved = false;
        scan.Read();
        while (!scan.EOF && scan.Depth > depth) {
            if (scan.NodeType == XmlNodeType.Element) {
                elements++;
                string name = scan.Name;
                if (name != previous && !seen.Add(name)) interleaved = true;
                previous = name;
                scan.Skip();
            } else {
                scan.Read();
            }
        }

        return (elements, interleaved);
    }

    // Reads up to the root element, carrying comments and reporting processing instructions on the way.
    private bool MoveToRoot() {
        while (_xml!.Read()) {
            switch (_xml.NodeType) {
                case XmlNodeType.Element:
                    return true;
                case XmlNodeType.Comment:
                    _queue.Enqueue(new Item(DataTokenType.Comment, text: _xml.Value));
                    break;
                case XmlNodeType.ProcessingInstruction:
                    ReportInstruction();
                    break;
            }
        }

        return false;
    }

    private void Finish() {
        _queue.Enqueue(new Item(DataTokenType.EndObject));
        _queue.Enqueue(new Item(DataTokenType.EndDocument));
        _phase = Phase.Done;
        while (_xml!.Read()) {
            if (_xml.NodeType == XmlNodeType.Comment) _queue.Enqueue(new Item(DataTokenType.Comment, text: _xml.Value));
            else if (_xml.NodeType == XmlNodeType.ProcessingInstruction) ReportInstruction();
        }
    }

    // Streams one child of the root. Same-name neighbors form an array, decided by peeking at the next sibling.
    private void RootChild() {
        XmlReader xml = _xml!;
        SkipNonElements(_rootText);
        if (xml.NodeType != XmlNodeType.Element) {
            if (_rootText.Count > 0) {
                ReportMixed();
                _queue.Enqueue(new Item(DataTokenType.PropertyName, text: "#text"));
                _queue.Enqueue(new Item(TextNode(_rootText)));
            }

            _queue.Enqueue(new Item(DataTokenType.EndObject));
            Finish();
            return;
        }

        string name = Name(xml.Name);
        DataNode value = ReadElement();
        SkipNonElements(_rootText);
        bool repeats = xml.NodeType == XmlNodeType.Element && Name(xml.Name) == name;
        if (_arrayName != name) {
            _queue.Enqueue(new Item(DataTokenType.PropertyName, text: name));
            if (repeats) {
                _arrayName = name;
                _queue.Enqueue(new Item(DataTokenType.StartArray));
            }
        }

        _queue.Enqueue(new Item(value));
        if (_arrayName == name && !repeats) {
            _arrayName = null;
            _queue.Enqueue(new Item(DataTokenType.EndArray));
        }
    }

    private void SkipNonElements(List<string> text) {
        XmlReader xml = _xml!;
        while (!xml.EOF) {
            switch (xml.NodeType) {
                case XmlNodeType.Text:
                case XmlNodeType.CDATA:
                case XmlNodeType.SignificantWhitespace:
                    text.Add(xml.Value);
                    break;
                case XmlNodeType.Comment:
                    Context.Report(OfficeDiagnosticCodes.CommentsOmitted, OfficeConversionLossKind.Omission,
                        "XML comments inside elements were dropped.");
                    break;
                case XmlNodeType.ProcessingInstruction:
                    ReportInstruction();
                    break;
                case XmlNodeType.Whitespace:
                    break;
                default:
                    return;
            }

            xml.Read();
        }
    }

    /// <summary>Reads the element at the reader into a buffered node, grouping repeated children into arrays.</summary>
    private DataNode ReadElement() {
        XmlReader xml = _xml!;
        bool nil = IsNil();
        bool empty = xml.IsEmptyElement;
        List<KeyValuePair<string, DataNode>> attributes = ReadAttributes();
        if (nil) {
            xml.Skip();
            return DataNode.Scalar(DataScalarKind.Null, string.Empty);
        }

        xml.Read();
        if (empty) return attributes.Count == 0 ? DataNode.Scalar(DataScalarKind.String, string.Empty) : Object(attributes, null, null);
        var text = new List<string>();
        Dictionary<string, List<DataNode>>? children = null;
        List<string>? order = null;
        string? previous = null;
        while (true) {
            SkipNonElements(text);
            if (xml.NodeType != XmlNodeType.Element) break;
            string name = Name(xml.Name);
            children ??= new Dictionary<string, List<DataNode>>(StringComparer.Ordinal);
            order ??= new List<string>();
            if (!children.TryGetValue(name, out List<DataNode>? list)) {
                list = new List<DataNode>();
                children.Add(name, list);
                order.Add(name);
            } else if (previous != name) {
                Context.Report(OfficeDiagnosticCodes.XmlSiblingOrderChanged, OfficeConversionLossKind.Approximation,
                    "Repeated elements separated by other elements were grouped into one array, which changes their order among siblings.");
            }

            previous = name;
            list.Add(ReadElement());
        }

        xml.Read();
        if (children == null && attributes.Count == 0) return Scalar(string.Concat(text));
        if (children != null && text.Count > 0) ReportMixed();
        return Object(attributes, order == null ? null : Group(children!, order), text);
    }

    private static List<KeyValuePair<string, DataNode>> Group(Dictionary<string, List<DataNode>> children, List<string> order) {
        var grouped = new List<KeyValuePair<string, DataNode>>(order.Count);
        foreach (string name in order) {
            List<DataNode> items = children[name];
            if (items.Count == 1) {
                grouped.Add(new KeyValuePair<string, DataNode>(name, items[0]));
            } else {
                DataNode array = DataNode.Array();
                array.Items!.AddRange(items);
                grouped.Add(new KeyValuePair<string, DataNode>(name, array));
            }
        }

        return grouped;
    }

    private DataNode Object(List<KeyValuePair<string, DataNode>> attributes, List<KeyValuePair<string, DataNode>>? children, List<string>? text) {
        DataNode node = DataNode.Object();
        node.Properties!.AddRange(attributes);
        if (children != null) node.Properties.AddRange(children);
        if (text is { Count: > 0 }) node.Properties.Add(new KeyValuePair<string, DataNode>("#text", TextNode(text)));
        return node;
    }

    private DataNode TextNode(List<string> runs) {
        if (runs.Count == 1) return Scalar(runs[0]);
        DataNode array = DataNode.Array();
        foreach (string run in runs) array.Items!.Add(Scalar(run));
        return array;
    }

    private bool IsNil() {
        string? nil = _xml!.GetAttribute("nil", XsiNamespace);
        return nil == "true" || nil == "1";
    }

    private List<KeyValuePair<string, DataNode>> ReadAttributes() {
        XmlReader xml = _xml!;
        var attributes = new List<KeyValuePair<string, DataNode>>();
        if (!xml.HasAttributes) return attributes;
        for (bool more = xml.MoveToFirstAttribute(); more; more = xml.MoveToNextAttribute()) {
            if (xml.LocalName == "nil" && xml.NamespaceURI == XsiNamespace) continue;
            attributes.Add(new KeyValuePair<string, DataNode>("@" + Name(xml.Name), Scalar(xml.Value)));
        }

        xml.MoveToElement();
        return attributes;
    }

    private DataNode Scalar(string text) {
        if (_infer && TableTypes.Infer(text.AsSpan(), '.') is { } kind && kind != DataScalarKind.String) {
            return DataNode.Scalar(kind, kind == DataScalarKind.Boolean ? (text.Length == 4 ? "true" : "false") : text);
        }

        return DataNode.Scalar(DataScalarKind.String, text);
    }

    private void ReportMixed() =>
        Context.Report(OfficeDiagnosticCodes.XmlMixedContent, OfficeConversionLossKind.Approximation,
            "Text between child elements became #text; its position among the elements was not kept.");

    private void ReportInstruction() =>
        Context.Report(OfficeDiagnosticCodes.XmlProcessingInstructionDropped, OfficeConversionLossKind.Omission,
            "XML processing instructions (such as <?" + _xml!.Name + "?>) were dropped.");

    private static string Name(string name) => XmlConvert.DecodeName(name);

    public override void Dispose() => _xml?.Dispose();
}
