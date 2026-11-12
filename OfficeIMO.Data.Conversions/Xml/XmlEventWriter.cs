using System.Xml;

namespace OfficeIMO.Data.Conversions;

/// <summary>
/// Writes data as XML with the convention <see cref="XmlEventReader"/> reads: a root object with one property names the
/// root element, <c>@name</c> members are attributes, <c>#text</c> is text, array items repeat their property's element
/// name (items of nested arrays are <c>item</c>), and null is <c>xsi:nil="true"</c>. Any other root is wrapped in
/// <see cref="DataConversionOptions.RootName"/>, which is reported.
/// </summary>
internal sealed class XmlEventWriter : DataEventWriter {
    private const string XsiNamespace = "http://www.w3.org/2001/XMLSchema-instance";
    private const string XmlnsNamespace = "http://www.w3.org/2000/xmlns/";
    private readonly XmlWriter _xml;
    private readonly bool _ownsWriter;
    private readonly bool _wrapDocuments;
    private readonly bool _rootIsNamed;
    private readonly Stack<Frame> _frames = new();
    private readonly List<(DataTokenType Type, DataScalarKind Kind, string? Text)> _buffer = new();
    private string? _name;
    private string? _pendingElement;
    private readonly List<(string Name, string Value)> _pendingAttributes = new();
    private readonly List<(int Depth, string Prefix, string Uri)> _scopes = new();
    private int _depth;
    private int _documents;
    private bool _buffering;

    private sealed class Frame {
        public Frame(string? element, bool isArray, string itemName) {
            Element = element;
            IsArray = isArray;
            ItemName = itemName;
        }

        /// <summary>The element this frame closes at its end; null when it opened none (a property's array, a named root).</summary>
        public string? Element { get; }
        public bool IsArray { get; }
        public string ItemName { get; }
        public int Items;
        public bool WroteChild;
        public bool IsArrayWrapper;
    }

    /// <param name="writer">The XML output.</param>
    /// <param name="documents">How many documents the source holds; null for a list of records such as JSON Lines.</param>
    /// <param name="rootIsNamed">Whether the source guarantees a root object with one property (XML read back), so nothing is buffered.</param>
    /// <param name="context">The conversion context.</param>
    /// <param name="ownsWriter">Whether disposing this writer disposes <paramref name="writer"/>.</param>
    public XmlEventWriter(XmlWriter writer, long? documents, bool rootIsNamed, DataConversionContext context, bool ownsWriter) : base(context) {
        _xml = writer;
        _ownsWriter = ownsWriter;
        _wrapDocuments = documents is null or > 1;
        _rootIsNamed = rootIsNamed;
    }

    public static XmlEventWriter Create(Stream output, long? documents, bool rootIsNamed, DataConversionContext context) {
        XmlWriterSettings settings = context.Options.XmlWriter?.Clone() ?? new XmlWriterSettings { Indent = true, IndentChars = "  ", NewLineChars = "\n" };
        if (context.Options.XmlWriter == null) settings.Encoding = new UTF8Encoding(false);
        settings.CloseOutput = false;
        return new XmlEventWriter(XmlWriter.Create(output, settings), documents, rootIsNamed, context, ownsWriter: true);
    }

    public override void WriteStartDocument() {
        if (_documents++ == 0) {
            _xml.WriteStartDocument();
            if (_wrapDocuments) {
                OpenWrapper("Several documents became <item> elements of one root element.");
                _frames.Peek().IsArrayWrapper = true;
            }
        }

        Context.ItemsWritten++;
    }

    public override void WriteEndDocument() {
    }

    public override void WriteStartObject() {
        if (Buffered(DataTokenType.StartObject)) return;
        if (_frames.Count == 0) {
            if (!_rootIsNamed) {
                _buffering = true;
                _buffer.Add((DataTokenType.StartObject, default, null));
                return;
            }

            _frames.Push(new Frame(null, false, string.Empty));
            return;
        }

        string element = TakeName();
        OpenElement(element);
        _frames.Push(new Frame(element, false, string.Empty));
    }

    public override void WriteEndObject() {
        if (Buffered(DataTokenType.EndObject)) return;
        Close();
    }

    public override void WriteStartArray() {
        if (Buffered(DataTokenType.StartArray)) return;
        if (_frames.Count == 0) {
            OpenWrapper("A root array was wrapped in a root element; its items became <item> elements.");
            _frames.Peek().IsArrayWrapper = true;
            return;
        }

        Frame parent = _frames.Peek();
        if (parent.IsArray) {
            string element = TakeName();
            OpenElement(element);
            _frames.Push(new Frame(element, true, "item"));
            ReportArrayShape();
            return;
        }

        _frames.Push(new Frame(null, true, TakeName()));
    }

    public override void WriteEndArray() {
        if (Buffered(DataTokenType.EndArray)) return;
        Frame frame = _frames.Peek();
        if (frame.Items <= 1 && !frame.IsArrayWrapper) ReportArrayShape();
        Close();
    }

    public override void WritePropertyName(ReadOnlySpan<char> name) {
        if (Buffered(DataTokenType.PropertyName, default, name)) return;
        _name = name.ToString();
    }

    public override void WriteScalar(DataScalarKind kind, ReadOnlySpan<char> text) {
        if (Buffered(DataTokenType.Scalar, kind, text)) return;
        if (_frames.Count == 0) {
            Report("A root value was wrapped in a root element.");
            WriteScalarElement(Context.Options.RootName, kind, text);
            return;
        }

        Frame frame = _frames.Peek();
        string name = TakeName();
        if (!frame.IsArray && frame.Element != null && name.Length > 1 && name[0] == '@' && _pendingElement != null) {
            if (kind == DataScalarKind.Null) return;
            CheckType(kind);
            _pendingAttributes.Add((name.Substring(1), Clean(text.ToString())));
            return;
        }

        if (!frame.IsArray && frame.Element != null && name == "#text") {
            CheckType(kind);
            FlushStartTag();
            _xml.WriteString(Clean(text.ToString()));
            return;
        }

        WriteScalarElement(name, kind, text);
    }

    public override void WriteComment(ReadOnlySpan<char> text) {
        if (_buffering) return;
        FlushStartTag();
        _xml.WriteComment(text.ToString().Replace("--", "- -"));
    }

    private void Close() {
        Frame frame = _frames.Pop();
        if (frame.Element != null) EndElement();
    }

    private void WriteScalarElement(string name, DataScalarKind kind, ReadOnlySpan<char> text) {
        OpenElement(name);
        FlushStartTag();
        if (kind == DataScalarKind.Null) {
            _xml.WriteAttributeString("xsi", "nil", XsiNamespace, "true");
        } else {
            CheckType(kind);
            if (text.Length > 0) _xml.WriteString(Clean(text.ToString()));
        }

        EndElement();
    }

    // Start tags wait for their attributes, so namespace declarations among them can resolve the element's own prefix.
    private void OpenElement(string name) {
        FlushStartTag();
        if (_frames.Count > 0) _frames.Peek().WroteChild = true;
        _pendingElement = name;
    }

    private void FlushStartTag() {
        if (_pendingElement == null) return;
        string name = _pendingElement;
        _pendingElement = null;
        _depth++;
        foreach ((string attribute, string value) in _pendingAttributes) {
            if (attribute == "xmlns") _scopes.Add((_depth, string.Empty, value));
            else if (attribute.StartsWith("xmlns:", StringComparison.Ordinal)) _scopes.Add((_depth, attribute.Substring(6), value));
        }

        (string prefix, string local) = Split(name);
        string? ns = Resolve(prefix);
        if (prefix.Length > 0 && ns == null) {
            _xml.WriteStartElement(XmlConvert.EncodeLocalName(name));
        } else {
            _xml.WriteStartElement(prefix, XmlConvert.EncodeLocalName(local), ns ?? string.Empty);
        }

        foreach ((string attribute, string value) in _pendingAttributes) {
            if (attribute == "xmlns") {
                _xml.WriteAttributeString("xmlns", XmlnsNamespace, value);
            } else if (attribute.StartsWith("xmlns:", StringComparison.Ordinal)) {
                _xml.WriteAttributeString("xmlns", attribute.Substring(6), XmlnsNamespace, value);
            } else {
                (string attributePrefix, string attributeLocal) = Split(attribute);
                string? attributeNs = attributePrefix.Length > 0 ? Resolve(attributePrefix) : null;
                if (attributeNs != null) _xml.WriteAttributeString(attributePrefix, XmlConvert.EncodeLocalName(attributeLocal), attributeNs, value);
                else _xml.WriteAttributeString(XmlConvert.EncodeLocalName(attribute), value);
            }
        }

        _pendingAttributes.Clear();
    }

    private void EndElement() {
        FlushStartTag();
        _xml.WriteEndElement();
        while (_scopes.Count > 0 && _scopes[_scopes.Count - 1].Depth >= _depth) _scopes.RemoveAt(_scopes.Count - 1);
        _depth--;
    }

    private string? Resolve(string prefix) {
        for (int i = _scopes.Count - 1; i >= 0; i--) {
            if (_scopes[i].Prefix == prefix) return _scopes[i].Uri;
        }

        return prefix == "xml" ? "http://www.w3.org/XML/1998/namespace" : null;
    }

    private static (string Prefix, string Local) Split(string name) {
        int colon = name.IndexOf(':');
        return colon > 0 && colon < name.Length - 1 ? (name.Substring(0, colon), name.Substring(colon + 1)) : (string.Empty, name);
    }

    // XML 1.0 cannot hold some characters at all, not even as references (NUL, most C0 controls, U+FFFE, U+FFFF).
    private string Clean(string text) {
        int bad = -1;
        for (int i = 0; i < text.Length; i++) {
            char c = text[i];
            if (char.IsHighSurrogate(c) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1])) {
                i++;
                continue;
            }

            if (!XmlConvert.IsXmlChar(c)) {
                bad = i;
                break;
            }
        }

        if (bad < 0) return text;
        Context.Report(OfficeDiagnosticCodes.XmlCharacterReplaced, OfficeConversionLossKind.Approximation,
            "Characters XML cannot hold (such as NUL, control characters or U+FFFF) were replaced with U+FFFD.");
        var builder = new StringBuilder(text.Length);
        for (int i = 0; i < text.Length; i++) {
            char c = text[i];
            if (char.IsHighSurrogate(c) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1])) {
                builder.Append(c).Append(text[++i]);
            } else {
                builder.Append(XmlConvert.IsXmlChar(c) ? c : '\uFFFD');
            }
        }

        return builder.ToString();
    }

    private string TakeName() {
        Frame frame = _frames.Peek();
        if (frame.IsArray) {
            frame.Items++;
            return frame.ItemName;
        }

        string name = _name ?? throw new InvalidOperationException("A value in an object has no property name.");
        _name = null;
        if (name.Length == 0) {
            Context.Report(OfficeDiagnosticCodes.XmlNameChanged, OfficeConversionLossKind.Approximation,
                "Empty keys have no XML element name; they became <_> elements.");
            return "_";
        }

        return name;
    }

    private void OpenWrapper(string message) {
        Report(message);
        string root = Context.Options.RootName;
        OpenElement(root);
        _frames.Push(new Frame(root, true, "item"));
    }

    private void Report(string message) =>
        Context.Report(OfficeDiagnosticCodes.RootWrapped, OfficeConversionLossKind.Approximation,
            message + " The root element is <" + Context.Options.RootName + ">.");

    private void CheckType(DataScalarKind kind) {
        if (kind == DataScalarKind.String) return;
        Context.Report(OfficeDiagnosticCodes.ScalarTypeDropped, OfficeConversionLossKind.Approximation,
            "XML text has no types; numbers, booleans and date-times became text.");
    }

    private void ReportArrayShape() =>
        Context.Report(OfficeDiagnosticCodes.XmlArrayShape, OfficeConversionLossKind.Approximation,
            "Empty, single-item or nested arrays cannot be told apart from plain elements in XML.");

    // The root object is held back until it shows whether it has exactly one property whose value is not an array.
    private bool Buffered(DataTokenType type, DataScalarKind kind = default, ReadOnlySpan<char> text = default) {
        if (!_buffering) return false;
        _buffer.Add((type, kind, type is DataTokenType.PropertyName or DataTokenType.Scalar ? text.ToString() : null));
        int depth = 0, properties = 0;
        bool firstValueIsArray = false;
        for (int i = 0; i < _buffer.Count; i++) {
            DataTokenType t = _buffer[i].Type;
            if (t == DataTokenType.PropertyName && depth == 1) properties++;
            if (t == DataTokenType.StartArray && depth == 1 && properties == 1) firstValueIsArray = true;
            if (t is DataTokenType.StartObject or DataTokenType.StartArray) depth++;
            else if (t is DataTokenType.EndObject or DataTokenType.EndArray) depth--;
        }

        if (properties > 1 || firstValueIsArray) Release(named: false);
        else if (depth == 0) Release(named: properties == 1);
        return true;
    }

    private void Release(bool named) {
        _buffering = false;
        var events = _buffer.ToArray();
        _buffer.Clear();
        if (named) {
            _frames.Push(new Frame(null, false, string.Empty));
        } else {
            Report("The root object did not have exactly one property, so it was wrapped in a root element.");
            string root = Context.Options.RootName;
            OpenElement(root);
            _frames.Push(new Frame(root, false, string.Empty));
        }

        for (int i = 1; i < events.Length; i++) {
            (DataTokenType type, DataScalarKind kind, string? text) = events[i];
            switch (type) {
                case DataTokenType.StartObject: WriteStartObject(); break;
                case DataTokenType.EndObject: WriteEndObject(); break;
                case DataTokenType.StartArray: WriteStartArray(); break;
                case DataTokenType.EndArray: WriteEndArray(); break;
                case DataTokenType.PropertyName: WritePropertyName(text.AsSpan()); break;
                case DataTokenType.Scalar: WriteScalar(kind, text.AsSpan()); break;
            }
        }
    }

    public override void Complete() {
        if (_buffering) Release(named: false);
        while (_frames.Count > 0) Close();
        FlushStartTag();
        if (_documents == 0) {
            _xml.WriteStartDocument();
            _xml.WriteStartElement(XmlConvert.EncodeName(Context.Options.RootName));
            _xml.WriteAttributeString("xsi", "nil", XsiNamespace, "true");
            _xml.WriteEndElement();
        }

        _xml.WriteEndDocument();
        _xml.Flush();
    }

    public override void Dispose() {
        if (_ownsWriter) _xml.Dispose();
    }
}
