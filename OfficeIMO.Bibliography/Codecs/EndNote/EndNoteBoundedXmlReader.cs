using System.Xml;

namespace OfficeIMO.Bibliography;

internal sealed class EndNoteBoundedXmlReader : XmlReader, IXmlLineInfo {
    private readonly XmlReader _reader;
    private readonly BibliographyLimitGuard _limits;
    private readonly BibliographyLimitGuard _materializationLimits;
    private readonly IList<BibliographyItem> _partialItems;
    private readonly EndNoteSourceOffsetMap _offsets;
    private readonly CancellationToken _cancellationToken;
    private readonly List<string> _elementNames = new List<string>();
    private readonly List<string> _elementNamespaces = new List<string>();
    private readonly List<bool> _elementHasChildren = new List<bool>();
    private readonly List<int> _elementTextLengths = new List<int>();
    private readonly List<int> _elementOffsets = new List<int>();

    internal EndNoteBoundedXmlReader(XmlReader reader, BibliographyLimitGuard limits, BibliographyLimitGuard materializationLimits, IList<BibliographyItem> partialItems, EndNoteSourceOffsetMap offsets, CancellationToken cancellationToken) {
        _reader = reader;
        _limits = limits;
        _materializationLimits = materializationLimits;
        _partialItems = partialItems;
        _offsets = offsets;
        _cancellationToken = cancellationToken;
    }

    public override bool Read() {
        _cancellationToken.ThrowIfCancellationRequested();
        bool read = _reader.Read();
        _cancellationToken.ThrowIfCancellationRequested();
        if (read && _reader.NodeType == XmlNodeType.Element) {
            int offset = _offsets.GetOffset(this);
            if (_elementHasChildren.Count > 0) _elementHasChildren[_elementHasChildren.Count - 1] = true;
            _limits.CheckDepth(_partialItems, _reader.Depth + 1, offset);
            if (IsAcceptedRecordElement()) _limits.AddItem(_partialItems, offset);
            for (int index = 0; index < _reader.AttributeCount; index++) _materializationLimits.AddValue(_partialItems, _reader.GetAttribute(index), offset);
            if (!_reader.IsEmptyElement) {
                _elementNames.Add(_reader.LocalName);
                _elementNamespaces.Add(_reader.NamespaceURI);
                _elementHasChildren.Add(false);
                _elementTextLengths.Add(0);
                _elementOffsets.Add(offset);
            } else {
                _materializationLimits.AddValue(_partialItems, string.Empty, offset);
            }
        } else if (read && _reader.NodeType == XmlNodeType.EndElement && _elementNames.Count > 0) {
            int last = _elementNames.Count - 1;
            if (!_elementHasChildren[last]) {
                _materializationLimits.CheckValueLength(_partialItems, _elementTextLengths[last], _elementOffsets[last]);
                _materializationLimits.AddValue(_partialItems, null, _elementOffsets[last]);
            }
            _elementNames.RemoveAt(_elementNames.Count - 1);
            _elementNamespaces.RemoveAt(_elementNamespaces.Count - 1);
            _elementHasChildren.RemoveAt(last);
            _elementTextLengths.RemoveAt(last);
            _elementOffsets.RemoveAt(last);
        } else if (read && _elementTextLengths.Count > 0 && (_reader.NodeType == XmlNodeType.Text || _reader.NodeType == XmlNodeType.CDATA)) {
            int last = _elementTextLengths.Count - 1;
            _materializationLimits.CheckAdditionalValueLength(_partialItems, _elementTextLengths[last], _reader.Value.Length, _elementOffsets[last]);
            _elementTextLengths[last] += _reader.Value.Length;
        } else if (read && (_reader.NodeType == XmlNodeType.Comment || _reader.NodeType == XmlNodeType.ProcessingInstruction)) {
            int offset = _offsets.GetOffset(this);
            _materializationLimits.AddValue(_partialItems, _reader.Value, offset);
        }
        return read;
    }

    private bool IsAcceptedRecordElement() {
        if (!string.Equals(_reader.LocalName, "record", StringComparison.OrdinalIgnoreCase)) return false;
        int parentIndex = _elementNames.Count - 1;
        if (parentIndex < 0 || !string.Equals(_elementNames[parentIndex], "records", StringComparison.OrdinalIgnoreCase) || !string.Equals(_elementNamespaces[parentIndex], _reader.NamespaceURI, StringComparison.Ordinal)) return false;
        if (_elementNames.Count == 1) return true;
        return _elementNames.Count == 2 && string.Equals(_elementNamespaces[0], _elementNamespaces[parentIndex], StringComparison.Ordinal);
    }

    public override bool ReadAttributeValue() {
        _cancellationToken.ThrowIfCancellationRequested();
        return _reader.ReadAttributeValue();
    }

    public bool HasLineInfo() => (_reader as IXmlLineInfo)?.HasLineInfo() == true;
    public int LineNumber => (_reader as IXmlLineInfo)?.LineNumber ?? 0;
    public int LinePosition => (_reader as IXmlLineInfo)?.LinePosition ?? 0;
    public override int AttributeCount => _reader.AttributeCount;
    public override string BaseURI => _reader.BaseURI;
    public override int Depth => _reader.Depth;
    public override bool EOF => _reader.EOF;
    public override bool HasValue => _reader.HasValue;
    public override bool IsEmptyElement => _reader.IsEmptyElement;
    public override string LocalName => _reader.LocalName;
    public override string NamespaceURI => _reader.NamespaceURI;
    public override XmlNameTable NameTable => _reader.NameTable;
    public override XmlNodeType NodeType => _reader.NodeType;
    public override string Prefix => _reader.Prefix;
    public override ReadState ReadState => _reader.ReadState;
    public override string Value => _reader.Value;
    public override string? GetAttribute(string name) => _reader.GetAttribute(name);
    public override string? GetAttribute(string name, string? namespaceURI) => _reader.GetAttribute(name, namespaceURI);
    public override string GetAttribute(int index) => _reader.GetAttribute(index);
    public override string? LookupNamespace(string prefix) => _reader.LookupNamespace(prefix);
    public override bool MoveToAttribute(string name) => _reader.MoveToAttribute(name);
    public override bool MoveToAttribute(string name, string? namespaceURI) => _reader.MoveToAttribute(name, namespaceURI);
    public override void MoveToAttribute(int index) => _reader.MoveToAttribute(index);
    public override bool MoveToElement() => _reader.MoveToElement();
    public override bool MoveToFirstAttribute() => _reader.MoveToFirstAttribute();
    public override bool MoveToNextAttribute() => _reader.MoveToNextAttribute();
    public override void ResolveEntity() => _reader.ResolveEntity();
    public override void Close() => _reader.Close();

    protected override void Dispose(bool disposing) {
        if (disposing) _reader.Dispose();
        base.Dispose(disposing);
    }
}

internal sealed class EndNoteSourceOffsetMap {
    private readonly string _source;
    private readonly int _baseOffset;
    private readonly CancellationToken _cancellationToken;
    private int _scanIndex;
    private int _currentLine = 1;
    private int _currentLineStart;

    internal EndNoteSourceOffsetMap(string source, int baseOffset, CancellationToken cancellationToken) {
        _source = source;
        _baseOffset = baseOffset;
        _currentLineStart = baseOffset;
        _cancellationToken = cancellationToken;
    }

    internal int GetOffset(IXmlLineInfo lineInfo) {
        if (!lineInfo.HasLineInfo() || lineInfo.LineNumber < 1 || lineInfo.LinePosition < 1) return -1;
        if (lineInfo.LineNumber < _currentLine) Reset();
        while (_currentLine < lineInfo.LineNumber && _scanIndex < _source.Length) {
            if ((_scanIndex & 4095) == 0) _cancellationToken.ThrowIfCancellationRequested();
            char character = _source[_scanIndex++];
            if (character == '\r') {
                if (_scanIndex < _source.Length && _source[_scanIndex] == '\n') _scanIndex++;
                AdvanceLine();
            } else if (character == '\n') AdvanceLine();
        }
        _cancellationToken.ThrowIfCancellationRequested();
        if (_currentLine != lineInfo.LineNumber) return -1;
        return Math.Max(_currentLineStart, _currentLineStart + lineInfo.LinePosition - 2);
    }

    private void AdvanceLine() {
        _currentLine++;
        _currentLineStart = _baseOffset + _scanIndex;
    }

    private void Reset() {
        _scanIndex = 0;
        _currentLine = 1;
        _currentLineStart = _baseOffset;
    }
}

internal sealed class EndNoteSourceOffset {
    internal EndNoteSourceOffset(int value) => Value = value;
    internal int Value { get; }
}
