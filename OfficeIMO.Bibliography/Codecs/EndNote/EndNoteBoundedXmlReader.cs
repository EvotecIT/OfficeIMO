using System.Xml;

namespace OfficeIMO.Bibliography;

internal sealed class EndNoteBoundedXmlReader : XmlReader, IXmlLineInfo {
    private readonly XmlReader _reader;
    private readonly BibliographyLimitGuard _limits;
    private readonly IList<BibliographyItem> _partialItems;
    private readonly EndNoteSourceOffsetMap _offsets;
    private readonly CancellationToken _cancellationToken;

    internal EndNoteBoundedXmlReader(XmlReader reader, BibliographyLimitGuard limits, IList<BibliographyItem> partialItems, EndNoteSourceOffsetMap offsets, CancellationToken cancellationToken) {
        _reader = reader;
        _limits = limits;
        _partialItems = partialItems;
        _offsets = offsets;
        _cancellationToken = cancellationToken;
    }

    public override bool Read() {
        _cancellationToken.ThrowIfCancellationRequested();
        bool read = _reader.Read();
        _cancellationToken.ThrowIfCancellationRequested();
        if (read && _reader.NodeType == XmlNodeType.Element) _limits.CheckDepth(_partialItems, _reader.Depth + 1, _offsets.GetOffset(this));
        return read;
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
    private readonly int[] _lineStarts;

    internal EndNoteSourceOffsetMap(string source, int baseOffset) {
        var lineStarts = new List<int> { baseOffset };
        for (int index = 0; index < source.Length; index++) {
            if (source[index] == '\r') {
                if (index + 1 < source.Length && source[index + 1] == '\n') index++;
                lineStarts.Add(baseOffset + index + 1);
            } else if (source[index] == '\n') lineStarts.Add(baseOffset + index + 1);
        }
        _lineStarts = lineStarts.ToArray();
    }

    internal int GetOffset(IXmlLineInfo lineInfo) {
        if (!lineInfo.HasLineInfo() || lineInfo.LineNumber < 1 || lineInfo.LineNumber > _lineStarts.Length || lineInfo.LinePosition < 1) return -1;
        int lineStart = _lineStarts[lineInfo.LineNumber - 1];
        return Math.Max(lineStart, lineStart + lineInfo.LinePosition - 2);
    }
}

internal sealed class EndNoteSourceOffset {
    internal EndNoteSourceOffset(int value) => Value = value;
    internal int Value { get; }
}
