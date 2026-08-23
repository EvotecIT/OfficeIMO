namespace OfficeIMO.OpenDocument;

/// <summary>Enforces an XML depth bound while a consumer reads the document once.</summary>
internal sealed class OdfDepthLimitingXmlReader : XmlReader {
    private readonly XmlReader _inner;
    private readonly string _partPath;
    private readonly int _maxDepth;

    internal OdfDepthLimitingXmlReader(XmlReader inner, string partPath, int maxDepth) {
        _inner = inner ?? throw new ArgumentNullException(nameof(inner));
        _partPath = partPath ?? throw new ArgumentNullException(nameof(partPath));
        _maxDepth = maxDepth;
    }

    public override int AttributeCount => _inner.AttributeCount;
    public override string BaseURI => _inner.BaseURI;
    public override int Depth => _inner.Depth;
    public override bool EOF => _inner.EOF;
    public override bool HasValue => _inner.HasValue;
    public override bool IsEmptyElement => _inner.IsEmptyElement;
    public override string LocalName => _inner.LocalName;
    public override string NamespaceURI => _inner.NamespaceURI;
    public override XmlNameTable NameTable => _inner.NameTable;
    public override XmlNodeType NodeType => _inner.NodeType;
    public override string Prefix => _inner.Prefix;
    public override ReadState ReadState => _inner.ReadState;
    public override string Value => _inner.Value;

    public override string GetAttribute(int i) => _inner.GetAttribute(i);
    public override string? GetAttribute(string name) => _inner.GetAttribute(name);
    public override string? GetAttribute(string name, string? namespaceURI) => _inner.GetAttribute(name, namespaceURI);
    public override string? LookupNamespace(string prefix) => _inner.LookupNamespace(prefix);
    public override bool MoveToAttribute(string name) => _inner.MoveToAttribute(name);
    public override bool MoveToAttribute(string name, string? ns) => _inner.MoveToAttribute(name, ns);
    public override void MoveToAttribute(int i) => _inner.MoveToAttribute(i);
    public override bool MoveToElement() => _inner.MoveToElement();
    public override bool MoveToFirstAttribute() => _inner.MoveToFirstAttribute();
    public override bool MoveToNextAttribute() => _inner.MoveToNextAttribute();
    public override bool ReadAttributeValue() => _inner.ReadAttributeValue();
    public override void ResolveEntity() => _inner.ResolveEntity();

    public override bool Read() {
        bool result = _inner.Read();
        if (result && _inner.Depth > _maxDepth) {
            throw new InvalidDataException(
                $"OpenDocument XML part '{_partPath}' exceeds MaxXmlDepth ({_maxDepth}).");
        }
        return result;
    }

    protected override void Dispose(bool disposing) {
        if (disposing) _inner.Dispose();
        base.Dispose(disposing);
    }
}
