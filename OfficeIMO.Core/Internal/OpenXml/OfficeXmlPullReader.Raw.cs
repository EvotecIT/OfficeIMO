using System.Text;

namespace OfficeIMO.Core.Internal.OpenXml;

/// <summary>Source positions and attribute access for callers that keep markup as it was written.</summary>
internal sealed partial class OfficeXmlPullReader {
    private int _markupStart;
    private int _markupEnd;

    /// <summary>The part bytes the reader works on; UTF-16 parts are transcoded to UTF-8 first.</summary>
    internal byte[] Data => _data;

    /// <summary>Whether the part was UTF-16 and <see cref="Data"/> holds it transcoded to UTF-8.</summary>
    internal bool Transcoded { get; }

    /// <summary>Number of valid bytes in <see cref="Data"/>.</summary>
    internal int DataLength => _end;

    /// <summary>
    /// Where the markup of the current node starts in <see cref="Data"/>: the <c>&lt;</c> of a start or end tag, the
    /// first byte of text, or the <c>&lt;![CDATA[</c> of a CDATA section. The end of an empty element has no markup;
    /// its start and end both sit where the start tag ends.
    /// </summary>
    internal int MarkupStart => _markupStart;

    /// <summary>Where the markup of the current node ends in <see cref="Data"/>, exclusive.</summary>
    internal int MarkupEnd => _markupEnd;

    /// <summary>The exact URI id of the current element's namespace; Strict and transitional spellings differ here.</summary>
    internal int ElementUriId {
        get {
            ref Frame frame = ref _frames[_frameCount - 1];
            int prefixLength = frame.LocalStart == frame.QNameStart ? 0 : frame.LocalStart - frame.QNameStart - 1;
            return UriIdOfPrefix(frame.QNameStart, prefixLength, frame.NamespaceId);
        }
    }

    /// <summary>The prefix the current element is written with, or an empty string.</summary>
    internal string ElementPrefix {
        get {
            ref Frame frame = ref _frames[_frameCount - 1];
            return frame.LocalStart == frame.QNameStart ? string.Empty : CachedName(frame.QNameStart, frame.LocalStart - frame.QNameStart - 1);
        }
    }

    /// <summary>Number of attributes on the current start element, namespace declarations included.</summary>
    internal int AttributeCount => NodeKind == OfficeXmlNodeKind.StartElement ? _attributeCount : 0;

    internal string GetAttributeLocalName(int index) {
        ref Attribute attribute = ref _attributes[index];
        return CachedName(attribute.LocalStart, attribute.LocalLength);
    }

    internal string GetAttributePrefix(int index) {
        ref Attribute attribute = ref _attributes[index];
        return attribute.PrefixLength == 0 ? string.Empty : CachedName(attribute.PrefixStart, attribute.PrefixLength);
    }

    /// <summary>The folded namespace id and exact URI id of an attribute. Namespace declarations report
    /// <see cref="OfficeXmlNamespaces.Xmlns"/>; unprefixed attributes have no namespace.</summary>
    internal int GetAttributeNamespace(int index, out int uriId) {
        ref Attribute attribute = ref _attributes[index];
        if (attribute.PrefixLength == 0) {
            int id = AsciiEquals(attribute.LocalStart, attribute.LocalLength, "xmlns") ? OfficeXmlNamespaces.Xmlns : OfficeXmlNamespaces.None;
            uriId = id;
            return id;
        }
        if (attribute.PrefixLength == 5 && AsciiEquals(attribute.PrefixStart, 5, "xmlns")) {
            uriId = OfficeXmlNamespaces.Xmlns;
            return OfficeXmlNamespaces.Xmlns;
        }
        int namespaceId = ResolvePrefix(attribute.PrefixStart, attribute.PrefixLength);
        uriId = UriIdOfPrefix(attribute.PrefixStart, attribute.PrefixLength, namespaceId);
        return namespaceId;
    }

    /// <summary>The bytes of an attribute value as written, between its quotes.</summary>
    internal void GetAttributeRawValue(int index, out int start, out int length) {
        ref Attribute attribute = ref _attributes[index];
        start = attribute.ValueStart;
        length = attribute.ValueLength;
    }

    /// <summary>The quote character an attribute value is written with.</summary>
    internal byte GetAttributeQuote(int index) => _data[_attributes[index].ValueStart - 1];

    internal string GetAttributeValue(int index) => GetAttributeValue(_attributes[index]);

    /// <summary>Decodes an attribute value written at <paramref name="start"/> in <see cref="Data"/>.</summary>
    internal string DecodeAttributeValue(int start, int length) {
        bool needsDecoding = false;
        for (int index = start; index < start + length && !needsDecoding; index++) {
            byte value = _data[index];
            needsDecoding = value == (byte)'&' || value == (byte)'\r' || value == (byte)'\n' || value == (byte)'\t';
        }
        return GetAttributeValue(new Attribute { ValueStart = start, ValueLength = length, NeedsDecoding = needsDecoding });
    }

    /// <summary>Decodes text written at <paramref name="start"/> in <see cref="Data"/>; for CDATA, the range between its delimiters.</summary>
    internal string DecodeTextValue(int start, int length, bool cdata) {
        // Most text has no references and no carriage returns, so its UTF-8 bytes are the value.
        if (!NeedsTextDecoding(start, length, cdata)) return Decode(start, length);
        var buffer = new OfficeCharBuffer(_decodeBuffer ??= new char[64]);
        if (cdata) AppendNormalizedLineEnds(buffer, start, length);
        else AppendDecoded(buffer, start, length, attribute: false);
        _decodeBuffer = buffer.Chars;
        return buffer.ToString();
    }

    private bool NeedsTextDecoding(int start, int length, bool cdata) {
        byte[] data = _data;
        int end = start + length;
        for (int index = start; index < end; index++) {
            byte value = data[index];
            if (value == (byte)'\r' || value == (byte)'&' && !cdata) return true;
        }
        return false;
    }

    private int UriIdOfPrefix(int start, int length, int namespaceId) {
        int found = FindBinding(start, length);
        return found >= 0 ? _bindings[found].UriId : OfficeXmlNamespaces.TransitionalUriId(namespaceId);
    }

    // Names repeat across the parts of a package, so the namespace table caches them.
    private string CachedName(int start, int length) {
        int hash = length;
        for (int index = 0; index < length; index++) hash = hash * 31 + _data[start + index];
        // Two slots per name, so two names whose hashes collide don't evict each other on every element.
        ref string? cached = ref _namespaces.NameSlot(hash);
        if (cached != null && AsciiEquals(start, length, cached)) return cached;
        ref string? alternate = ref _namespaces.NameSlot((int)((uint)hash * 2654435761u >> 7));
        if (alternate != null && AsciiEquals(start, length, alternate)) return alternate;
        string name = Encoding.UTF8.GetString(_data, start, length);
        if (cached == null) cached = name;
        else alternate = name;
        return name;
    }
}
