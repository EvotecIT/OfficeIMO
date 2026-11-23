using System.Text;

namespace OfficeIMO.Core.Internal.OpenXml;

internal enum OfficeXmlNodeKind : byte {
    None,
    StartElement,
    EndElement,
    Text
}

/// <summary>
/// Forward-only reader over the UTF-8 bytes of one Open XML part. It resolves namespace prefixes per element scope,
/// decodes entity and character references on demand, rejects document type declarations, and applies ECMA-376
/// Part 3 markup compatibility: <c>mc:AlternateContent</c> selects the first <c>mc:Choice</c> whose
/// <c>Requires</c> namespaces are understood or else <c>mc:Fallback</c>, and elements in ignorable namespaces that
/// are not understood are skipped unless <c>mc:ProcessContent</c> keeps their children. Compatibility wrappers are
/// transparent: the content they select appears as children of the wrapper's parent.
/// </summary>
/// <remarks>An empty element is reported as a <see cref="OfficeXmlNodeKind.StartElement"/> with
/// <see cref="IsEmptyElement"/> set, followed by its <see cref="OfficeXmlNodeKind.EndElement"/>.</remarks>
internal sealed partial class OfficeXmlPullReader {
    private const int MaximumDepth = 1024;

    private readonly byte[] _data;
    private readonly int _end;
    private readonly OfficeXmlNamespaces.Table _namespaces;
    private readonly bool _processMarkupCompatibility;
    private int _position;

    private Frame[] _frames = new Frame[16];
    private int _frameCount;
    private int _depth;
    private bool _pendingEnd;
    private bool _decrementDepth;
    private bool _rootSeen;

    private Binding[] _bindings = new Binding[8];
    private int _bindingCount;
    private int[] _ignorable = Array.Empty<int>();
    private int _ignorableCount;
    private ProcessContentEntry[] _processContent = Array.Empty<ProcessContentEntry>();
    private int _processContentCount;

    private Attribute[] _attributes = new Attribute[8];
    private int _attributeCount;

    private int _nameStart;
    private int _nameLength;
    private int _textStart;
    private int _textLength;
    private bool _textNeedsDecoding;

    /// <summary>Creates a reader over <paramref name="length"/> bytes of a part. UTF-16 parts are transcoded first.</summary>
    /// <param name="data">The part bytes.</param>
    /// <param name="length">Number of valid bytes in <paramref name="data"/>.</param>
    /// <param name="namespaces">Namespace ids shared with other parts of the same package.</param>
    /// <param name="processMarkupCompatibility">False reports <c>mc:</c> elements and elements in ignorable
    /// namespaces as ordinary elements, for callers that decline such markup themselves.</param>
    internal OfficeXmlPullReader(byte[] data, int length, OfficeXmlNamespaces.Table? namespaces = null, bool processMarkupCompatibility = true) {
        ArgumentNullException.ThrowIfNull(data);
        if (length < 0 || length > data.Length) throw new ArgumentOutOfRangeException(nameof(length));
        _namespaces = namespaces ?? new OfficeXmlNamespaces.Table();
        _processMarkupCompatibility = processMarkupCompatibility;

        int start = 0;
        if (length >= 2 && (data[0] == 0xFF && data[1] == 0xFE || data[0] == 0xFE && data[1] == 0xFF
            || data[0] == 0x3C && data[1] == 0x00 || data[0] == 0x00 && data[1] == 0x3C)) {
            // ECMA-376 Part 2 allows UTF-16 parts. They are rare, so they are transcoded rather than read natively.
            bool bigEndian = data[0] == 0xFE || data[0] == 0x00;
            string text = (bigEndian ? Encoding.BigEndianUnicode : Encoding.Unicode).GetString(data, 0, length);
            if (text.Length > 0 && text[0] == '﻿') text = text.Substring(1);
            data = Encoding.UTF8.GetBytes(text);
            length = data.Length;
        } else if (length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF) {
            start = 3;
        }
        _data = data;
        _position = start;
        _end = length;
    }

    /// <summary>The namespace table shared by readers of the same package.</summary>
    internal OfficeXmlNamespaces.Table Namespaces => _namespaces;

    internal OfficeXmlNodeKind NodeKind { get; private set; }

    /// <summary>Depth of the current element counting only elements reported to the caller; the root is 1.
    /// For text, the depth of the element that contains it.</summary>
    internal int Depth => _depth;

    /// <summary>Namespace id of the current element.</summary>
    internal int NamespaceId { get; private set; }

    internal bool IsEmptyElement { get; private set; }

    /// <summary>For text: whether an <c>xml:space="preserve"</c> scope contains it.</summary>
    internal bool PreserveSpace { get; private set; }

    /// <summary>For text: whether a comment or processing instruction precedes it in the same element.</summary>
    internal bool FollowsCommentOrInstruction { get; private set; }

    /// <summary>For text: whether it came from a CDATA section.</summary>
    internal bool IsCData { get; private set; }

    /// <summary>For text outside CDATA: whether it is only XML whitespace (space, tab, carriage return, line feed).</summary>
    internal bool IsWhitespace { get; private set; }

    /// <summary>Compares the local name of the current element with an ASCII name.</summary>
    internal bool LocalNameEquals(string name) => AsciiEquals(_nameStart, _nameLength, name);

    internal bool IsElement(int namespaceId, string localName) => NamespaceId == namespaceId && AsciiEquals(_nameStart, _nameLength, localName);

    /// <summary>The local name of the current element. Names repeat across the parts of a package, so the
    /// package's namespace table caches them rather than decoding them again.</summary>
    internal string LocalName {
        get {
            int hash = _nameLength;
            for (int index = 0; index < _nameLength; index++) hash = hash * 31 + _data[_nameStart + index];
            ref string? cached = ref _namespaces.NameSlot(hash);
            if (cached != null && AsciiEquals(_nameStart, _nameLength, cached)) return cached;
            return cached = Encoding.UTF8.GetString(_data, _nameStart, _nameLength);
        }
    }

    /// <summary>Advances to the next element boundary or text node the caller should see.</summary>
    /// <returns><see langword="false"/> at the end of the part.</returns>
    /// <exception cref="System.Xml.XmlException">The part is not well-formed XML, declares a DTD or uses an unbound prefix.</exception>
    /// <exception cref="InvalidDataException">The part requires a namespace through <c>mc:MustUnderstand</c> that is not understood.</exception>
    internal bool Read() {
        if (_decrementDepth) {
            _decrementDepth = false;
            _depth--;
        }
        while (true) {
            if (_pendingEnd) {
                _pendingEnd = false;
                if (PopFrame()) return true;
                continue;
            }
            if (_position >= _end) {
                if (_frameCount != 0) throw Malformed("The part ends inside an element.");
                if (!_rootSeen) throw Malformed("The part has no root element.");
                NodeKind = OfficeXmlNodeKind.None;
                return false;
            }
            if (_data[_position] != (byte)'<') {
                if (ReadCharacterData()) return true;
                continue;
            }
            if (_position + 1 >= _end) throw Malformed("The part ends inside markup.");
            byte next = _data[_position + 1];
            if (next == (byte)'/') {
                ReadEndTag();
                if (PopFrame()) return true;
            } else if (next == (byte)'?') {
                SkipPast("?>", 2);
                MarkCommentOrInstruction();
            } else if (next == (byte)'!') {
                if (ReadDeclaration()) return true;
            } else if (ReadStartTag()) {
                return true;
            }
        }
    }

    /// <summary>Skips the rest of the current element, leaving the reader on its end tag.</summary>
    internal void SkipElement() {
        if (NodeKind != OfficeXmlNodeKind.StartElement) return;
        int depth = _depth;
        while (Read()) {
            if (NodeKind == OfficeXmlNodeKind.EndElement && _depth == depth) return;
        }
    }

    private bool ReadCharacterData() {
        int start = _position;
        bool needsDecoding = false;
        bool whitespace = true;
        int position = start;
        byte[] data = _data;
        while (position < _end) {
            byte value = data[position];
            if (value == (byte)'<') break;
            if (value == (byte)'&' || value == (byte)'\r') {
                needsDecoding = true;
                if (value == (byte)'&') whitespace = false;
            } else if (value != (byte)' ' && value != (byte)'\n' && value != (byte)'\t') {
                whitespace = false;
                if (value < 0x20 || value == 0xEF) RejectNonCharacter(position);
            }
            position++;
        }
        _position = position;
        if (_frameCount == 0) {
            if (!whitespace) throw Malformed("The part has text outside its root element.");
            return false;
        }
        // Only compatibility elements may appear directly inside mc:AlternateContent.
        if (_frames[_frameCount - 1].Kind == FrameKind.AlternateContent) return false;
        SetText(start, position - start, needsDecoding, cdata: false, whitespace);
        return true;
    }

    private bool ReadDeclaration() {
        if (StartsWith(_position, "<!--")) {
            SkipPast("-->", 4);
            MarkCommentOrInstruction();
            return false;
        }
        if (StartsWith(_position, "<![CDATA[")) {
            int start = _position + 9;
            int end = IndexOf(start, "]]>");
            if (end < 0) throw Malformed("A CDATA section is not closed.");
            for (int index = start; index < end; index++) {
                byte value = _data[index];
                if (value < 0x20 && value != (byte)'\t' && value != (byte)'\n' && value != (byte)'\r' || value == 0xEF) RejectNonCharacter(index);
            }
            _position = end + 3;
            if (_frameCount == 0) throw Malformed("The part has a CDATA section outside its root element.");
            if (_frames[_frameCount - 1].Kind == FrameKind.AlternateContent) return false;
            SetText(start, end - start, needsDecoding: HasCarriageReturn(start, end - start), cdata: true, whitespace: false);
            return true;
        }
        // ECMA-376 Part 2 forbids document type declarations in package parts; they also enable entity expansion.
        throw Malformed("The part has a document type declaration.");
    }

    private void MarkCommentOrInstruction() {
        if (_frameCount > 0) _frames[_frameCount - 1].SawCommentOrInstruction = true;
    }

    private void SetText(int start, int length, bool needsDecoding, bool cdata, bool whitespace) {
        NodeKind = OfficeXmlNodeKind.Text;
        _textStart = start;
        _textLength = length;
        _textNeedsDecoding = needsDecoding;
        IsCData = cdata;
        IsWhitespace = whitespace;
        PreserveSpace = _frames[_frameCount - 1].PreserveSpace;
        FollowsCommentOrInstruction = _frames[_frameCount - 1].SawCommentOrInstruction;
        IsEmptyElement = false;
    }

    private void ReadEndTag() {
        int start = _position + 2;
        int position = start;
        while (position < _end && !IsNameTerminator(_data[position])) position++;
        int length = position - start;
        while (position < _end && IsXmlSpace(_data[position])) position++;
        if (position >= _end || _data[position] != (byte)'>') throw Malformed("An end tag is not closed.");
        _position = position + 1;
        if (_frameCount == 0) throw Malformed("An end tag has no matching start tag.");
        ref Frame frame = ref _frames[_frameCount - 1];
        if (frame.QNameLength != length || !BytesEqual(frame.QNameStart, start, length)) {
            throw Malformed("An end tag does not match its start tag.");
        }
    }

    // Pops the innermost frame and reports whether it produced an EndElement the caller should see.
    private bool PopFrame() {
        ref Frame frame = ref _frames[--_frameCount];
        _bindingCount = frame.BindingMark;
        _ignorableCount = frame.IgnorableMark;
        _processContentCount = frame.ProcessContentMark;
        if (frame.Kind != FrameKind.Element) return false;
        NodeKind = OfficeXmlNodeKind.EndElement;
        NamespaceId = frame.NamespaceId;
        _nameStart = frame.LocalStart;
        _nameLength = frame.LocalLength;
        IsEmptyElement = false;
        _attributeCount = 0;
        // Depth still reports the closing element; its parent becomes current on the next read.
        _decrementDepth = true;
        return true;
    }

    // XML 1.0 characters exclude C0 controls other than tab, line feed and carriage return, and U+FFFE and U+FFFF
    // (EF BF BE and EF BF BF in UTF-8). Other lead bytes reach here only to be checked.
    private void RejectNonCharacter(int position) {
        byte value = _data[position];
        if (value == 0xEF) {
            if (position + 2 < _end && _data[position + 1] == 0xBF && (_data[position + 2] == 0xBE || _data[position + 2] == 0xBF)) {
                throw Malformed("The part contains a character XML does not allow.");
            }
            return;
        }
        throw Malformed("The part contains a character XML does not allow.");
    }

    /// <summary>
    /// UTF-16 code units of the XML text in a part, the unit XML readers count against a character limit: the byte order
    /// mark is not counted, and UTF-16 parts count two bytes per unit.
    /// </summary>
    internal static long CountCharacters(byte[] data, int length) {
        if (length >= 2 && (data[0] == 0xFF && data[1] == 0xFE || data[0] == 0xFE && data[1] == 0xFF)) return (length - 2) / 2;
        if (length >= 2 && (data[0] == 0x3C && data[1] == 0x00 || data[0] == 0x00 && data[1] == 0x3C)) return length / 2;
        int start = length >= 3 && data[0] == 0xEF && data[1] == 0xBB && data[2] == 0xBF ? 3 : 0;
        long count = 0;
        for (int index = start; index < length; index++) {
            byte value = data[index];
            if ((value & 0xC0) != 0x80) count += value >= 0xF0 ? 2 : 1;
        }
        return count;
    }

    private static System.Xml.XmlException Malformed(string message) => new("The Open XML part is not well-formed: " + message);
}
