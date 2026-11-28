using System.Text;

namespace OfficeIMO.Core.Internal.OpenXml;

internal sealed partial class OfficeXmlPullReader {
    private static readonly UTF8Encoding StrictUtf8 = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);

    private char[]? _decodeBuffer;

    /// <summary>Reads an attribute of the current start element by namespace id and ASCII local name.</summary>
    /// <remarks>Unprefixed attributes have no namespace (<see cref="OfficeXmlNamespaces.None"/>).</remarks>
    internal bool TryGetAttribute(int namespaceId, string localName, out string value) {
        for (int index = 0; index < _attributeCount; index++) {
            ref Attribute attribute = ref _attributes[index];
            if (!AsciiEquals(attribute.LocalStart, attribute.LocalLength, localName)) continue;
            int attributeNamespace = attribute.PrefixLength == 0
                ? OfficeXmlNamespaces.None
                : ResolvePrefix(attribute.PrefixStart, attribute.PrefixLength);
            if (attributeNamespace != namespaceId) continue;
            value = GetAttributeValue(attribute);
            return true;
        }
        value = string.Empty;
        return false;
    }

    /// <summary>Returns the decoded text of the current text node.</summary>
    internal string GetText() {
        if (!_textNeedsDecoding) return Decode(_textStart, _textLength);
        var buffer = new OfficeCharBuffer();
        AppendText(buffer);
        return buffer.ToString();
    }

    /// <summary>Appends the decoded text of the current text node.</summary>
    internal void AppendText(OfficeCharBuffer target) {
        if (!_textNeedsDecoding) {
            target.AppendUtf8(_data, _textStart, _textLength);
            return;
        }
        if (IsCData) {
            AppendNormalizedLineEnds(target, _textStart, _textLength);
            return;
        }
        AppendDecoded(target, _textStart, _textLength, attribute: false);
    }

    private string GetAttributeValue(in Attribute attribute) =>
        attribute.NeedsDecoding ? DecodeAttribute(attribute) : Decode(attribute.ValueStart, attribute.ValueLength);

    private string DecodeAttribute(in Attribute attribute) {
        var buffer = new OfficeCharBuffer(_decodeBuffer ??= new char[64]);
        AppendDecoded(buffer, attribute.ValueStart, attribute.ValueLength, attribute: true);
        _decodeBuffer = buffer.Chars;
        return buffer.ToString();
    }

    private string Decode(int start, int length) {
        try {
            return StrictUtf8.GetString(_data, start, length);
        } catch (DecoderFallbackException exception) {
            throw new System.Xml.XmlException("The Open XML part is not valid UTF-8.", exception);
        }
    }

    // Applies XML end-of-line handling and, for attributes, value normalization (XML 1.0 sections 2.11 and 3.3.3),
    // and expands the predefined entities and character references.
    private void AppendDecoded(OfficeCharBuffer target, int start, int length, bool attribute) {
        int end = start + length;
        int run = start;
        int position = start;
        byte[] data = _data;
        while (position < end) {
            byte value = data[position];
            if (value == (byte)'&') {
                target.AppendUtf8(data, run, position - run);
                position = AppendReference(target, position + 1, end);
                run = position;
            } else if (value == (byte)'\r' || attribute && (value == (byte)'\n' || value == (byte)'\t')) {
                target.AppendUtf8(data, run, position - run);
                if (value == (byte)'\r' && position + 1 < end && data[position + 1] == (byte)'\n') position++;
                target.Append(attribute ? ' ' : '\n');
                position++;
                run = position;
            } else {
                position++;
            }
        }
        target.AppendUtf8(data, run, end - run);
    }

    private void AppendNormalizedLineEnds(OfficeCharBuffer target, int start, int length) {
        int end = start + length;
        int run = start;
        for (int position = start; position < end; position++) {
            if (_data[position] != (byte)'\r') continue;
            target.AppendUtf8(_data, run, position - run);
            target.Append('\n');
            if (position + 1 < end && _data[position + 1] == (byte)'\n') position++;
            run = position + 1;
        }
        target.AppendUtf8(_data, run, end - run);
    }

    private int AppendReference(OfficeCharBuffer target, int start, int end) {
        int semicolon = start;
        // Numeric references may carry any number of leading zeros, so only the value is bounded, not the spelling.
        while (semicolon < end && _data[semicolon] != (byte)';') semicolon++;
        if (semicolon >= end || _data[semicolon] != (byte)';') throw Malformed("A reference is not terminated.");
        int length = semicolon - start;
        if (length >= 2 && _data[start] == (byte)'#') {
            int codePoint = 0;
            bool hex = _data[start + 1] == (byte)'x';
            for (int index = start + (hex ? 2 : 1); index < semicolon; index++) {
                int digit = HexOrDecimalDigit(_data[index], hex);
                if (digit < 0 || codePoint > 0x10FFFF) throw Malformed("A character reference is not a number.");
                codePoint = codePoint * (hex ? 16 : 10) + digit;
            }
            if (hex && length == 2 || !IsXmlCharacter(codePoint)) throw Malformed("A character reference names an invalid character.");
            if (codePoint < 0x10000) {
                target.Append((char)codePoint);
            } else {
                codePoint -= 0x10000;
                target.Append((char)(0xD800 + (codePoint >> 10)));
                target.Append((char)(0xDC00 + (codePoint & 0x3FF)));
            }
        } else if (AsciiEquals(start, length, "lt")) {
            target.Append('<');
        } else if (AsciiEquals(start, length, "gt")) {
            target.Append('>');
        } else if (AsciiEquals(start, length, "amp")) {
            target.Append('&');
        } else if (AsciiEquals(start, length, "quot")) {
            target.Append('"');
        } else if (AsciiEquals(start, length, "apos")) {
            target.Append('\'');
        } else {
            // Without a DTD only the predefined entities exist.
            throw Malformed("The entity '" + Encoding.UTF8.GetString(_data, start, length) + "' is not defined.");
        }
        return semicolon + 1;
    }

    private static int HexOrDecimalDigit(byte value, bool hex) {
        if (value >= (byte)'0' && value <= (byte)'9') return value - '0';
        if (!hex) return -1;
        if (value >= (byte)'a' && value <= (byte)'f') return value - 'a' + 10;
        if (value >= (byte)'A' && value <= (byte)'F') return value - 'A' + 10;
        return -1;
    }

    private static bool IsXmlCharacter(int codePoint) =>
        codePoint == 0x9 || codePoint == 0xA || codePoint == 0xD
        || codePoint >= 0x20 && codePoint <= 0xD7FF
        || codePoint >= 0xE000 && codePoint <= 0xFFFD
        || codePoint >= 0x10000 && codePoint <= 0x10FFFF;

    private static bool IsXmlSpace(byte value) => value == (byte)' ' || value == (byte)'\n' || value == (byte)'\t' || value == (byte)'\r';

    private static bool IsNameTerminator(byte value) =>
        value == (byte)'>' || value == (byte)'/' || value == (byte)' ' || value == (byte)'\n' || value == (byte)'\t' || value == (byte)'\r';

    private bool AsciiEquals(int start, int length, string name) {
        if (length != name.Length) return false;
        byte[] data = _data;
        for (int index = 0; index < length; index++) {
            if (data[start + index] != name[index]) return false;
        }
        return true;
    }

    private bool Utf8Equals(int start, int length, byte[] value) {
        if (length != value.Length) return false;
        for (int index = 0; index < length; index++) {
            if (_data[start + index] != value[index]) return false;
        }
        return true;
    }

    private bool BytesEqual(int first, int second, int length) {
        byte[] data = _data;
        for (int index = 0; index < length; index++) {
            if (data[first + index] != data[second + index]) return false;
        }
        return true;
    }

    private bool HasCarriageReturn(int start, int length) {
        for (int index = start; index < start + length; index++) {
            if (_data[index] == (byte)'\r') return true;
        }
        return false;
    }

    private bool StartsWith(int position, string value) => position + value.Length <= _end && AsciiEquals(position, value.Length, value);

    private int IndexOf(int start, string value) {
        for (int position = start; position <= _end - value.Length; position++) {
            if (_data[position] == value[0] && AsciiEquals(position, value.Length, value)) return position;
        }
        return -1;
    }

    private void SkipComment() {
        int start = _position + 4;
        int end = IndexOf(start, "--");
        if (end < 0 || end + 2 >= _end || _data[end + 2] != (byte)'>') {
            throw Malformed("A comment contains an invalid double hyphen or is not closed.");
        }
        ValidateLiteralCharacters(start, end);
        _position = end + 3;
    }

    private void ValidateLiteralCharacters(int start, int end) {
        for (int index = start; index < end; index++) {
            byte value = _data[index];
            if (value < 0x20 && value != (byte)'\t' && value != (byte)'\n' && value != (byte)'\r' || value == 0xEF) {
                RejectNonCharacter(index);
            }
        }
    }

    private void SkipPast(string terminator, int offset) {
        int index = IndexOf(_position + offset, terminator);
        if (index < 0) throw Malformed("Markup is not closed.");
        _position = index + terminator.Length;
    }
}

/// <summary>A growable character buffer that text readers reuse between strings.</summary>
internal sealed class OfficeCharBuffer {
    private static readonly UTF8Encoding StrictUtf8 = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);
    private char[] _chars;

    internal OfficeCharBuffer(int capacity = 256) => _chars = new char[capacity];

    internal OfficeCharBuffer(char[] chars) => _chars = chars;

    internal char[] Chars => _chars;

    internal int Length { get; private set; }

    internal void Clear() => Length = 0;

    internal char this[int index] => _chars[index];

    /// <summary>Removes <paramref name="count"/> characters starting at <paramref name="start"/>.</summary>
    internal void Remove(int start, int count) {
        Array.Copy(_chars, start + count, _chars, start, Length - start - count);
        Length -= count;
    }

    /// <summary>Drops characters appended after <paramref name="length"/>.</summary>
    internal void Truncate(int length) => Length = length;

    internal void Append(char value) {
        if (Length == _chars.Length) Grow(1);
        _chars[Length++] = value;
    }

    internal void AppendUtf8(byte[] data, int start, int count) {
        if (count == 0) return;
        // UTF-8 never yields more UTF-16 code units than bytes.
        if (_chars.Length - Length < count) Grow(count);
        try {
            Length += StrictUtf8.GetChars(data, start, count, _chars, Length);
        } catch (DecoderFallbackException exception) {
            throw new System.Xml.XmlException("The Open XML part is not valid UTF-8.", exception);
        }
    }

    private void Grow(int additional) {
        int capacity = Math.Max(_chars.Length * 2, Length + additional);
        Array.Resize(ref _chars, capacity);
    }

    public override string ToString() => Length == 0 ? string.Empty : new string(_chars, 0, Length);
}
