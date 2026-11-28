namespace OfficeIMO.Core.Internal.OpenXml;

/// <summary>
/// Buffered UTF-8 output for XML parts, with the namespace scope of the element being written. One writer serves every
/// part of a save: <see cref="Target"/> moves to the next part's entry stream.
/// </summary>
internal sealed class OfficeXmlUtf8Writer {
    private readonly byte[] _buffer;
    private int _count;
    private Stream? _target;
    private Binding[] _scope = new Binding[16];
    private int _scopeCount;

    internal OfficeXmlUtf8Writer(int bufferLength = 16 * 1024) => _buffer = new byte[Math.Max(64, bufferLength)];

    /// <summary>Where <see cref="Flush"/> sends the bytes.</summary>
    internal Stream? Target {
        get => _target;
        set {
            if (_count > 0) Flush();
            _target = value;
        }
    }

    internal void Flush() {
        if (_count == 0) return;
        if (_target == null) throw new InvalidOperationException("The writer has no target.");
        _target.Write(_buffer, 0, _count);
        _count = 0;
    }

    internal void WriteByte(byte value) {
        if (_count == _buffer.Length) Flush();
        _buffer[_count++] = value;
    }

    internal void WriteRaw(byte[] data, int start, int length) {
        if (length <= 0) return;
        if (length > _buffer.Length - _count) {
            Flush();
            if (length >= _buffer.Length) {
                _target!.Write(data, start, length);
                return;
            }
        }
        Buffer.BlockCopy(data, start, _buffer, _count, length);
        _count += length;
    }

    /// <summary>Writes <c>prefix:local</c>, or the local name alone for an empty prefix.</summary>
    internal void WriteQualifiedName(string prefix, byte[] localName) {
        if (prefix.Length > 0) {
            byte[] prefixBytes = PrefixBytes(prefix);
            WriteRaw(prefixBytes, 0, prefixBytes.Length);
            WriteByte((byte)':');
        }
        WriteRaw(localName, 0, localName.Length);
    }

    private readonly string?[] _prefixes = new string?[8];
    private readonly byte[]?[] _prefixBytes = new byte[]?[8];
    private int _nextPrefix;

    // A part uses a handful of prefixes, held by reference, so a short list finds them without hashing.
    private byte[] PrefixBytes(string prefix) {
        for (int index = 0; index < _prefixes.Length; index++) {
            if (ReferenceEquals(_prefixes[index], prefix)) return _prefixBytes[index]!;
        }
        byte[] bytes = System.Text.Encoding.UTF8.GetBytes(prefix);
        _prefixes[_nextPrefix] = prefix;
        _prefixBytes[_nextPrefix] = bytes;
        _nextPrefix = (_nextPrefix + 1) % _prefixes.Length;
        return bytes;
    }

    internal void WriteAscii(string text) {
        for (int index = 0; index < text.Length; index++) WriteByte((byte)text[index]);
    }

    /// <summary>
    /// Writes text as UTF-8 with the escapes XML needs to read it back unchanged: <c>&amp;</c>, <c>&lt;</c> and
    /// <c>&gt;</c>, carriage returns as references, and in attributes also quotes, tabs and line feeds.
    /// </summary>
    internal void WriteEscaped(string text, bool attribute) {
        // Most values are short ASCII with nothing to escape: copy them in one pass.
        if (text.Length <= _buffer.Length) {
            if (_buffer.Length - _count < text.Length) Flush();
            int start = _count;
            int position = 0;
            for (; position < text.Length; position++) {
                char value = text[position];
                if (value >= 0x80 || value < 0x20 || value == '&' || value == '<' || value == '>' || value == '"') break;
                _buffer[start + position] = (byte)value;
            }
            _count = start + position;
            if (position == text.Length) return;
            WriteEscapedFrom(text, position, attribute);
            return;
        }
        WriteEscapedFrom(text, 0, attribute);
    }

    private void WriteEscapedFrom(string text, int from, bool attribute) {
        for (int index = from; index < text.Length; index++) {
            char value = text[index];
            if (value < 0x80) {
                switch (value) {
                    case '&': WriteAscii("&amp;"); continue;
                    case '<': WriteAscii("&lt;"); continue;
                    case '>': WriteAscii("&gt;"); continue;
                    case '\r': WriteAscii("&#xD;"); continue;
                    case '"' when attribute: WriteAscii("&quot;"); continue;
                    case '\n' when attribute: WriteAscii("&#xA;"); continue;
                    case '\t' when attribute: WriteAscii("&#x9;"); continue;
                }
                if (value < 0x20 && value != '\n' && value != '\t') {
                    throw new ArgumentException("The text contains a character XML does not allow.", nameof(text));
                }
                WriteByte((byte)value);
                continue;
            }
            if (_buffer.Length - _count < 4) Flush();
            if (value < 0x800) {
                _buffer[_count++] = (byte)(0xC0 | value >> 6);
                _buffer[_count++] = (byte)(0x80 | value & 0x3F);
            } else if (char.IsHighSurrogate(value) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) {
                int codePoint = char.ConvertToUtf32(value, text[++index]);
                _buffer[_count++] = (byte)(0xF0 | codePoint >> 18);
                _buffer[_count++] = (byte)(0x80 | codePoint >> 12 & 0x3F);
                _buffer[_count++] = (byte)(0x80 | codePoint >> 6 & 0x3F);
                _buffer[_count++] = (byte)(0x80 | codePoint & 0x3F);
            } else if (char.IsSurrogate(value) || value >= 0xFFFE) {
                throw new ArgumentException("The text contains a character XML does not allow.", nameof(text));
            } else {
                _buffer[_count++] = (byte)(0xE0 | value >> 12);
                _buffer[_count++] = (byte)(0x80 | value >> 6 & 0x3F);
                _buffer[_count++] = (byte)(0x80 | value & 0x3F);
            }
        }
    }

    // ---- Namespace scope --------------------------------------------------------------------------------------

    internal int ScopeMark => _scopeCount;

    internal void ResetScope() => _scopeCount = 0;

    internal void PopScope(int mark) => _scopeCount = mark;

    internal void PushScope(string prefix, int namespaceId, int uriId) {
        if (_scopeCount == _scope.Length) Array.Resize(ref _scope, _scope.Length * 2);
        _scope[_scopeCount++] = new Binding(prefix, namespaceId, uriId);
    }

    /// <summary>The innermost prefix bound to a namespace and not hidden by an inner binding of the same prefix.</summary>
    internal bool TryFindPrefix(int namespaceId, bool allowDefault, out string prefix) {
        for (int index = _scopeCount - 1; index >= 0; index--) {
            ref Binding binding = ref _scope[index];
            if (binding.NamespaceId != namespaceId) continue;
            if (binding.Prefix.Length == 0 && !allowDefault) continue;
            if (IsHidden(binding.Prefix, index)) continue;
            prefix = binding.Prefix;
            return true;
        }
        prefix = string.Empty;
        return false;
    }

    /// <summary>The URI id of the innermost binding of <paramref name="prefix"/>, if it is bound.</summary>
    internal bool TryGetBinding(string prefix, out int uriId) {
        for (int index = _scopeCount - 1; index >= 0; index--) {
            if (string.Equals(_scope[index].Prefix, prefix, StringComparison.Ordinal)) {
                uriId = _scope[index].UriId;
                return true;
            }
        }
        uriId = OfficeXmlNamespaces.None;
        return false;
    }

    internal bool IsPrefixBound(string prefix) {
        for (int index = _scopeCount - 1; index >= 0; index--) {
            if (string.Equals(_scope[index].Prefix, prefix, StringComparison.Ordinal)) return true;
        }
        return prefix.Length == 0 ? false : prefix == "xml" || prefix == "xmlns";
    }

    private bool IsHidden(string prefix, int below) {
        for (int index = _scopeCount - 1; index > below; index--) {
            if (string.Equals(_scope[index].Prefix, prefix, StringComparison.Ordinal)) return true;
        }
        return false;
    }

    private readonly struct Binding {
        internal Binding(string prefix, int namespaceId, int uriId) {
            Prefix = prefix;
            NamespaceId = namespaceId;
            UriId = uriId;
        }

        internal string Prefix { get; }
        internal int NamespaceId { get; }
        internal int UriId { get; }
    }
}
