using AngleSharp.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Html;

/// <summary>Canonical bounded-prefix encoding resolution for HTML, CSS, and textual data URIs.</summary>
internal static class HtmlTextEncodingResolver {
    private const int HtmlPrescanLength = 1024;
    private const int CssSniffLength = 4096;
    private static readonly Encoding Utf8 = new UTF8Encoding(false);
    private static readonly Regex ContentTypeCharset = new(
        "(?:^|;)\\s*charset\\s*=\\s*['\\\"]?\\s*([^\\s;'\\\"]+)",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
    private static readonly Regex CssCharset = new(
        "^@charset\\s+['\\\"]([^'\\\"]+)['\\\"];",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    static HtmlTextEncodingResolver() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    internal static Encoding ResolveHtmlEncoding(Stream stream, Encoding? explicitEncoding = null) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (explicitEncoding != null) return explicitEncoding;
        if (!stream.CanSeek) return Utf8;

        long position = stream.Position;
        try {
            stream.Position = 0;
            var prefix = new byte[HtmlPrescanLength];
            int count = ReadPrefix(stream, prefix);
            return ResolveHtmlEncoding(prefix, count) ?? Utf8;
        } finally {
            stream.Position = position;
        }
    }

    internal static Stream PrepareHtmlStream(Stream stream, Encoding? explicitEncoding, out Encoding encoding) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (explicitEncoding != null || stream.CanSeek) {
            encoding = ResolveHtmlEncoding(stream, explicitEncoding);
            return stream;
        }

        var prefix = new byte[HtmlPrescanLength];
        int count = ReadPrefix(stream, prefix);
        encoding = ResolveHtmlEncoding(prefix, count) ?? Utf8;
        return new PrefixReplayStream(prefix, count, stream);
    }

    internal static async Task<(Stream Stream, Encoding Encoding)> PrepareHtmlStreamAsync(
        Stream stream,
        Encoding? explicitEncoding,
        CancellationToken cancellationToken) {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (explicitEncoding != null) return (stream, explicitEncoding);
        if (stream.CanSeek) {
            long position = stream.Position;
            try {
                stream.Position = 0;
                var seekablePrefix = new byte[HtmlPrescanLength];
                int seekableCount = await ReadPrefixAsync(stream, seekablePrefix, cancellationToken).ConfigureAwait(false);
                return (stream, ResolveHtmlEncoding(seekablePrefix, seekableCount) ?? Utf8);
            } finally {
                stream.Position = position;
            }
        }

        var prefix = new byte[HtmlPrescanLength];
        int count = await ReadPrefixAsync(stream, prefix, cancellationToken).ConfigureAwait(false);
        Encoding encoding = ResolveHtmlEncoding(prefix, count) ?? Utf8;
        return (new PrefixReplayStream(prefix, count, stream), encoding);
    }

    internal static Encoding ResolveDataUriEncoding(string metadata) {
        string? charset = ReadCharset(ContentTypeCharset, metadata);
        return charset == null ? Utf8 : GetEncoding(charset);
    }

    internal static bool TryDecodeCss(byte[] bytes, string? contentType, out string css) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        try {
            Encoding encoding = ResolveBomEncoding(bytes)
                ?? ResolveCharsetEncoding(ContentTypeCharset, contentType)
                ?? ResolveCharsetEncoding(CssCharset, GetAsciiPrefix(bytes))
                ?? new UTF8Encoding(false, true);
            int preambleLength = GetPreambleLength(bytes, encoding);
            css = encoding.GetString(bytes, preambleLength, bytes.Length - preambleLength);
            return true;
        } catch (DecoderFallbackException) {
            css = string.Empty;
            return false;
        } catch (ArgumentException) {
            css = string.Empty;
            return false;
        } catch (NotSupportedException) {
            css = string.Empty;
            return false;
        }
    }

    internal static string DecodeCss(byte[] bytes, string? contentType = null) {
        if (TryDecodeCss(bytes, contentType, out string css)) return css;
        throw new DecoderFallbackException("The stylesheet encoding is unsupported or its byte sequence is invalid.");
    }

    private static Encoding? ResolveHtmlEncoding(byte[] prefix, int count) {
        Encoding? bom = ResolveBomEncoding(prefix, count);
        if (bom != null) return bom;
        return PrescanHtmlEncoding(prefix, count) ?? ResolveXmlDeclarationEncoding(prefix, count);
    }

    private static Encoding? PrescanHtmlEncoding(byte[] bytes, int count) {
        if (StartsWith(bytes, count, 0, 0x3C, 0x00, 0x3F, 0x00, 0x78, 0x00)) {
            return new UnicodeEncoding(false, false, true);
        }
        if (StartsWith(bytes, count, 0, 0x00, 0x3C, 0x00, 0x3F, 0x00, 0x78)) {
            return new UnicodeEncoding(true, false, true);
        }

        int position = 0;
        while (position < count) {
            if (StartsWith(bytes, count, position, 0x3C, 0x21, 0x2D, 0x2D)) {
                int commentEnd = FindSequence(bytes, position + 4, count, 0x2D, 0x2D, 0x3E);
                if (commentEnd < 0) return null;
                position = commentEnd + 3;
                continue;
            }

            if (IsMetaStart(bytes, count, position)) {
                position += 5;
                Encoding? encoding = PrescanMetaAttributes(bytes, count, ref position);
                if (encoding != null) return NormalizeHtmlDeclaredEncoding(encoding);
                if (position >= count) return null;
                position++;
                continue;
            }

            if (IsTagStart(bytes, count, position)) {
                position++;
                if (position < count && bytes[position] == 0x2F) position++;
                while (position < count && !IsHtmlSpace(bytes[position]) && bytes[position] != 0x3E) position++;
                while (TryReadPrescanAttribute(bytes, count, ref position, out _, out _)) { }
                if (position >= count) return null;
                position++;
                continue;
            }

            if (StartsWith(bytes, count, position, 0x3C, 0x21)
                || StartsWith(bytes, count, position, 0x3C, 0x2F)
                || StartsWith(bytes, count, position, 0x3C, 0x3F)) {
                int tagEnd = Array.IndexOf(bytes, (byte)0x3E, position + 2, count - position - 2);
                if (tagEnd < 0) return null;
                position = tagEnd + 1;
                continue;
            }

            position++;
        }
        return null;
    }

    private static Encoding? PrescanMetaAttributes(byte[] bytes, int count, ref int position) {
        var attributeNames = new HashSet<string>(StringComparer.Ordinal);
        bool gotPragma = false;
        bool? needPragma = null;
        Encoding? charset = null;
        bool charsetFailed = false;

        while (TryReadPrescanAttribute(bytes, count, ref position, out string name, out string value)) {
            if (!attributeNames.Add(name)) continue;
            if (name == "http-equiv") {
                if (value == "content-type") gotPragma = true;
            } else if (name == "content") {
                Encoding? contentEncoding = TextEncoding.Parse(value);
                if (contentEncoding != null && charset == null && !charsetFailed) {
                    charset = contentEncoding;
                    needPragma = true;
                }
            } else if (name == "charset") {
                charset = ResolveHtmlLabel(value);
                charsetFailed = charset == null;
                needPragma = false;
            }
        }

        if (needPragma == null || needPragma == true && !gotPragma || charsetFailed) return null;
        return charset;
    }

    private static bool TryReadPrescanAttribute(
        byte[] bytes,
        int count,
        ref int position,
        out string name,
        out string value) {
        name = string.Empty;
        value = string.Empty;
        while (position < count && (IsHtmlSpace(bytes[position]) || bytes[position] == 0x2F)) position++;
        if (position >= count || bytes[position] == 0x3E) return false;

        var nameBuilder = new StringBuilder();
        while (position < count) {
            byte current = bytes[position];
            if (current == 0x3D && nameBuilder.Length > 0) {
                position++;
                break;
            }
            if (IsHtmlSpace(current)) {
                while (position < count && IsHtmlSpace(bytes[position])) position++;
                if (position >= count || bytes[position] != 0x3D) {
                    name = nameBuilder.ToString();
                    return name.Length > 0;
                }
                position++;
                break;
            }
            if (current == 0x2F || current == 0x3E) {
                name = nameBuilder.ToString();
                return name.Length > 0;
            }
            nameBuilder.Append((char)ToLowerAscii(current));
            position++;
        }
        if (position > count) return false;

        while (position < count && IsHtmlSpace(bytes[position])) position++;
        if (position >= count) return false;
        name = nameBuilder.ToString();
        if (bytes[position] == 0x3E) return name.Length > 0;

        var valueBuilder = new StringBuilder();
        byte quote = bytes[position];
        if (quote == 0x22 || quote == 0x27) {
            position++;
            while (position < count && bytes[position] != quote) {
                valueBuilder.Append((char)ToLowerAscii(bytes[position]));
                position++;
            }
            if (position >= count) return false;
            position++;
        } else {
            while (position < count && !IsHtmlSpace(bytes[position]) && bytes[position] != 0x3E) {
                valueBuilder.Append((char)ToLowerAscii(bytes[position]));
                position++;
            }
        }
        value = valueBuilder.ToString();
        return name.Length > 0;
    }

    private static Encoding? ResolveXmlDeclarationEncoding(byte[] bytes, int count) {
        if (!StartsWith(bytes, count, 0, 0x3C, 0x3F, 0x78, 0x6D, 0x6C)) return null;
        int declarationEnd = Array.IndexOf(bytes, (byte)0x3E, 5, Math.Max(0, count - 5));
        if (declarationEnd < 0) return null;
        int encodingPosition = FindSequence(bytes, 5, declarationEnd, 0x65, 0x6E, 0x63, 0x6F, 0x64, 0x69, 0x6E, 0x67);
        if (encodingPosition < 0) return null;
        encodingPosition += 8;
        while (encodingPosition < declarationEnd && bytes[encodingPosition] <= 0x20) encodingPosition++;
        if (encodingPosition >= declarationEnd || bytes[encodingPosition] != 0x3D) return null;
        encodingPosition++;
        while (encodingPosition < declarationEnd && bytes[encodingPosition] <= 0x20) encodingPosition++;
        if (encodingPosition >= declarationEnd || bytes[encodingPosition] != 0x22 && bytes[encodingPosition] != 0x27) return null;
        byte quote = bytes[encodingPosition++];
        int encodingEnd = Array.IndexOf(bytes, quote, encodingPosition, declarationEnd - encodingPosition);
        if (encodingEnd < 0) return null;
        for (int index = encodingPosition; index < encodingEnd; index++) {
            if (bytes[index] <= 0x20) return null;
        }
        string label = Encoding.ASCII.GetString(bytes, encodingPosition, encodingEnd - encodingPosition);
        Encoding? encoding = ResolveHtmlLabel(label);
        return encoding == null ? null : NormalizeHtmlDeclaredEncoding(encoding);
    }

    private static Encoding? ResolveHtmlLabel(string label) {
        if (string.Equals(label.Trim(), "x-user-defined", StringComparison.OrdinalIgnoreCase)) {
            return TextEncoding.Resolve("windows-1252");
        }
        return TextEncoding.IsSupported(label) ? TextEncoding.Resolve(label) : null;
    }

    private static Encoding NormalizeHtmlDeclaredEncoding(Encoding encoding) =>
        IsUtf16(encoding) ? Utf8 : encoding;

    private static bool IsMetaStart(byte[] bytes, int count, int position) =>
        StartsWithAsciiCaseInsensitive(bytes, count, position, "<meta")
        && position + 5 < count
        && (IsHtmlSpace(bytes[position + 5]) || bytes[position + 5] == 0x2F);

    private static bool IsTagStart(byte[] bytes, int count, int position) {
        if (position >= count || bytes[position] != 0x3C) return false;
        int namePosition = position + 1;
        if (namePosition < count && bytes[namePosition] == 0x2F) namePosition++;
        return namePosition < count && IsAsciiAlpha(bytes[namePosition]);
    }

    private static bool StartsWithAsciiCaseInsensitive(byte[] bytes, int count, int position, string value) {
        if (position < 0 || count - position < value.Length) return false;
        for (int index = 0; index < value.Length; index++) {
            if (ToLowerAscii(bytes[position + index]) != (byte)value[index]) return false;
        }
        return true;
    }

    private static bool StartsWith(byte[] bytes, int count, int position, params byte[] value) {
        if (position < 0 || count - position < value.Length) return false;
        for (int index = 0; index < value.Length; index++) {
            if (bytes[position + index] != value[index]) return false;
        }
        return true;
    }

    private static int FindSequence(byte[] bytes, int start, int end, params byte[] value) {
        for (int position = start; position <= end - value.Length; position++) {
            if (StartsWith(bytes, end, position, value)) return position;
        }
        return -1;
    }

    private static bool IsHtmlSpace(byte value) =>
        value == 0x09 || value == 0x0A || value == 0x0C || value == 0x0D || value == 0x20;

    private static bool IsAsciiAlpha(byte value) =>
        value >= 0x41 && value <= 0x5A || value >= 0x61 && value <= 0x7A;

    private static byte ToLowerAscii(byte value) =>
        value >= 0x41 && value <= 0x5A ? (byte)(value + 0x20) : value;

    private static Encoding? ResolveBomEncoding(byte[] bytes, int? count = null) {
        int length = count ?? bytes.Length;
        if (length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF) return new UTF8Encoding(true, true);
        if (length >= 2 && bytes[0] == 0xFF && bytes[1] == 0xFE) return new UnicodeEncoding(false, true, true);
        if (length >= 2 && bytes[0] == 0xFE && bytes[1] == 0xFF) return new UnicodeEncoding(true, true, true);
        return null;
    }

    private static Encoding? ResolveCharsetEncoding(Regex pattern, string? source) {
        string? charset = ReadCharset(pattern, source);
        return charset == null ? null : GetEncoding(charset);
    }

    private static string? ReadCharset(Regex pattern, string? source) {
        if (string.IsNullOrWhiteSpace(source)) return null;
        Match match = pattern.Match(source!);
        return match.Success ? match.Groups[1].Value.Trim() : null;
    }

    private static Encoding GetEncoding(string charset) {
        string label = charset.Trim().Trim('\'', '"');
        Encoding? encoding = ResolveHtmlLabel(label);
        if (encoding == null) throw new ArgumentException($"Unsupported character encoding label '{label}'.", nameof(charset));
        return Encoding.GetEncoding(
            encoding.CodePage,
            EncoderFallback.ExceptionFallback,
            DecoderFallback.ExceptionFallback);
    }

    private static string GetAsciiPrefix(byte[] bytes) =>
        Encoding.ASCII.GetString(bytes, 0, Math.Min(bytes.Length, CssSniffLength));

    private static bool IsUtf16(Encoding encoding) => encoding.CodePage == 1200 || encoding.CodePage == 1201;

    private static int GetPreambleLength(byte[] bytes, Encoding encoding) {
        byte[] preamble = encoding.GetPreamble();
        if (preamble.Length == 0 || bytes.Length < preamble.Length) return 0;
        for (int index = 0; index < preamble.Length; index++) {
            if (bytes[index] != preamble[index]) return 0;
        }
        return preamble.Length;
    }

    private static int ReadPrefix(Stream stream, byte[] prefix) {
        int count = 0;
        while (count < prefix.Length) {
            int read = stream.Read(prefix, count, prefix.Length - count);
            if (read == 0) break;
            count += read;
        }
        return count;
    }

    private static async Task<int> ReadPrefixAsync(Stream stream, byte[] prefix, CancellationToken cancellationToken) {
        int count = 0;
        while (count < prefix.Length) {
            int read = await stream.ReadAsync(prefix, count, prefix.Length - count, cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            count += read;
        }
        return count;
    }

    private sealed class PrefixReplayStream : Stream {
        private readonly byte[] _prefix;
        private readonly int _prefixLength;
        private readonly Stream _remainder;
        private int _prefixPosition;

        internal PrefixReplayStream(byte[] prefix, int prefixLength, Stream remainder) {
            _prefix = prefix;
            _prefixLength = prefixLength;
            _remainder = remainder;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) {
            ValidateReadArguments(buffer, offset, count);
            int copied = CopyPrefix(buffer, offset, count);
            return copied == count ? copied : copied + _remainder.Read(buffer, offset + copied, count - copied);
        }

        public override async Task<int> ReadAsync(
            byte[] buffer,
            int offset,
            int count,
            CancellationToken cancellationToken) {
            ValidateReadArguments(buffer, offset, count);
            int copied = CopyPrefix(buffer, offset, count);
            if (copied == count) return copied;
            int read = await _remainder.ReadAsync(
                buffer,
                offset + copied,
                count - copied,
                cancellationToken).ConfigureAwait(false);
            return copied + read;
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        private int CopyPrefix(byte[] buffer, int offset, int count) {
            int available = _prefixLength - _prefixPosition;
            int copied = Math.Min(available, count);
            if (copied > 0) {
                Buffer.BlockCopy(_prefix, _prefixPosition, buffer, offset, copied);
                _prefixPosition += copied;
            }
            return copied;
        }

        private static void ValidateReadArguments(byte[] buffer, int offset, int count) {
            if (buffer == null) throw new ArgumentNullException(nameof(buffer));
            if (offset < 0) throw new ArgumentOutOfRangeException(nameof(offset));
            if (count < 0) throw new ArgumentOutOfRangeException(nameof(count));
            if (buffer.Length - offset < count) throw new ArgumentException("The offset and count exceed the buffer length.");
        }
    }
}
