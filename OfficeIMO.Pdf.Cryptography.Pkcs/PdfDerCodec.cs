using System.Globalization;
using System.IO;
using System.Text;

#pragma warning disable CA1510, CA1512 // Cross-target guard code supports netstandard2.0 and net472.

namespace OfficeIMO.Pdf.Cryptography;

internal static class PdfDerCodec {
    internal static byte[] Sequence(params byte[][] values) => Wrap(0x30, Concatenate(values));
    internal static byte[] Set(params byte[][] values) {
        Array.Sort(values, CompareEncoded);
        return Wrap(0x31, Concatenate(values));
    }
    internal static byte[] Context(int number, byte[] content) => Wrap((byte)(0xA0 + number), content);
    internal static byte[] Integer(int value) => Integer(new[] { (byte)value });
    internal static byte[] Integer(byte[] bigEndian) {
        if (bigEndian == null) throw new ArgumentNullException(nameof(bigEndian));
        int offset = 0;
        while (offset + 1 < bigEndian.Length && bigEndian[offset] == 0) offset++;
        int length = bigEndian.Length - offset;
        bool prefix = length == 0 || (bigEndian[offset] & 0x80) != 0;
        var content = new byte[length + (prefix ? 1 : 0)];
        if (length > 0) Buffer.BlockCopy(bigEndian, offset, content, prefix ? 1 : 0, length);
        return Wrap(0x02, content);
    }
    internal static byte[] Null() => new byte[] { 0x05, 0x00 };
    internal static byte[] OctetString(byte[] value) => Wrap(0x04, value ?? throw new ArgumentNullException(nameof(value)));
    internal static byte[] UtcTime(DateTimeOffset value) => Wrap(0x17, Encoding.ASCII.GetBytes(value.UtcDateTime.ToString("yyMMddHHmmss'Z'", CultureInfo.InvariantCulture)));
    internal static byte[] ObjectIdentifier(string value) {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Object identifier cannot be empty.", nameof(value));
        string[] parts = value.Split('.');
        if (parts.Length < 2 || !uint.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out uint first) || first > 2 ||
            !uint.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out uint second) || (first < 2 && second > 39)) {
            throw new ArgumentException("Object identifier is invalid.", nameof(value));
        }

        var content = new List<byte>();
        AppendBase128(content, first * 40U + second);
        for (int i = 2; i < parts.Length; i++) {
            if (!uint.TryParse(parts[i], NumberStyles.None, CultureInfo.InvariantCulture, out uint component)) {
                throw new ArgumentException("Object identifier is invalid.", nameof(value));
            }
            AppendBase128(content, component);
        }
        return Wrap(0x06, content.ToArray());
    }
    internal static byte[] AlgorithmIdentifier(string oid, bool includeNull = true) =>
        includeNull ? Sequence(ObjectIdentifier(oid), Null()) : Sequence(ObjectIdentifier(oid));
    internal static byte[] ReplaceTag(byte[] encoded, byte tag) {
        if (encoded == null || encoded.Length < 2) throw new ArgumentException("DER value is incomplete.", nameof(encoded));
        var result = (byte[])encoded.Clone();
        result[0] = tag;
        return result;
    }
    internal static byte[] Wrap(byte tag, byte[] content) {
        byte[] length = EncodeLength(content.Length);
        var result = new byte[1 + length.Length + content.Length];
        result[0] = tag;
        Buffer.BlockCopy(length, 0, result, 1, length.Length);
        Buffer.BlockCopy(content, 0, result, 1 + length.Length, content.Length);
        return result;
    }
    internal static byte[] Concatenate(params byte[][] values) {
        int length = 0;
        for (int i = 0; i < values.Length; i++) checked { length += values[i].Length; }
        var result = new byte[length];
        int offset = 0;
        for (int i = 0; i < values.Length; i++) {
            Buffer.BlockCopy(values[i], 0, result, offset, values[i].Length);
            offset += values[i].Length;
        }
        return result;
    }

    private static byte[] EncodeLength(int length) {
        if (length < 0) throw new ArgumentOutOfRangeException(nameof(length));
        if (length < 128) return new[] { (byte)length };
        int count = 0;
        int remaining = length;
        while (remaining > 0) { count++; remaining >>= 8; }
        var result = new byte[count + 1];
        result[0] = (byte)(0x80 | count);
        for (int i = count; i > 0; i--) { result[i] = (byte)(length & 0xFF); length >>= 8; }
        return result;
    }
    private static void AppendBase128(List<byte> output, uint value) {
        var buffer = new byte[5];
        int index = buffer.Length;
        buffer[--index] = (byte)(value & 0x7F);
        while ((value >>= 7) > 0) buffer[--index] = (byte)(0x80 | (value & 0x7F));
        for (; index < buffer.Length; index++) output.Add(buffer[index]);
    }
    private static int CompareEncoded(byte[] left, byte[] right) {
        int count = Math.Min(left.Length, right.Length);
        for (int i = 0; i < count; i++) {
            int comparison = left[i].CompareTo(right[i]);
            if (comparison != 0) return comparison;
        }
        return left.Length.CompareTo(right.Length);
    }
}

internal readonly struct PdfDerElement {
    internal PdfDerElement(byte[] source, int offset, int headerLength, int contentLength) {
        Source = source;
        Offset = offset;
        HeaderLength = headerLength;
        ContentLength = contentLength;
    }
    internal byte[] Source { get; }
    internal int Offset { get; }
    internal byte Tag => Source[Offset];
    internal int HeaderLength { get; }
    internal int ContentLength { get; }
    internal int TotalLength => HeaderLength + ContentLength;
    internal int ContentOffset => Offset + HeaderLength;
    internal byte[] Encoded() => Slice(Source, Offset, TotalLength);
    internal byte[] Content() => Slice(Source, ContentOffset, ContentLength);
    internal PdfDerReader Reader() => new PdfDerReader(Source, ContentOffset, ContentLength);

    private static byte[] Slice(byte[] source, int offset, int length) {
        var result = new byte[length];
        Buffer.BlockCopy(source, offset, result, 0, length);
        return result;
    }
}

internal sealed class PdfDerReader {
    private readonly byte[] _source;
    private readonly int _end;
    private int _offset;

    internal PdfDerReader(byte[] source) : this(source, 0, source?.Length ?? 0) { }
    internal PdfDerReader(byte[] source, int offset, int length) {
        _source = source ?? throw new ArgumentNullException(nameof(source));
        if (offset < 0 || length < 0 || offset > source.Length - length) throw new ArgumentOutOfRangeException(nameof(offset));
        _offset = offset;
        _end = offset + length;
    }
    internal bool HasData => _offset < _end;
    internal int Remaining => _end - _offset;
    internal PdfDerElement Read(byte? expectedTag = null) {
        if (_offset >= _end) throw new InvalidDataException("DER value ended unexpectedly.");
        int start = _offset++;
        byte tag = _source[start];
        if ((tag & 0x1F) == 0x1F) throw new InvalidDataException("High-tag-number DER values are not supported.");
        if (_offset >= _end) throw new InvalidDataException("DER length is missing.");
        int firstLength = _source[_offset++];
        int contentLength;
        if ((firstLength & 0x80) == 0) {
            contentLength = firstLength;
        } else {
            int count = firstLength & 0x7F;
            if (count == 0 || count > 4 || _offset > _end - count) throw new InvalidDataException("DER length is invalid.");
            contentLength = 0;
            for (int i = 0; i < count; i++) {
                if (contentLength > (int.MaxValue >> 8)) throw new InvalidDataException("DER value is too large.");
                contentLength = (contentLength << 8) | _source[_offset++];
            }
            if (contentLength < 128) throw new InvalidDataException("DER length is not minimally encoded.");
        }
        if (contentLength < 0 || _offset > _end - contentLength) throw new InvalidDataException("DER content is truncated.");
        int headerLength = _offset - start;
        var element = new PdfDerElement(_source, start, headerLength, contentLength);
        _offset += contentLength;
        if (expectedTag.HasValue && tag != expectedTag.Value) throw new InvalidDataException("Unexpected DER tag 0x" + tag.ToString("X2", CultureInfo.InvariantCulture) + ".");
        return element;
    }
    internal void EnsureEnd() {
        if (HasData) throw new InvalidDataException("DER value contains unexpected trailing data.");
    }
}
