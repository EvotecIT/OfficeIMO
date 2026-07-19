using System.Globalization;

namespace OfficeIMO.Tests.Pdf;

/// <summary>Small test-only DER fixture encoder used to construct malformed and edge-case CMS inputs.</summary>
internal static class PdfDerCodec {
    internal static byte[] Sequence(params byte[][] values) => Wrap(0x30, Concatenate(values));

    internal static byte[] Set(params byte[][] values) {
        Array.Sort(values, CompareEncoded);
        return Wrap(0x31, Concatenate(values));
    }

    internal static byte[] Context(int number, byte[] content) => Wrap((byte)(0xA0 + number), content);

    internal static byte[] Integer(int value) => Integer(new[] { (byte)value });

    internal static byte[] Integer(byte[] bigEndian) {
        int offset = 0;
        while (offset + 1 < bigEndian.Length && bigEndian[offset] == 0) offset++;
        int length = bigEndian.Length - offset;
        bool prefix = length == 0 || (bigEndian[offset] & 0x80) != 0;
        var content = new byte[length + (prefix ? 1 : 0)];
        if (length > 0) Buffer.BlockCopy(bigEndian, offset, content, prefix ? 1 : 0, length);
        return Wrap(0x02, content);
    }

    internal static byte[] OctetString(byte[] value) => Wrap(0x04, value);

    internal static byte[] ObjectIdentifier(string value) {
        string[] parts = value.Split('.');
        if (parts.Length < 2 ||
            !uint.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out uint first) ||
            !uint.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out uint second)) {
            throw new ArgumentException("Object identifier is invalid.", nameof(value));
        }
        var content = new List<byte>();
        AppendBase128(content, first * 40U + second);
        for (int index = 2; index < parts.Length; index++) {
            if (!uint.TryParse(parts[index], NumberStyles.None, CultureInfo.InvariantCulture, out uint component)) {
                throw new ArgumentException("Object identifier is invalid.", nameof(value));
            }
            AppendBase128(content, component);
        }
        return Wrap(0x06, content.ToArray());
    }

    internal static byte[] AlgorithmIdentifier(string oid, bool includeNull = true) =>
        includeNull
            ? Sequence(ObjectIdentifier(oid), new byte[] { 0x05, 0x00 })
            : Sequence(ObjectIdentifier(oid));

    internal static byte[] ReplaceTag(byte[] encoded, byte tag) {
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

    private static byte[] Concatenate(params byte[][] values) {
        int length = values.Sum(static value => value.Length);
        var result = new byte[length];
        int offset = 0;
        foreach (byte[] value in values) {
            Buffer.BlockCopy(value, 0, result, offset, value.Length);
            offset += value.Length;
        }
        return result;
    }

    private static byte[] EncodeLength(int length) {
        if (length < 128) return new[] { (byte)length };
        int count = 0;
        int remaining = length;
        while (remaining > 0) {
            count++;
            remaining >>= 8;
        }
        var result = new byte[count + 1];
        result[0] = (byte)(0x80 | count);
        for (int index = count; index > 0; index--) {
            result[index] = (byte)(length & 0xFF);
            length >>= 8;
        }
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
        for (int index = 0; index < count; index++) {
            int comparison = left[index].CompareTo(right[index]);
            if (comparison != 0) return comparison;
        }
        return left.Length.CompareTo(right.Length);
    }
}
