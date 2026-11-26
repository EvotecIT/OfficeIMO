using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Archives.SevenZip;

/// <summary>Bounded cursor over one 7z metadata record.</summary>
internal sealed class SevenZipInput {
    private static readonly UnicodeEncoding Utf16 = new(bigEndian: false, byteOrderMark: false, throwOnInvalidBytes: true);
    private readonly byte[] _data;
    private readonly int _end;
    private int _offset;

    internal SevenZipInput(byte[] data) : this(data, 0, data.Length) { }

    private SevenZipInput(byte[] data, int offset, int length) {
        _data = data;
        _offset = offset;
        _end = offset + length;
    }

    internal static uint UInt32At(byte[] data, int offset) =>
        data[offset] | ((uint)data[offset + 1] << 8) | ((uint)data[offset + 2] << 16) | ((uint)data[offset + 3] << 24);

    internal int Remaining => _end - _offset;

    internal byte Byte() {
        Need(1);
        return _data[_offset++];
    }

    internal ulong Number() {
        byte first = Byte();
        ulong value = 0;
        int highBit = 0x80;
        for (int shift = 0; shift < 64; shift += 8) {
            if ((first & highBit) == 0) return value | ((ulong)(first & (highBit - 1)) << shift);
            value |= (ulong)Byte() << shift;
            highBit >>= 1;
        }
        return value;
    }

    internal long Size() {
        ulong value = Number();
        if (value > long.MaxValue) throw Invalid("A 7z size or offset exceeds the supported integer range.");
        return (long)value;
    }

    internal int Count(int maximum, string limitName) {
        ulong value = Number();
        if (value > (ulong)maximum) throw ArchiveLimits.Exceeded(limitName, maximum);
        return (int)value;
    }

    internal uint U32() {
        Need(4);
        uint value = UInt32At(_data, _offset);
        _offset += 4;
        return value;
    }

    internal long I64() {
        Need(8);
        long value = unchecked((long)(UInt32At(_data, _offset) | ((ulong)UInt32At(_data, _offset + 4) << 32)));
        _offset += 8;
        return value;
    }

    internal void Expect(byte property) {
        if (Byte() != property) throw Invalid($"Expected 7z field 0x{property:x2}.");
    }

    internal void Finish() {
        if (Remaining != 0) throw Invalid("A 7z metadata record contains trailing bytes.");
    }

    internal SevenZipInput Property() {
        int length = TakeSize(Size());
        var result = new SevenZipInput(_data, _offset, length);
        _offset += length;
        return result;
    }

    internal byte[] Bytes(int length) {
        Need(length);
        var result = new byte[length];
        Buffer.BlockCopy(_data, _offset, result, 0, length);
        _offset += length;
        return result;
    }

    internal void SkipRemaining() => _offset = _end;

    internal bool[] Bits(int count) {
        var result = new bool[count];
        int bits = 0;
        for (int i = 0; i < count; i++) {
            if ((i & 7) == 0) bits = Byte();
            result[i] = (bits & (0x80 >> (i & 7))) != 0;
        }
        return result;
    }

    internal bool[]? Defined(int count) => Byte() switch {
        0 => Bits(count),
        1 => null,
        _ => throw Invalid("A 7z digest or metadata presence flag is invalid.")
    };

    internal uint?[] Checksums(int count) {
        bool[]? defined = Defined(count);
        var result = new uint?[count];
        for (int i = 0; i < count; i++) {
            if (defined == null || defined[i]) result[i] = U32();
        }
        return result;
    }

    internal void RequireInline(string field) {
        if (Byte() != 0) throw new NotSupportedException($"External 7z {field} records are unsupported.");
    }

    internal string Name(int maximum, CancellationToken token) {
        int start = _offset;
        while (true) {
            Need(2);
            if (_data[_offset] == 0 && _data[_offset + 1] == 0) break;
            _offset += 2;
            int chars = (_offset - start) / 2;
            if ((chars & 4095) == 0) token.ThrowIfCancellationRequested();
            ArchiveLimits.Check(chars, maximum, nameof(ArchiveReadOptions.MaxPathLength));
        }
        int length = _offset - start;
        _offset += 2;
        try { return Utf16.GetString(_data, start, length); }
        catch (DecoderFallbackException error) { throw new InvalidDataException("A 7z entry name is invalid UTF-16.", error); }
    }

    private int TakeSize(long size) {
        if (size > Remaining) throw Invalid("A 7z property exceeds its containing header.");
        return (int)size;
    }

    private void Need(int bytes) {
        if (bytes < 0 || bytes > Remaining) throw Invalid("The 7z header is truncated.");
    }

    internal static void Fill(Stream source, byte[] data, int offset, int count, CancellationToken token) {
        while (count > 0) {
            token.ThrowIfCancellationRequested();
            int read = source.Read(data, offset, Math.Min(count, 65536));
            if (read <= 0) throw Invalid("The 7z archive ended before the declared stream size.");
            offset += read;
            count -= read;
        }
    }

    internal static void VerifyChecksum(byte[] data, int offset, int length, uint? expected, string field, CancellationToken token) {
        if (!expected.HasValue) return;
        uint crc = OfficePngCrc32.Begin();
        while (length > 0) {
            token.ThrowIfCancellationRequested();
            int block = Math.Min(length, 65536);
            crc = OfficePngCrc32.Append(crc, data, offset, block);
            length -= block;
            offset += block;
        }
        if (OfficePngCrc32.Complete(crc) != expected.Value) throw Invalid("The 7z " + field + " checksum does not match.");
    }

    internal static InvalidDataException Invalid(string message) => new(message);
}
