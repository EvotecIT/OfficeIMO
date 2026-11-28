using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Archives.Rar5;

/// <summary>Reads a RAR5 header or a bounded subrecord after its enclosing CRC has been checked.</summary>
internal sealed class Rar5HeaderReader {
    private static readonly UTF8Encoding Utf8 = new(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true);
    private readonly byte[] _bytes;
    private readonly int _end;
    private int _offset;

    internal Rar5HeaderReader(byte[] bytes) : this(bytes, 0, bytes.Length) { }

    private Rar5HeaderReader(byte[] bytes, int offset, int length) {
        _bytes = bytes;
        _offset = offset;
        _end = offset + length;
    }

    internal int Remaining => _end - _offset;

    internal ulong Number() {
        ulong value = 0;
        for (int shift = 0; shift <= 63; shift += 7) {
            int next = _bytes[Take(1)];
            int bits = next & 127;
            if (shift == 63 && next > 1)
                throw new InvalidDataException("A RAR5 variable integer exceeds 64 bits.");
            value |= (ulong)bits << shift;
            if ((next & 128) == 0) return value;
        }
        throw new InvalidDataException("A RAR5 variable integer is too long.");
    }

    internal long Size() {
        ulong value = Number();
        if (value > long.MaxValue) throw new InvalidDataException("A RAR5 size exceeds the supported integer range.");
        return (long)value;
    }

    internal ulong UInt64() {
        uint low = UInt32();
        return low | ((ulong)UInt32() << 32);
    }

    internal byte[] Bytes(long length) {
        int start = Take(length);
        var result = new byte[(int)length];
        Buffer.BlockCopy(_bytes, start, result, 0, result.Length);
        return result;
    }

    internal uint UInt32() => UInt32At(_bytes, Take(4));

    internal Rar5HeaderReader Slice(long size) {
        int start = Take(size);
        return new Rar5HeaderReader(_bytes, start, (int)size);
    }

    internal string ReadName(int characterLimit) {
        long length = Size();
        ArchiveLimits.Check(length, (long)characterLimit * 4, nameof(ArchiveExtractionOptions.MaxPathLength));
        int start = Take(length);
        try { return Utf8.GetString(_bytes, start, (int)length); }
        catch (DecoderFallbackException error) { throw new InvalidDataException("A RAR5 name is not valid UTF-8.", error); }
    }

    internal static string DecodeText(byte[] bytes) {
        try { return Utf8.GetString(bytes); }
        catch (DecoderFallbackException error) { throw new InvalidDataException("RAR5 text is not valid UTF-8.", error); }
    }

    internal void Finish() {
        if (Remaining != 0) throw new InvalidDataException("A RAR5 header contains unexpected trailing fields.");
    }

    private int Take(long count) {
        if (count < 0 || count > Remaining) throw new InvalidDataException("A RAR5 record extends beyond its header.");
        int start = _offset;
        _offset += (int)count;
        return start;
    }

    internal static uint UInt32At(byte[] bytes, int offset) =>
        bytes[offset] | ((uint)bytes[offset + 1] << 8) | ((uint)bytes[offset + 2] << 16) | ((uint)bytes[offset + 3] << 24);

    internal static void Fill(Stream source, byte[] destination, int count, CancellationToken token, int start = 0) {
        int total = 0;
        while (total < count) {
            token.ThrowIfCancellationRequested();
            int read = source.Read(destination, start + total, Math.Min(65536, count - total));
            if (read == 0) throw new InvalidDataException("The RAR5 archive ended before the declared record size.");
            total += read;
        }
    }

    internal static uint Checksum(uint crc, byte[] input, CancellationToken token) {
        for (int offset = 0; offset < input.Length;) {
            token.ThrowIfCancellationRequested();
            int count = Math.Min(65536, input.Length - offset);
            crc = OfficePngCrc32.Append(crc, input, offset, count);
            offset += count;
        }
        return OfficePngCrc32.Complete(crc);
    }
}
