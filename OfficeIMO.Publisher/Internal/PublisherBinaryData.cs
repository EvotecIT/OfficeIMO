using System.Text;

namespace OfficeIMO.Publisher.Internal;

internal sealed class PublisherBinaryData {
    internal PublisherBinaryData(byte[] bytes, string name) { Bytes = bytes; Name = name; }
    internal byte[] Bytes { get; }
    internal string Name { get; }
    internal int Length => Bytes.Length;
    internal void Range(int offset, int length, int? boundary = null) {
        int end = boundary ?? Length;
        if (end < 0 || end > Length || offset < 0 || length < 0 || offset > end - length)
            throw new InvalidDataException($"Truncated or out-of-range Publisher record in {Name} at 0x{offset:X}.");
    }
    internal byte U8(int offset) { Range(offset, 1); return Bytes[offset]; }
    internal ushort U16(int offset) { Range(offset, 2); return OfficeLegacyImportBuffer.ReadUInt16(Bytes, offset); }
    internal uint U32(int offset) { Range(offset, 4); return unchecked((uint)OfficeLegacyImportBuffer.ReadInt32(Bytes, offset)); }
    internal int I32(int offset) => unchecked((int)U32(offset));
    internal int Offset(uint value) {
        if (value > int.MaxValue || value > Length) throw new InvalidDataException($"Publisher offset exceeds {Name} bounds.");
        return (int)value;
    }
    internal string Utf16(int offset, int bytes) {
        Range(offset, bytes);
        if ((bytes & 1) != 0) throw new InvalidDataException($"Odd Publisher UTF-16 byte count in {Name}.");
        try { return new UnicodeEncoding(false, false, true).GetString(Bytes, offset, bytes); }
        catch (DecoderFallbackException exception) { throw new InvalidDataException($"Invalid Publisher UTF-16 in {Name}.", exception); }
    }
    internal string Tag(int offset) { Range(offset, 4); return Encoding.ASCII.GetString(Bytes, offset, 4); }
}
