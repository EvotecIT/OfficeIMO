namespace OfficeIMO.OneNote;

internal static class OneNoteBinary {
    public static Guid ReadGuid(byte[] data, int offset) {
        EnsureRange(data, offset, 16);
        var bytes = new byte[16];
        Buffer.BlockCopy(data, offset, bytes, 0, bytes.Length);
        return new Guid(bytes);
    }

    public static ushort ReadUInt16(byte[] data, int offset) {
        EnsureRange(data, offset, 2);
        return (ushort)(data[offset] | (data[offset + 1] << 8));
    }

    public static uint ReadUInt32(byte[] data, int offset) {
        EnsureRange(data, offset, 4);
        return (uint)(data[offset] |
                      (data[offset + 1] << 8) |
                      (data[offset + 2] << 16) |
                      (data[offset + 3] << 24));
    }

    public static ulong ReadUInt64(byte[] data, int offset) {
        EnsureRange(data, offset, 8);
        uint low = ReadUInt32(data, offset);
        uint high = ReadUInt32(data, offset + 4);
        return low | ((ulong)high << 32);
    }

    public static OneNoteFileChunkReference ReadFileChunkReference64x32(byte[] data, int offset) {
        return new OneNoteFileChunkReference(ReadUInt64(data, offset), ReadUInt32(data, offset + 8));
    }

    public static void EnsureRange(byte[] data, int offset, int length) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (offset < 0 || length < 0 || offset > data.Length - length) {
            throw new OneNoteFormatException(
                "ONENOTE_TRUNCATED_STRUCTURE",
                "The OneNote file ended before a required structure could be read.",
                offset);
        }
    }
}
