using System.Text;

namespace OfficeIMO.OneNote;

internal static class OneNoteCrc32 {
    private const uint OnePolynomial = 0xEDB88320U;
    private const uint MsoPolynomial = 0x000000AFU;

    internal static uint ComputeFileName(string? fileName) {
        if (string.IsNullOrEmpty(fileName)) return 0;
        return Compute(Encoding.Unicode.GetBytes(fileName + "\0"));
    }

    internal static uint Compute(IEnumerable<uint> values) {
        using (var stream = new MemoryStream()) {
            foreach (uint value in values) FssHttpStreamObjectWriter.WriteUInt32(stream, value);
            return Compute(stream.ToArray());
        }
    }

    internal static uint Compute(byte[] data) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        return Continue(0, data, 0, data.Length, OneNoteFileKind.Section);
    }

    internal static uint ComputeMso(byte[] data) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        return Continue(0, data, 0, data.Length, OneNoteFileKind.TableOfContents);
    }

    internal static uint Continue(uint crc, byte[] data, int offset, int count, OneNoteFileKind fileKind) {
        if (data == null) throw new ArgumentNullException(nameof(data));
        if (offset < 0 || count < 0 || offset > data.Length - count) throw new ArgumentOutOfRangeException(nameof(offset));
        if (fileKind == OneNoteFileKind.Section) return ContinueOne(crc, data, offset, count);
        if (fileKind == OneNoteFileKind.TableOfContents) return ContinueMso(crc, data, offset, count);
        throw new ArgumentOutOfRangeException(nameof(fileKind));
    }

    private static uint ContinueOne(uint crc, byte[] data, int offset, int count) {
        uint state = ~crc;
        int end = offset + count;
        for (int index = offset; index < end; index++) {
            state ^= data[index];
            for (int bit = 0; bit < 8; bit++) state = (state >> 1) ^ ((state & 1) != 0 ? OnePolynomial : 0U);
        }
        return ~state;
    }

    private static uint ContinueMso(uint crc, byte[] data, int offset, int count) {
        int end = offset + count;
        for (int index = offset; index < end; index++) {
            crc ^= (uint)data[index] << 24;
            for (int bit = 0; bit < 8; bit++) crc = (crc << 1) ^ ((crc & 0x80000000U) != 0 ? MsoPolynomial : 0U);
        }
        return crc;
    }
}
