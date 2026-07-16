using System.Text;

namespace OfficeIMO.OneNote.Tests;

public sealed class CabinetReaderTests {
    [Fact]
    public void ReadsBoundedUncompressedCabinetEntry() {
        byte[] payload = Encoding.UTF8.GetBytes("offline OneNote package entry");
        byte[] cabinet = BuildUncompressedCabinet("Notebook\\Section.one", payload);

        OneNoteCabinetEntry entry = Assert.Single(OneNoteCabinetArchiveReader.Read(cabinet, 1024 * 1024, 1024 * 1024, 10));

        Assert.Equal("Notebook\\Section.one", entry.Name);
        Assert.Equal(payload, entry.Data);
    }

    [Fact]
    public void LzxUncompressedBlockRoundTrips() {
        byte[] payload = Encoding.UTF8.GetBytes("uncompressed LZX block payload, raw bytes copied verbatim");
        byte[] chunk = BuildUncompressedLzxChunk(payload);

        byte[] decoded = OneNoteLzxDecoder.Decompress(new[] { chunk }, new[] { payload.Length }, 15, 1024 * 1024);

        Assert.Equal(payload, decoded);
    }

    [Fact]
    public void RejectsCabinetEntryPastConfiguredLimit() {
        byte[] cabinet = BuildUncompressedCabinet("large.one", new byte[128]);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteCabinetArchiveReader.Read(cabinet, 1024, 64, 10));

        Assert.Equal("ONENOTE_CAB_ENTRY_LIMIT", exception.Code);
    }

    [Theory]
    [InlineData("/absolute.one")]
    [InlineData("../escape.one")]
    [InlineData("group/../escape.one")]
    [InlineData("C:\\drive.one")]
    public void RejectsUnsafePackageEntryNames(string name) {
        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() => OneNotePackageReader.NormalizeEntryName(name));

        Assert.Equal("ONENOTE_PACKAGE_ENTRY_PATH", exception.Code);
    }

    private static byte[] BuildUncompressedCabinet(string name, byte[] payload) {
        byte[] nameBytes = Encoding.UTF8.GetBytes(name + "\0");
        const int folderOffset = 36;
        const int filesOffset = folderOffset + 8;
        int dataOffset = filesOffset + 16 + nameBytes.Length;
        int cabinetSize = dataOffset + 8 + payload.Length;
        var data = new byte[cabinetSize];
        data[0] = (byte)'M'; data[1] = (byte)'S'; data[2] = (byte)'C'; data[3] = (byte)'F';
        WriteUInt32(data, 8, (uint)cabinetSize);
        WriteUInt32(data, 16, filesOffset);
        data[24] = 3; data[25] = 1;
        WriteUInt16(data, 26, 1);
        WriteUInt16(data, 28, 1);
        WriteUInt32(data, folderOffset, (uint)dataOffset);
        WriteUInt16(data, folderOffset + 4, 1);
        WriteUInt16(data, folderOffset + 6, 0);
        WriteUInt32(data, filesOffset, (uint)payload.Length);
        WriteUInt32(data, filesOffset + 4, 0);
        WriteUInt16(data, filesOffset + 8, 0);
        WriteUInt16(data, filesOffset + 14, 0x80);
        Buffer.BlockCopy(nameBytes, 0, data, filesOffset + 16, nameBytes.Length);
        WriteUInt16(data, dataOffset + 4, (ushort)payload.Length);
        WriteUInt16(data, dataOffset + 6, (ushort)payload.Length);
        Buffer.BlockCopy(payload, 0, data, dataOffset + 8, payload.Length);
        return data;
    }

    private static byte[] BuildUncompressedLzxChunk(byte[] payload) {
        var bits = new List<int>();
        AppendBits(bits, 0, 1);
        AppendBits(bits, 3, 3);
        AppendBits(bits, payload.Length >> 8, 16);
        AppendBits(bits, payload.Length & 0xFF, 8);
        byte[] header = PackWords(bits);
        var result = new byte[header.Length + 12 + payload.Length];
        Buffer.BlockCopy(header, 0, result, 0, header.Length);
        WriteUInt32(result, header.Length, 1);
        WriteUInt32(result, header.Length + 4, 1);
        WriteUInt32(result, header.Length + 8, 1);
        Buffer.BlockCopy(payload, 0, result, header.Length + 12, payload.Length);
        return result;
    }

    private static void AppendBits(ICollection<int> bits, int value, int count) {
        for (int index = count - 1; index >= 0; index--) bits.Add((value >> index) & 1);
    }

    private static byte[] PackWords(List<int> bits) {
        while (bits.Count % 16 != 0) bits.Add(0);
        var output = new byte[bits.Count / 8];
        int outputOffset = 0;
        for (int offset = 0; offset < bits.Count; offset += 16) {
            ushort word = 0;
            for (int index = 0; index < 16; index++) word = (ushort)((word << 1) | bits[offset + index]);
            output[outputOffset++] = (byte)word;
            output[outputOffset++] = (byte)(word >> 8);
        }
        return output;
    }

    private static void WriteUInt16(byte[] data, int offset, ushort value) {
        data[offset] = (byte)value;
        data[offset + 1] = (byte)(value >> 8);
    }

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)value;
        data[offset + 1] = (byte)(value >> 8);
        data[offset + 2] = (byte)(value >> 16);
        data[offset + 3] = (byte)(value >> 24);
    }
}
