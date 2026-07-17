using System.Text;
using OfficeIMO.OneNote.Markdown;

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
    public void LzxUncompressedBlocksUsePerBlockPadding() {
        byte[][] blocks = {
            new byte[] { 1, 2, 3 },
            new byte[] { 4, 5 },
            new byte[] { 6 }
        };
        byte[] chunk = BuildUncompressedLzxChunk(blocks);

        byte[] decoded = OneNoteLzxDecoder.Decompress(
            new[] { chunk },
            new[] { blocks.Sum(block => block.Length) },
            15,
            1024 * 1024);

        Assert.Equal(blocks.SelectMany(block => block).ToArray(), decoded);
    }

    [Theory]
    [InlineData("makecab-lzx-testOneNote2016.cab", "testOneNote2016.one")]
    [InlineData("makecab-lzx-testOneNoteFromOffice365-2.cab", "testOneNoteFromOffice365-2.one")]
    public void MicrosoftMakeCabLzxArchivesRoundTripPublicFixtures(string cabinetName, string sourceName) {
        byte[] cabinet = File.ReadAllBytes(FixturePath(cabinetName));
        byte[] expected = File.ReadAllBytes(FixturePath(sourceName));

        OneNoteCabinetEntry entry = Assert.Single(
            OneNoteCabinetArchiveReader.Read(cabinet, 1024 * 1024, 1024 * 1024, 10));

        Assert.Equal(sourceName, entry.Name);
        Assert.Equal(expected, entry.Data);
        Assert.NotEmpty(OneNoteSectionReader.Read(new MemoryStream(entry.Data, writable: false)).Pages);
    }

    [Fact]
    public void RejectsMicrosoftMakeCabBlockWithInvalidChecksum() {
        byte[] cabinet = File.ReadAllBytes(FixturePath("makecab-lzx-testOneNote2016.cab"));
        int dataOffset = checked((int)BitConverter.ToUInt32(cabinet, 36));
        cabinet[dataOffset + 8] ^= 0x01;

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteCabinetArchiveReader.Read(cabinet, 1024 * 1024, 1024 * 1024, 10));

        Assert.Equal("ONENOTE_CAB_CHECKSUM", exception.Code);
    }

    [Theory]
    [InlineData("makecab-lzx15-e8.cab")]
    [InlineData("makecab-lzx16-e8.cab")]
    [InlineData("makecab-lzx17-e8.cab")]
    [InlineData("makecab-lzx18-e8.cab")]
    [InlineData("makecab-lzx19-e8.cab")]
    [InlineData("makecab-lzx20-e8.cab")]
    [InlineData("makecab-lzx-e8.cab")]
    public void MicrosoftMakeCabLzxReversesE8TranslationAcrossWindowSizes(string cabinetName) {
        byte[] cabinet = File.ReadAllBytes(FixturePath(cabinetName));

        OneNoteCabinetEntry entry = Assert.Single(
            OneNoteCabinetArchiveReader.Read(cabinet, 1024 * 1024, 1024 * 1024, 10));

        Assert.Equal("officeimo-lzx-e8.bin", entry.Name);
        Assert.Equal(CreateE8OraclePayload(), entry.Data);
    }

    [Fact]
    public void OneNotePackageReaderOpensCompleteMicrosoftMakeCabLzxNotebook() {
        OneNoteNotebook notebook = OneNotePackageReader.Read(FixturePath("makecab-lzx-notebook.onepkg"));

        OneNoteSection section = Assert.Single(notebook.Sections);
        Assert.Equal("Compressed section", section.Name);
        OneNotePage page = Assert.Single(section.Pages);
        Assert.Equal("Compressed page", page.Title);
        Assert.Contains("LZX package content", OneNoteMarkdownProjection.ToText(page), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(14)]
    [InlineData(22)]
    public void LzxRejectsWindowOutsideCabinetRange(int windowBits) {
        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteLzxDecoder.Decompress(new[] { Array.Empty<byte>() }, new[] { 0 }, windowBits, 1024));

        Assert.Equal("ONENOTE_CAB_LZX_WINDOW", exception.Code);
    }

    [Fact]
    public void LzxCompressedStreamRejectsTruncatedTreeOrTokenData() {
        byte[] cabinet = File.ReadAllBytes(FixturePath("makecab-lzx-testOneNote2016.cab"));
        int dataOffset = checked((int)BitConverter.ToUInt32(cabinet, 36));
        ushort compressedLength = BitConverter.ToUInt16(cabinet, dataOffset + 4);
        WriteUInt16(cabinet, dataOffset + 4, checked((ushort)(compressedLength / 2)));

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteCabinetArchiveReader.Read(cabinet, 1024 * 1024, 1024 * 1024, 10));

        Assert.Contains(exception.Code, new[] { "ONENOTE_CAB_CHECKSUM", "ONENOTE_CAB_LZX_TRUNCATED", "ONENOTE_CAB_LZX_CORRUPT" });
    }

    [Fact]
    public void RejectsCabinetEntryPastConfiguredLimit() {
        byte[] cabinet = BuildUncompressedCabinet("large.one", new byte[128]);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteCabinetArchiveReader.Read(cabinet, 1024, 64, 10));

        Assert.Equal("ONENOTE_CAB_ENTRY_LIMIT", exception.Code);
    }

    [Fact]
    public void RejectsAliasedCabinetEntriesPastAggregateExpansionLimit() {
        byte[] cabinet = BuildUncompressedCabinet(new[] { "first.one", "alias.one" }, new byte[80]);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteCabinetArchiveReader.Read(cabinet, 100, 100, 10));

        Assert.Equal("ONENOTE_CAB_EXPANDED_LIMIT", exception.Code);
    }

    [Fact]
    public void RejectsUncompressedFolderPastManagedBufferCapacityBeforeNarrowing() {
        byte[] cabinet = BuildOversizedDeclaredUncompressedCabinet();

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteCabinetArchiveReader.Read(cabinet, long.MaxValue, 1, 1));

        Assert.Equal("ONENOTE_CAB_EXPANDED_LIMIT", exception.Code);
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

    private static byte[] BuildUncompressedCabinet(string name, byte[] payload) =>
        BuildUncompressedCabinet(new[] { name }, payload);

    private static byte[] BuildUncompressedCabinet(IReadOnlyList<string> names, byte[] payload) {
        byte[][] nameBytes = names.Select(name => Encoding.UTF8.GetBytes(name + "\0")).ToArray();
        const int folderOffset = 36;
        const int filesOffset = folderOffset + 8;
        int dataOffset = filesOffset + nameBytes.Sum(bytes => 16 + bytes.Length);
        int cabinetSize = dataOffset + 8 + payload.Length;
        var data = new byte[cabinetSize];
        data[0] = (byte)'M'; data[1] = (byte)'S'; data[2] = (byte)'C'; data[3] = (byte)'F';
        WriteUInt32(data, 8, (uint)cabinetSize);
        WriteUInt32(data, 16, filesOffset);
        data[24] = 3; data[25] = 1;
        WriteUInt16(data, 26, 1);
        WriteUInt16(data, 28, checked((ushort)names.Count));
        WriteUInt32(data, folderOffset, (uint)dataOffset);
        WriteUInt16(data, folderOffset + 4, 1);
        WriteUInt16(data, folderOffset + 6, 0);
        int fileOffset = filesOffset;
        foreach (byte[] encodedName in nameBytes) {
            WriteUInt32(data, fileOffset, (uint)payload.Length);
            WriteUInt32(data, fileOffset + 4, 0);
            WriteUInt16(data, fileOffset + 8, 0);
            WriteUInt16(data, fileOffset + 14, 0x80);
            Buffer.BlockCopy(encodedName, 0, data, fileOffset + 16, encodedName.Length);
            fileOffset += 16 + encodedName.Length;
        }
        WriteUInt16(data, dataOffset + 4, (ushort)payload.Length);
        WriteUInt16(data, dataOffset + 6, (ushort)payload.Length);
        Buffer.BlockCopy(payload, 0, data, dataOffset + 8, payload.Length);
        return data;
    }

    private static byte[] BuildOversizedDeclaredUncompressedCabinet() {
        const int blockCount = 32769;
        const int folderOffset = 36;
        const int dataOffset = folderOffset + 8;
        var data = new byte[dataOffset + blockCount * 8];
        data[0] = (byte)'M'; data[1] = (byte)'S'; data[2] = (byte)'C'; data[3] = (byte)'F';
        WriteUInt32(data, 8, (uint)data.Length);
        WriteUInt32(data, 16, dataOffset);
        data[24] = 3; data[25] = 1;
        WriteUInt16(data, 26, 1);
        WriteUInt16(data, 28, 0);
        WriteUInt32(data, folderOffset, dataOffset);
        WriteUInt16(data, folderOffset + 4, blockCount);
        WriteUInt16(data, folderOffset + 6, 0);
        for (int index = 0; index < blockCount; index++) {
            int blockOffset = dataOffset + index * 8;
            WriteUInt16(data, blockOffset + 4, 0);
            WriteUInt16(data, blockOffset + 6, ushort.MaxValue);
        }
        return data;
    }

    private static byte[] BuildUncompressedLzxChunk(byte[] payload) =>
        BuildUncompressedLzxChunk(new[] { payload });

    private static byte[] BuildUncompressedLzxChunk(IReadOnlyList<byte[]> blocks) {
        using var output = new MemoryStream();
        for (int blockIndex = 0; blockIndex < blocks.Count; blockIndex++) {
            byte[] payload = blocks[blockIndex];
            var bits = new List<int>();
            if (blockIndex == 0) AppendBits(bits, 0, 1);
            AppendBits(bits, 3, 3);
            AppendBits(bits, payload.Length >> 8, 16);
            AppendBits(bits, payload.Length & 0xFF, 8);
            byte[] header = PackWords(bits);
            output.Write(header, 0, header.Length);
            WriteUInt32(output, 1);
            WriteUInt32(output, 1);
            WriteUInt32(output, 1);
            output.Write(payload, 0, payload.Length);
            if ((payload.Length & 1) != 0) output.WriteByte(0);
        }
        return output.ToArray();
    }

    private static string FixturePath(string fileName) =>
        Path.Combine(AppContext.BaseDirectory, "Fixtures", fileName);

    private static byte[] CreateE8OraclePayload() {
        byte[] pattern = Encoding.ASCII.GetBytes("OfficeIMO-LZX-E8-independent-oracle-");
        var payload = new byte[4096];
        for (int index = 0; index < payload.Length; index++) payload[index] = pattern[index % pattern.Length];
        WriteE8Call(payload, 64, 500);
        WriteE8Call(payload, 1024, -50);
        WriteE8Call(payload, 2048, 1_000_000);
        return payload;
    }

    private static void WriteE8Call(byte[] payload, int offset, int displacement) {
        payload[offset] = 0xE8;
        byte[] value = BitConverter.GetBytes(displacement);
        Buffer.BlockCopy(value, 0, payload, offset + 1, value.Length);
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

    private static void WriteUInt32(Stream stream, uint value) {
        stream.WriteByte((byte)value);
        stream.WriteByte((byte)(value >> 8));
        stream.WriteByte((byte)(value >> 16));
        stream.WriteByte((byte)(value >> 24));
    }
}
