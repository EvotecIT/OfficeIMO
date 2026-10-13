namespace OfficeIMO.OneNote.Tests;

public sealed class CabinetWriterTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WritesCabinetNamesAtThe256EncodedByteLimit(bool nonAscii) {
        string name = BoundaryName(nonAscii, extraBytes: 0);
        Assert.Equal(256, System.Text.Encoding.UTF8.GetByteCount(name));

        byte[] cabinet = OneNoteCabinetArchiveWriter.Write(
            new[] { new OneNoteCabinetEntry(name, new byte[] { 1 }) }, 1024);

        Assert.Equal(name.Replace('/', '\\'), Assert.Single(OneNoteCabinetArchiveReader.Read(cabinet, 1024, 1024, 10)).Name);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void RejectsCabinetNamesAbove256EncodedBytes(bool nonAscii) {
        string name = BoundaryName(nonAscii, extraBytes: 1);
        Assert.Equal(257, System.Text.Encoding.UTF8.GetByteCount(name));

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteCabinetArchiveWriter.Write(new[] { new OneNoteCabinetEntry(name, new byte[] { 1 }) }, 1024));

        Assert.Equal("ONENOTE_CAB_ENTRY_NAME", exception.Code);
    }

    [Fact]
    public void WritesAndValidatesCabinetDataChecksums() {
        byte[] cabinet = OneNoteCabinetArchiveWriter.Write(
            new[] { new OneNoteCabinetEntry("Section.one", new byte[] { 1, 2, 3, 4, 5 }) },
            1024 * 1024);
        int dataOffset = checked((int)BitConverter.ToUInt32(cabinet, 36));

        Assert.NotEqual(0u, BitConverter.ToUInt32(cabinet, dataOffset));
        Assert.Single(OneNoteCabinetArchiveReader.Read(cabinet, 1024, 1024, 10));

        cabinet[dataOffset + 8] ^= 0x01;
        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteCabinetArchiveReader.Read(cabinet, 1024, 1024, 10));
        Assert.Equal("ONENOTE_CAB_CHECKSUM", exception.Code);
    }

    [Fact]
    public void RejectsCabinetOutputBeyondInMemoryCapacityBeforeNarrowing() {
        long cabinetSize = (long)int.MaxValue + 1;

        IOException exception = Assert.Throws<IOException>(() =>
            OneNoteCabinetArchiveWriter.GetOutputCapacity(cabinetSize, uint.MaxValue));

        Assert.Contains("supported in-memory size", exception.Message);
    }

    private static string BoundaryName(bool nonAscii, int extraBytes) =>
        new string('a', 121) + "/" + new string('b', (nonAscii ? 128 : 130) + extraBytes) + (nonAscii ? "\u0105" : "") + ".one";
}
