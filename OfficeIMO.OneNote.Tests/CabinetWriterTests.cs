namespace OfficeIMO.OneNote.Tests;

public sealed class CabinetWriterTests {
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
}
