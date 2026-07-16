namespace OfficeIMO.OneNote.Tests;

public sealed class CabinetWriterTests {
    [Fact]
    public void RejectsCabinetOutputBeyondInMemoryCapacityBeforeNarrowing() {
        long cabinetSize = (long)int.MaxValue + 1;

        IOException exception = Assert.Throws<IOException>(() =>
            OneNoteCabinetArchiveWriter.GetOutputCapacity(cabinetSize, uint.MaxValue));

        Assert.Contains("supported in-memory size", exception.Message);
    }
}
