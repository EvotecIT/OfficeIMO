namespace OfficeIMO.Archives.Tests;

public sealed class CabinetLzxContractTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void TrailingLzxBytesAreRejectedEvenWhenAllMembersAreExcluded(bool literal, bool exclude) {
        using var fixture = new ArchiveFixtures();
        byte[] cabinet = ArchiveFixtures.LzxCab(literal, new byte[] { 0x55, 0x55 });
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(fixture.Package(cabinet), fixture.Destination,
            new ArchiveExtractionOptions { IncludeFile = _ => !exclude }));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Fact]
    public void CompressedLzxFrameAlignmentBitsMustBeZero() {
        using var fixture = new ArchiveFixtures();
        Assert.Throws<InvalidDataException>(() => ArchiveExtractor.Extract(
            fixture.Package(ArchiveFixtures.LzxCab(true, corruptPadding: true)), fixture.Destination));
        Assert.False(Directory.Exists(fixture.Destination));
    }

    [Theory]
    [InlineData(false, "ABCD")]
    [InlineData(true, "AAAA")]
    public void CompleteLzxFramesRetainTheirDeclaredPayload(bool literal, string text) {
        using var fixture = new ArchiveFixtures();
        Assert.Single(ArchiveExtractor.Extract(fixture.Package(ArchiveFixtures.LzxCab(literal)), fixture.Destination).Files);
        Assert.Equal(text, File.ReadAllText(Path.Combine(fixture.Destination, "Template.admx")));
    }
}
