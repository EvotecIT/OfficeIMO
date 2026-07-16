namespace OfficeIMO.OneNote.Tests;

public sealed class FssHttpPackageTests {
    [Theory]
    [InlineData("testOneNoteFromOffice365.one")]
    [InlineData("testOneNoteFromOffice365-2.one")]
    public void ParsesBoundedStreamObjectTree(string fixture) {
        using FileStream stream = File.OpenRead(FixturePath(fixture));
        var options = new OneNoteReaderOptions();

        FssHttpStreamObject packaging = FssHttpStreamObjectReader.ReadPackaging(stream, options);

        Assert.Equal(0x7A, packaging.Type);
        Assert.True(packaging.Compound);
        Assert.Equal(33UL, packaging.DataLength);
        FssHttpStreamObject package = Assert.Single(packaging.Children);
        Assert.Equal(0x15, package.Type);
        Assert.True(package.Compound);
        Assert.NotEmpty(package.Children);
        Assert.All(package.Children, element => Assert.Equal(0x01, element.Type));
        Assert.All(package.Children, element => {
            byte[] prefix = FssHttpStreamObjectReader.ReadData(stream, element, 128, "data-element prefix");
            var cursor = new FssHttpDataCursor(prefix, element.DataOffset);
            OneNoteExtendedGuid id = cursor.ReadExtendedGuid();
            cursor.SkipSerialNumber();
            ulong type = cursor.ReadCompactUInt64();
            cursor.EnsureEnd("data-element prefix");
            Assert.NotEqual(Guid.Empty, id.Identifier);
            Assert.Contains(type, new ulong[] { 1, 2, 3, 4, 5, 10 });
        });
    }

    [Fact]
    public void StreamObjectCountLimitStopsPackageTraversal() {
        using FileStream stream = File.OpenRead(FixturePath("testOneNoteFromOffice365.one"));
        var options = new OneNoteReaderOptions { MaxStreamObjects = 2 };

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() => FssHttpStreamObjectReader.ReadPackaging(stream, options));

        Assert.Equal("ONENOTE_PACKAGE_OBJECT_LIMIT", exception.Code);
    }

    private static string FixturePath(string fileName) => Path.Combine(AppContext.BaseDirectory, "Fixtures", fileName);
}
