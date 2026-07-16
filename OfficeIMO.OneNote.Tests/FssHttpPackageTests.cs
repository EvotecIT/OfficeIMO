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

    [Fact]
    public void DuplicateRevisionManifestIdentifiersAreBoundedFormatErrors() {
        OneNoteExtendedGuid rootSpaceId = Id();
        var graph = new OneNoteWriteGraph(Guid.NewGuid(), OneNoteFileKind.Section, rootSpaceId, Guid.Empty, 0);
        graph.ObjectSpaces.Add(new OneNoteWriteObjectSpace(rootSpaceId, Id()));
        graph.ObjectSpaces.Add(new OneNoteWriteObjectSpace(Id(), Id()));

        byte[] data = OneNotePackageStoreWriter.Write(graph);
        DuplicateSecondUserRevisionIdentifier(data);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteRevisionStoreReader.Read(new MemoryStream(data)));

        Assert.Equal("ONENOTE_PACKAGE_REVISION_ID", exception.Code);
    }

    [Fact]
    public void DuplicateStorageIndexesAreBoundedFormatErrors() {
        OneNoteExtendedGuid rootSpaceId = Id();
        var graph = new OneNoteWriteGraph(Guid.NewGuid(), OneNoteFileKind.Section, rootSpaceId, Guid.Empty, 0);
        graph.ObjectSpaces.Add(new OneNoteWriteObjectSpace(rootSpaceId, Id()));

        byte[] data = OneNotePackageStoreWriter.Write(graph);
        ChangeFirstNonStorageElementToStorageIndex(data);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteRevisionStoreReader.Read(new MemoryStream(data)));

        Assert.Equal("ONENOTE_PACKAGE_STORAGE_INDEX", exception.Code);
    }

    private static string FixturePath(string fileName) => Path.Combine(AppContext.BaseDirectory, "Fixtures", fileName);

    private static OneNoteExtendedGuid Id() => new OneNoteExtendedGuid(Guid.NewGuid(), 1, 17);

    private static void DuplicateSecondUserRevisionIdentifier(byte[] data) {
        using var stream = new MemoryStream(data);
        FssHttpStreamObject packaging = FssHttpStreamObjectReader.ReadPackaging(stream, new OneNoteReaderOptions());
        FssHttpStreamObject package = Assert.Single(packaging.Children, item => item.Type == 0x15);
        FssHttpStreamObject[] revisionDeclarations = package.Children
            .Where(item => DataElementType(stream, item) == 4)
            .Select(item => Assert.Single(item.Children, child => child.Type == 0x1A))
            .Skip(1)
            .Take(2)
            .ToArray();
        Assert.Equal(2, revisionDeclarations.Length);

        const int encodedGuidBytes = 17;
        Buffer.BlockCopy(
            data,
            checked((int)revisionDeclarations[0].DataOffset),
            data,
            checked((int)revisionDeclarations[1].DataOffset),
            encodedGuidBytes);
    }

    private static void ChangeFirstNonStorageElementToStorageIndex(byte[] data) {
        using var stream = new MemoryStream(data);
        FssHttpStreamObject packaging = FssHttpStreamObjectReader.ReadPackaging(stream, new OneNoteReaderOptions());
        FssHttpStreamObject package = Assert.Single(packaging.Children, item => item.Type == 0x15);
        FssHttpStreamObject element = Assert.Single(
            package.Children.Where(item => DataElementType(stream, item) != 1).Take(1));
        byte[] prefix = FssHttpStreamObjectReader.ReadData(stream, element, 128, "data-element prefix");
        var cursor = new FssHttpDataCursor(prefix, element.DataOffset);
        cursor.ReadExtendedGuid();
        cursor.SkipSerialNumber();

        data[checked((int)element.DataOffset + cursor.Position)] = 0x03; // compact uint64 value 1
    }

    private static ulong DataElementType(Stream stream, FssHttpStreamObject element) {
        byte[] prefix = FssHttpStreamObjectReader.ReadData(stream, element, 128, "data-element prefix");
        var cursor = new FssHttpDataCursor(prefix, element.DataOffset);
        cursor.ReadExtendedGuid();
        cursor.SkipSerialNumber();
        ulong type = cursor.ReadCompactUInt64();
        cursor.EnsureEnd("data-element prefix");
        return type;
    }
}
