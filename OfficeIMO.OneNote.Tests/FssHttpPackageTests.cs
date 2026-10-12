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

    [Fact]
    public void ObjectSpaceReferenceCellsMayNameAnotherContext() {
        OneNoteExtendedGuid rootSpaceId = Id();
        OneNoteExtendedGuid objectId = Id();
        var graph = new OneNoteWriteGraph(Guid.NewGuid(), OneNoteFileKind.Section, rootSpaceId, Guid.Empty, 0);
        graph.ObjectSpaces.Add(new OneNoteWriteObjectSpace(rootSpaceId, Id()));
        var space = new OneNoteWriteObjectSpace(Id(), Id(), contextId: Id());
        space.Objects.Add(new OneNoteWriteObject(objectId, OneNoteSchema.JcidPropertyContainer, new[] {
            new OneNoteWriteProperty(ObjectSpaceReferenceProperty, references: new[] { rootSpaceId }, referenceKind: OneNoteWriteReferenceKind.ObjectSpace)
        }));
        graph.ObjectSpaces.Add(space);

        byte[] data = OneNotePackageStoreWriter.Write(graph);
        ReplaceFirstCellReferenceContext(data, objectId, DefaultContext);

        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(new MemoryStream(data));

        OneNoteRevisionStoreObject record = Assert.Single(store.Objects, item => item.Id.Equals(objectId));
        OneNotePropertyValue? property = record.PropertySet?.Find(ObjectSpaceReferenceProperty);
        Assert.NotNull(property);
        Assert.Equal(rootSpaceId, Assert.Single(property!.ReferencedIds));
    }

    [Fact]
    public void NullObjectReferencesWithoutExtendedGuidsAreMapped() {
        OneNoteExtendedGuid targetId = Id();
        OneNoteExtendedGuid objectId = Id();
        byte[] data = WriteObjectWithTrailingNullReference(targetId, objectId);
        RemoveTrailingNullExtendedGuid(data, objectId, nullCompactId: true);

        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(new MemoryStream(data));

        OneNoteRevisionStoreObject record = Assert.Single(store.Objects, item => item.Id.Equals(objectId));
        OneNotePropertyValue? property = record.PropertySet?.Find(ObjectReferencesProperty);
        Assert.NotNull(property);
        Assert.Equal(2, property!.ReferencedIds.Count);
        Assert.Equal(targetId, property.ReferencedIds[0]);
    }

    [Fact]
    public void UnpairedObjectReferencesAreBoundedFormatErrors() {
        OneNoteExtendedGuid objectId = Id();
        byte[] data = WriteObjectWithTrailingNullReference(Id(), objectId);
        RemoveTrailingNullExtendedGuid(data, objectId, nullCompactId: false);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteRevisionStoreReader.Read(new MemoryStream(data)));

        Assert.Equal("ONENOTE_PACKAGE_MAPPING_COUNT", exception.Code);
    }

    private const uint ObjectReferencesProperty = ((uint)OneNotePropertyType.ObjectIdArray << 26) | 0x1001;
    private const uint ObjectSpaceReferenceProperty = ((uint)OneNotePropertyType.ObjectSpaceId << 26) | 0x1002;
    private static readonly OneNoteExtendedGuid DefaultContext = new OneNoteExtendedGuid(new Guid("84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073"), 1, 17);

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

    private static byte[] WriteObjectWithTrailingNullReference(OneNoteExtendedGuid targetId, OneNoteExtendedGuid objectId) {
        OneNoteExtendedGuid rootSpaceId = Id();
        var graph = new OneNoteWriteGraph(Guid.NewGuid(), OneNoteFileKind.Section, rootSpaceId, Guid.Empty, 0);
        var space = new OneNoteWriteObjectSpace(rootSpaceId, Id());
        space.Objects.Add(new OneNoteWriteObject(targetId, OneNoteSchema.JcidPropertyContainer));
        space.Objects.Add(new OneNoteWriteObject(objectId, OneNoteSchema.JcidPropertyContainer, new[] {
            new OneNoteWriteProperty(ObjectReferencesProperty, references: new[] { targetId, new OneNoteExtendedGuid(Guid.Empty, 0, 1) })
        }));
        graph.ObjectSpaces.Add(space);
        return OneNotePackageStoreWriter.Write(graph);
    }

    // Drops the trailing null extended GUID and re-encodes the empty cell-ID count in two bytes, keeping every length.
    private static void RemoveTrailingNullExtendedGuid(byte[] data, OneNoteExtendedGuid objectId, bool nullCompactId) {
        FssHttpStreamObject objectData = FindPropertySetObject(data, objectId, out FssHttpStreamObject declaration);
        int start = checked((int)objectData.DataOffset);
        int nullEntry = start + 1 + 17;
        Assert.Equal(0x05, data[start]); // compact uint64 value 2
        Assert.Equal(0x00, data[nullEntry]);
        Assert.Equal(0x01, data[nullEntry + 1]); // compact uint64 value 0
        data[start] = 0x03;
        data[nullEntry] = 0x02;
        data[nullEntry + 1] = 0x00;

        using var stream = new MemoryStream(data);
        byte[] prefix = FssHttpStreamObjectReader.ReadData(stream, declaration, 256, "object declaration");
        var cursor = new FssHttpDataCursor(prefix, declaration.DataOffset);
        cursor.ReadExtendedGuid();
        cursor.ReadCompactUInt64();
        cursor.ReadCompactUInt64();
        int objectReferenceCount = checked((int)declaration.DataOffset + cursor.Position);
        Assert.Equal(0x05, data[objectReferenceCount]);
        data[objectReferenceCount] = 0x03;

        if (nullCompactId) {
            // Second CompactID: after the property-data length byte, the stream header and the first CompactID.
            int compactId = nullEntry + 2 + 1 + 4 + 4;
            Assert.Equal(0x100U, BitConverter.ToUInt32(data, compactId));
            Array.Clear(data, compactId, 4);
        }
    }

    private static void ReplaceFirstCellReferenceContext(byte[] data, OneNoteExtendedGuid objectId, OneNoteExtendedGuid contextId) {
        FssHttpStreamObject objectData = FindPropertySetObject(data, objectId, out _);
        using var stream = new MemoryStream(data);
        byte[] prefix = FssHttpStreamObjectReader.ReadData(stream, objectData, 4096, "object data");
        var cursor = new FssHttpDataCursor(prefix, objectData.DataOffset);
        Assert.Equal(0UL, cursor.ReadCompactUInt64());
        Assert.Equal(1UL, cursor.ReadCompactUInt64());
        int first = cursor.Position;
        Assert.NotEqual(contextId, cursor.ReadExtendedGuid());

        using var encoded = new MemoryStream();
        FssHttpStreamObjectWriter.WriteExtendedGuid(encoded, contextId);
        Assert.Equal(cursor.Position - first, checked((int)encoded.Length));
        Buffer.BlockCopy(encoded.ToArray(), 0, data, checked((int)objectData.DataOffset + first), checked((int)encoded.Length));
    }

    private static FssHttpStreamObject FindPropertySetObject(byte[] data, OneNoteExtendedGuid objectId, out FssHttpStreamObject declaration) {
        using var stream = new MemoryStream(data);
        FssHttpStreamObject packaging = FssHttpStreamObjectReader.ReadPackaging(stream, new OneNoteReaderOptions());
        FssHttpStreamObject package = Assert.Single(packaging.Children, item => item.Type == 0x15);
        foreach (FssHttpStreamObject group in package.Children.Where(item => DataElementType(stream, item) == 5)) {
            FssHttpStreamObject[] declarations = Assert.Single(group.Children, item => item.Type == 0x1D).Children
                .Where(item => item.Type == 0x18 || item.Type == 0x05)
                .ToArray();
            FssHttpStreamObject[] objectData = Assert.Single(group.Children, item => item.Type == 0x1E).Children
                .Where(item => item.Type == 0x16 || item.Type == 0x1C)
                .ToArray();
            for (int index = 0; index < declarations.Length; index++) {
                if (declarations[index].Type != 0x18) continue;
                byte[] prefix = FssHttpStreamObjectReader.ReadData(stream, declarations[index], 256, "object declaration");
                var cursor = new FssHttpDataCursor(prefix, declarations[index].DataOffset);
                if (!cursor.ReadExtendedGuid().Equals(objectId) || cursor.ReadCompactUInt64() != 1) continue;
                declaration = declarations[index];
                return objectData[index];
            }
        }
        throw new InvalidOperationException("The property-set object was not written.");
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
