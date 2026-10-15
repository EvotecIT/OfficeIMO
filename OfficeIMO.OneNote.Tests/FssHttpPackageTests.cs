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
            new OneNoteWriteProperty(ObjectSpaceReferenceProperty, references: new[] { rootSpaceId }, referenceKind: OneNoteWriteReferenceKind.ObjectSpace),
            new OneNoteWriteProperty(ContextReferenceProperty, references: new[] { DefaultContext }, referenceKind: OneNoteWriteReferenceKind.Context)
        }));
        graph.ObjectSpaces.Add(space);

        byte[] data = OneNotePackageStoreWriter.Write(graph);
        ReplaceFirstCellReferenceContext(data, objectId, DefaultContext);

        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(new MemoryStream(data));

        OneNoteRevisionStoreObject record = Assert.Single(store.Objects, item => item.Id.Equals(objectId));
        OneNotePropertyValue? property = record.PropertySet?.Find(ObjectSpaceReferenceProperty);
        Assert.NotNull(property);
        Assert.Equal(rootSpaceId, Assert.Single(property!.ReferencedIds));
        Assert.Equal(DefaultContext, Assert.Single(record.PropertySet!.Find(ContextReferenceProperty)!.ReferencedIds));
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
        Assert.Equal(Guid.Empty, property.ReferencedIds[1].Identifier);
        Assert.Equal(0U, property.ReferencedIds[1].Value);
    }

    [Fact]
    public void ExplicitNullObjectReferencesArePreserved() {
        OneNoteExtendedGuid targetId = Id();
        OneNoteExtendedGuid objectId = Id();
        byte[] data = WriteObjectWithTrailingNullReference(targetId, objectId);

        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(new MemoryStream(data));

        OneNotePropertyValue property = Assert.Single(store.Objects, item => item.Id.Equals(objectId)).PropertySet!.Find(ObjectReferencesProperty)!;
        Assert.Equal(targetId, property.ReferencedIds[0]);
        Assert.Equal(Guid.Empty, property.ReferencedIds[1].Identifier);
        Assert.Equal(0U, property.ReferencedIds[1].Value);
    }

    [Theory]
    [InlineData(256U, 1U)]
    [InlineData(uint.MaxValue, 0xFFFFFFFEU)]
    public void PackageMappingsPreserveTheFullExtendedGuidValue(uint value, uint compact) {
        OneNoteExtendedGuid targetId = Id();
        OneNoteExtendedGuid objectId = Id();
        var properties = new[] {
            new OneNoteWriteProperty(ObjectReferencesProperty, references: new[] { targetId })
        };
        OneNoteEncodedPropertySet encoded = OneNotePropertySetWriter.Write(properties, Id(), DefaultContext);
        using var compactBytes = new MemoryStream();
        FssHttpStreamObjectWriter.WriteUInt32(compactBytes, compact);
        Buffer.BlockCopy(compactBytes.ToArray(), 0, encoded.Data, 4, 4);
        var mappedId = new OneNoteExtendedGuid(targetId.Identifier, value, 21);
        byte[] data = ReplacePropertySet(WriteObjectWithTrailingNullReference(targetId, objectId), objectId,
            new OneNoteEncodedPropertySet(encoded.Data, new[] { mappedId }, encoded.CellReferences));

        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(new MemoryStream(data));

        OneNotePropertyValue property = Assert.Single(store.Objects, item => item.Id.Equals(objectId)).PropertySet!.Find(ObjectReferencesProperty)!;
        Assert.Equal(mappedId, Assert.Single(property.ReferencedIds));
    }

    [Fact]
    public void ConflictingValuesForTheSamePackageCompactIdAreRejected() {
        OneNoteExtendedGuid targetId = Id();
        OneNoteExtendedGuid objectId = Id();
        var properties = new[] {
            new OneNoteWriteProperty(ObjectReferencesProperty, references: new[] { targetId, targetId })
        };
        OneNoteEncodedPropertySet encoded = OneNotePropertySetWriter.Write(properties, Id(), DefaultContext);
        Buffer.BlockCopy(encoded.Data, 4, encoded.Data, 8, 4);
        byte[] data = ReplacePropertySet(WriteObjectWithTrailingNullReference(targetId, objectId), objectId,
            new OneNoteEncodedPropertySet(encoded.Data,
                new[] { targetId, new OneNoteExtendedGuid(targetId.Identifier, targetId.Value + 1, 17) }, encoded.CellReferences));

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() => OneNoteRevisionStoreReader.Read(new MemoryStream(data)));

        Assert.Equal("ONENOTE_PACKAGE_MAPPING", exception.Code);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NullOnlyObjectReferencesDoNotRequireGlobalMappings(bool omitExtendedGuid) {
        OneNoteExtendedGuid objectId = Id();
        OneNoteExtendedGuid rootSpaceId = Id();
        var properties = new[] {
            new OneNoteWriteProperty(ObjectReferencesProperty, references: new[] { new OneNoteExtendedGuid(Guid.Empty, 0, 1) })
        };
        var graph = new OneNoteWriteGraph(Guid.NewGuid(), OneNoteFileKind.Section, rootSpaceId, Guid.Empty, 0);
        var space = new OneNoteWriteObjectSpace(rootSpaceId, Id());
        space.Objects.Add(new OneNoteWriteObject(objectId, OneNoteSchema.JcidPropertyContainer, properties));
        graph.ObjectSpaces.Add(space);
        byte[] data = OneNotePackageStoreWriter.Write(graph);
        if (omitExtendedGuid) {
            OneNoteEncodedPropertySet encoded = OneNotePropertySetWriter.Write(properties, rootSpaceId, DefaultContext);
            data = ReplacePropertySet(data, objectId, new OneNoteEncodedPropertySet(encoded.Data,
                Array.Empty<OneNoteExtendedGuid>(), encoded.CellReferences));
        }

        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(new MemoryStream(data));

        OneNoteExtendedGuid reference = Assert.Single(Assert.Single(store.Objects, item => item.Id.Equals(objectId))
            .PropertySet!.Find(ObjectReferencesProperty)!.ReferencedIds);
        Assert.Equal(Guid.Empty, reference.Identifier);
        Assert.Equal(0U, reference.Value);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void OmittedNullReferencesDoNotShiftOtherReferenceStreams(int nullKind) {
        OneNoteExtendedGuid rootSpaceId = Id();
        OneNoteExtendedGuid objectId = Id();
        OneNoteExtendedGuid targetId = Id();
        OneNoteExtendedGuid contextId = Id();
        OneNoteExtendedGuid nullId = new OneNoteExtendedGuid(Guid.Empty, 0, 1);
        OneNoteExtendedGuid[][] references = {
            new[] { targetId }, new[] { rootSpaceId }, new[] { contextId }
        };
        references[nullKind] = new[] { references[nullKind][0], nullId };
        uint[] propertyIds = {
            ObjectReferencesProperty,
            ((uint)OneNotePropertyType.ObjectSpaceIdArray << 26) | 0x1002,
            ((uint)OneNotePropertyType.ContextIdArray << 26) | 0x1003
        };
        OneNoteWriteProperty[] properties = Enumerable.Range(0, 3).Select(index =>
            new OneNoteWriteProperty(propertyIds[index], references: references[index], referenceKind: (OneNoteWriteReferenceKind)index)).ToArray();
        var graph = new OneNoteWriteGraph(Guid.NewGuid(), OneNoteFileKind.Section, rootSpaceId, Guid.Empty, 0);
        var space = new OneNoteWriteObjectSpace(rootSpaceId, Id());
        space.Objects.Add(new OneNoteWriteObject(objectId, OneNoteSchema.JcidPropertyContainer, properties));
        graph.ObjectSpaces.Add(space);
        OneNoteEncodedPropertySet encoded = OneNotePropertySetWriter.Write(properties, rootSpaceId, DefaultContext);
        byte[] propertyData = (byte[])encoded.Data.Clone();
        int position = 0;
        for (int kind = 0; kind < 3; kind++) {
            position += 4; // Stream header.
            for (int index = 0; index < references[kind].Length; index++, position += 4) {
                if (references[kind][index].Identifier == Guid.Empty) Array.Clear(propertyData, position, 4);
            }
        }
        var replacement = new OneNoteEncodedPropertySet(propertyData,
            encoded.ObjectReferences.Where(id => id.Identifier != Guid.Empty).ToArray(),
            encoded.CellReferences.Where(cell => cell.First.Identifier != Guid.Empty && cell.Second.Identifier != Guid.Empty).ToArray());
        byte[] data = ReplacePropertySet(OneNotePackageStoreWriter.Write(graph), objectId, replacement);

        OneNoteRevisionStore store = OneNoteRevisionStoreReader.Read(new MemoryStream(data));

        OneNotePropertySet result = Assert.Single(store.Objects, item => item.Id.Equals(objectId)).PropertySet!;
        for (int kind = 0; kind < 3; kind++) {
            OneNotePropertyValue property = result.Find(propertyIds[kind])!;
            Assert.Equal(references[kind], property.ReferencedIds);
        }
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

    [Fact]
    public void EmbeddedFileBlobsAreLengthPrefixedBinaryItems() {
        byte[] payload = Enumerable.Range(0, 300).Select(index => (byte)index).ToArray();
        byte[] data = WriteSectionWithEmbeddedFiles(payload);

        using (var stream = new MemoryStream(data)) {
            FssHttpStreamObject blob = Assert.Single(FindObjectDataBlobs(stream));
            byte[] raw = FssHttpStreamObjectReader.ReadData(stream, blob, int.MaxValue, "object-data BLOB");
            Assert.Equal(302, raw.Length);
            Assert.Equal(new byte[] { 0xB2, 0x04 }, raw.Take(2).ToArray()); // compact uint64 300
            Assert.Equal(payload, raw.Skip(2).ToArray());
        }

        OneNoteSection section = OneNoteSectionReader.Read(new MemoryStream(data));
        OneNoteEmbeddedFile file = Assert.Single(section.Pages.SelectMany(page => page.Outlines).SelectMany(outline => outline.Children).OfType<OneNoteEmbeddedFile>());
        Assert.Equal(payload, file.Payload!.ToArray(payload.Length));
    }

    [Fact]
    public void AssetLimitsApplyToTheBlobPayloadLength() {
        byte[] first = Enumerable.Repeat((byte)1, 300).ToArray();
        byte[] second = Enumerable.Repeat((byte)2, 300).ToArray();
        byte[] data = WriteSectionWithEmbeddedFiles(first, second);

        OneNoteSection section = OneNoteSectionReader.Read(new MemoryStream(data), new OneNoteReaderOptions { MaxAssetBytes = 300, MaxTotalAssetBytes = 600 });

        Assert.Equal(new[] { first, second }, section.Pages.SelectMany(page => page.Outlines).SelectMany(outline => outline.Children).OfType<OneNoteEmbeddedFile>().Select(file => file.Payload!.ToArray(300)).ToArray());
        Assert.Equal("ONENOTE_ASSET_LIMIT", Assert.Throws<OneNoteFormatException>(() => OneNoteRevisionStoreReader.Read(
            new MemoryStream(data),
            new OneNoteReaderOptions { MaxAssetBytes = 299 })).Code);
        Assert.Equal("ONENOTE_ASSET_LIMIT", Assert.Throws<OneNoteFormatException>(() => OneNoteRevisionStoreReader.Read(
            new MemoryStream(data),
            new OneNoteReaderOptions { MaxAssetBytes = 300, MaxTotalAssetBytes = 599 })).Code);
    }

    [Fact]
    public void BlobLengthThatDisagreesWithItsStreamObjectIsABoundedFormatError() {
        byte[] data = WriteSectionWithEmbeddedFiles(new byte[300]);
        int offset;
        using (var stream = new MemoryStream(data)) {
            offset = checked((int)Assert.Single(FindObjectDataBlobs(stream)).DataOffset);
        }
        Assert.Equal((300 << 2) | 0b10, BitConverter.ToUInt16(data, offset)); // two-byte compact uint64
        byte[] shorter = BitConverter.GetBytes((ushort)((299 << 2) | 0b10));
        Buffer.BlockCopy(shorter, 0, data, offset, shorter.Length);

        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(() =>
            OneNoteRevisionStoreReader.Read(new MemoryStream(data)));

        Assert.Equal("ONENOTE_PACKAGE_ITEM_LENGTH", exception.Code);
    }

    private const uint ObjectReferencesProperty = ((uint)OneNotePropertyType.ObjectIdArray << 26) | 0x1001;
    private const uint ObjectSpaceReferenceProperty = ((uint)OneNotePropertyType.ObjectSpaceId << 26) | 0x1002;
    private const uint ContextReferenceProperty = ((uint)OneNotePropertyType.ContextId << 26) | 0x1003;
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
        Assert.Equal(2UL, cursor.ReadCompactUInt64());
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

    private static byte[] ReplacePropertySet(byte[] data, OneNoteExtendedGuid objectId, OneNoteEncodedPropertySet replacement) {
        FssHttpStreamObject objectData = FindPropertySetObject(data, objectId, out FssHttpStreamObject declaration);
        using var source = new MemoryStream(data);
        FssHttpStreamObject packaging = FssHttpStreamObjectReader.ReadPackaging(source, new OneNoteReaderOptions());
        using var declarationData = new MemoryStream();
        FssHttpStreamObjectWriter.WriteExtendedGuid(declarationData, objectId);
        FssHttpStreamObjectWriter.WriteCompactUInt64(declarationData, 1);
        FssHttpStreamObjectWriter.WriteCompactUInt64(declarationData, (ulong)replacement.Data.Length);
        FssHttpStreamObjectWriter.WriteCompactUInt64(declarationData, (ulong)replacement.ObjectReferences.Count);
        FssHttpStreamObjectWriter.WriteCompactUInt64(declarationData, (ulong)replacement.CellReferences.Count);
        using var payload = new MemoryStream();
        FssHttpStreamObjectWriter.WriteCompactUInt64(payload, (ulong)replacement.ObjectReferences.Count);
        foreach (OneNoteExtendedGuid id in replacement.ObjectReferences) FssHttpStreamObjectWriter.WriteExtendedGuid(payload, id);
        FssHttpStreamObjectWriter.WriteCompactUInt64(payload, (ulong)replacement.CellReferences.Count);
        foreach (FssHttpCellId cell in replacement.CellReferences) {
            FssHttpStreamObjectWriter.WriteExtendedGuid(payload, cell.First);
            FssHttpStreamObjectWriter.WriteExtendedGuid(payload, cell.Second);
        }
        FssHttpStreamObjectWriter.WriteCompactUInt64(payload, (ulong)replacement.Data.Length);
        payload.Write(replacement.Data, 0, replacement.Data.Length);
        FssHttpWriteObject Rewrite(FssHttpStreamObject node) {
            byte[] bytes = node.HeaderOffset == declaration.HeaderOffset ? declarationData.ToArray()
                : node.HeaderOffset == objectData.HeaderOffset ? payload.ToArray()
                : FssHttpStreamObjectReader.ReadData(source, node, int.MaxValue, "test package node");
            return new FssHttpWriteObject(node.Type, bytes, node.Children.Select(Rewrite));
        }
        using var output = new MemoryStream();
        output.Write(data, 0, checked((int)packaging.HeaderOffset));
        FssHttpStreamObjectWriter.WriteObject(output, Rewrite(packaging));
        return output.ToArray();
    }

    private static byte[] WriteSectionWithEmbeddedFiles(params byte[][] payloads) {
        var section = new OneNoteSection { Name = "Attachments" };
        var page = new OneNotePage { Title = "Attachment page" };
        for (int index = 0; index < payloads.Length; index++) {
            page.DirectContent.Add(new OneNoteEmbeddedFile {
                FileName = "attachment-" + index + ".bin",
                Payload = OneNoteBinaryPayload.FromBytes(payloads[index])
            });
        }
        section.Pages.Add(page);
        return OneNoteSectionWriter.Write(section, new OneNoteWriterOptions { StorageFormat = OneNoteStorageFormat.FileSynchronizationPackage });
    }

    private static IEnumerable<FssHttpStreamObject> FindObjectDataBlobs(Stream stream) {
        FssHttpStreamObject packaging = FssHttpStreamObjectReader.ReadPackaging(stream, new OneNoteReaderOptions());
        FssHttpStreamObject package = Assert.Single(packaging.Children, item => item.Type == 0x15);
        return package.Children
            .Where(item => DataElementType(stream, item) == 10)
            .Select(item => Assert.Single(item.Children, child => child.Type == 0x02))
            .ToArray();
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
