namespace OfficeIMO.OneNote;

internal static class OneNotePackageStoreWriter {
    private static readonly OneNoteExtendedGuid DefaultContext = Extended("84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073", 1);
    private static readonly OneNoteExtendedGuid DataRoot = Extended("84DEFAB9-AAA3-4A0D-A3A8-520C77AC7073", 2);
    private static readonly OneNoteExtendedGuid HeaderRoot = Extended("1A5A319C-C26B-41AA-B9C5-9BD8C44E07D4", 1);
    private static readonly OneNoteExtendedGuid HeaderSpaceId = Extended("111E4CF3-7FEF-4087-AF6A-B9544ACD334D", 1);
    private static readonly Guid RevisionRootGuid = new Guid("4A3717F8-1C14-49E7-9526-81D942DE1741");
    private static readonly Guid HeaderRevisionXor = new Guid("F5367D2F-F167-4830-A9C3-68F8336A2A09");
    private static readonly Guid HeaderObjectGroupXor = new Guid("26402DF1-68C7-43C0-A37C-04B9100893CE");
    private static readonly OneNoteExtendedGuid HeaderObjectId = Extended("B4760B1A-FBDF-4AE3-9D08-53219D8A8D21", 1);
    private const uint FileFormatVersion = 42;

    internal static byte[] Write(OneNoteWriteGraph graph, long maxOutputBytes = long.MaxValue) {
        if (maxOutputBytes < 1) throw new ArgumentOutOfRangeException(nameof(maxOutputBytes), "MaxOutputBytes must be greater than zero.");
        var ids = new OneNoteWriteIdFactory();
        OneNoteExtendedGuid storageIndexId = ids.New();
        OneNoteExtendedGuid storageManifestId = ids.New();
        var storageIndexChildren = new List<FssHttpWriteObject>();
        var packageElements = new List<FssHttpWriteObject>();

        OneNoteWriteObjectSpace headerSpace = BuildHeaderSpace(graph, out OneNoteExtendedGuid headerGroupId);
        AddObjectSpace(headerSpace, headerGroupId, ids, storageIndexChildren, packageElements);
        foreach (OneNoteWriteObjectSpace space in graph.ObjectSpaces) {
            AddObjectSpace(space, ids.New(), ids, storageIndexChildren, packageElements);
        }

        packageElements.Insert(0, DataElement(storageIndexId, 1, storageIndexChildren));
        packageElements.Insert(1, DataElement(storageManifestId, 2, new[] {
            Simple(0x0C, (graph.FileKind == OneNoteFileKind.TableOfContents ? OneNoteFormatConstants.TableOfContentsCellSchema : OneNoteFormatConstants.SectionCellSchema).ToByteArray()),
            Simple(0x07, Data(stream => { WriteExtended(stream, HeaderRoot); WriteCellId(stream, new FssHttpCellId(DefaultContext, HeaderSpaceId)); })),
            Simple(0x07, Data(stream => { WriteExtended(stream, DataRoot); WriteCellId(stream, new FssHttpCellId(DefaultContext, graph.RootObjectSpaceId)); }))
        }));

        byte[] rootData = Data(stream => {
            WriteExtended(stream, storageIndexId);
            FssHttpStreamObjectWriter.WriteGuid(stream, graph.FileKind == OneNoteFileKind.TableOfContents ? OneNoteFormatConstants.TableOfContentsCellSchema : OneNoteFormatConstants.SectionCellSchema);
        });
        var root = new FssHttpWriteObject(0x7A, rootData, new[] { new FssHttpWriteObject(0x15, children: packageElements) });
        long packagingLength = FssHttpStreamObjectWriter.GetEncodedLength(root);
        long outputLength = checked(68L + packagingLength);
        if (outputLength > maxOutputBytes) throw new IOException("FSSHTTP OneNote output exceeds MaxOutputBytes.");
        if (outputLength > int.MaxValue) throw new IOException("FSSHTTP OneNote output exceeds the supported in-memory size.");
        var output = new byte[(int)outputLength];
        // MS-ONESTORE package-store envelopes use the .one file-type GUID for both
        // section and table-of-contents payloads; guidCellSchemaId distinguishes them.
        Buffer.BlockCopy(OneNoteFormatConstants.SectionFileType.ToByteArray(), 0, output, 0, 16);
        Buffer.BlockCopy(graph.FileId.ToByteArray(), 0, output, 16, 16);
        Buffer.BlockCopy(Guid.NewGuid().ToByteArray(), 0, output, 32, 16);
        Buffer.BlockCopy(OneNoteFormatConstants.PackageStoreFormat.ToByteArray(), 0, output, 48, 16);
        using (var stream = new MemoryStream(output, true)) {
            stream.Position = 68;
            FssHttpStreamObjectWriter.WriteObject(stream, root);
        }
        return output;
    }

    private static void AddObjectSpace(
        OneNoteWriteObjectSpace space,
        OneNoteExtendedGuid groupElementId,
        OneNoteWriteIdFactory ids,
        ICollection<FssHttpWriteObject> storageIndexChildren,
        ICollection<FssHttpWriteObject> packageElements) {
        OneNoteExtendedGuid cellManifestElementId = ids.New();
        OneNoteExtendedGuid revisionElementId = ids.New();
        var cellId = new FssHttpCellId(space.ContextId ?? DefaultContext, space.Id);
        storageIndexChildren.Add(Simple(0x0E, Data(stream => { WriteCellId(stream, cellId); WriteExtended(stream, cellManifestElementId); stream.WriteByte(0); })));
        storageIndexChildren.Add(Simple(0x0D, Data(stream => { WriteExtended(stream, space.RevisionId); WriteExtended(stream, revisionElementId); stream.WriteByte(0); })));
        packageElements.Add(DataElement(cellManifestElementId, 3, new[] { Simple(0x0B, Data(stream => WriteExtended(stream, space.RevisionId))) }));
        var revisionChildren = new List<FssHttpWriteObject> {
            Simple(0x1A, Data(stream => { WriteExtended(stream, space.RevisionId); WriteExtended(stream, space.DependencyId ?? Null()); }))
        };
        foreach (KeyValuePair<uint, OneNoteExtendedGuid> rootEntry in space.Roots.OrderBy(item => item.Key)) {
            revisionChildren.Add(Simple(0x0A, Data(stream => { WriteExtended(stream, new OneNoteExtendedGuid(RevisionRootGuid, rootEntry.Key, 17)); WriteExtended(stream, rootEntry.Value); })));
        }
        revisionChildren.Add(Simple(0x19, Data(stream => WriteExtended(stream, groupElementId))));
        packageElements.Add(DataElement(revisionElementId, 4, revisionChildren));
        packageElements.Add(DataElement(groupElementId, 5, BuildObjectGroup(space, ids, packageElements)));
    }

    private static OneNoteWriteObjectSpace BuildHeaderSpace(OneNoteWriteGraph graph, out OneNoteExtendedGuid objectGroupId) {
        Guid revisionGuid = Xor(Xor(graph.FileId, graph.AncestorId), HeaderRevisionXor);
        uint revisionValue = graph.FileNameCrc ^ FileFormatVersion;
        var revisionId = new OneNoteExtendedGuid(revisionGuid, revisionValue, 21);
        objectGroupId = new OneNoteExtendedGuid(Xor(revisionGuid, HeaderObjectGroupXor), revisionValue ^ 0x80000000U, 21);
        var space = new OneNoteWriteObjectSpace(HeaderSpaceId, revisionId);
        space.Roots[1] = HeaderObjectId;
        space.Objects.Add(new OneNoteWriteObject(HeaderObjectId, OneNoteSchema.JcidPropertyContainer, new[] {
            DataProperty(OneNoteSchema.FileIdentityGuid, graph.FileId.ToByteArray()),
            DataProperty(OneNoteSchema.FileAncestorIdentityGuid, graph.AncestorId.ToByteArray()),
            ScalarProperty(OneNoteSchema.FileLastCodeVersion, FileFormatVersion),
            ScalarProperty(OneNoteSchema.FileNameCrc, graph.FileNameCrc)
        }));
        return space;
    }

    private static IReadOnlyList<FssHttpWriteObject> BuildObjectGroup(OneNoteWriteObjectSpace space, OneNoteWriteIdFactory ids, ICollection<FssHttpWriteObject> packageElements) {
        var declarations = new List<FssHttpWriteObject>();
        var data = new List<FssHttpWriteObject>();
        foreach (OneNoteWriteObject item in space.Objects) {
            byte[] jcid = Data(stream => FssHttpStreamObjectWriter.WriteUInt32(stream, item.Jcid));
            declarations.Add(Declaration(item.Id, 4, jcid.Length, 0, 0));
            data.Add(ObjectData(Array.Empty<OneNoteExtendedGuid>(), Array.Empty<FssHttpCellId>(), jcid));
            OneNoteEncodedPropertySet propertySet = OneNotePropertySetWriter.Write(
                item.Properties,
                space.Id,
                space.ContextId ?? DefaultContext);
            declarations.Add(Declaration(item.Id, 1, propertySet.Data.Length, propertySet.ObjectReferences.Count, propertySet.CellReferences.Count));
            data.Add(ObjectData(propertySet.ObjectReferences, propertySet.CellReferences, propertySet.Data));
            if (item.Blob != null) {
                OneNoteExtendedGuid blobElementId = ids.New();
                packageElements.Add(DataElement(blobElementId, 0x0A, new[] { Simple(0x02, item.Blob) }));
                declarations.Add(BlobDeclaration(item.Id, blobElementId));
                data.Add(ObjectBlobReference(blobElementId));
            }
        }
        return new[] {
            new FssHttpWriteObject(0x1D, children: declarations),
            new FssHttpWriteObject(0x1E, children: data)
        };
    }

    private static FssHttpWriteObject BlobDeclaration(OneNoteExtendedGuid id, OneNoteExtendedGuid blobElementId) {
        return Simple(0x05, Data(stream => {
            WriteExtended(stream, id);
            WriteExtended(stream, blobElementId);
            FssHttpStreamObjectWriter.WriteCompactUInt64(stream, 2);
            FssHttpStreamObjectWriter.WriteCompactUInt64(stream, 0);
            FssHttpStreamObjectWriter.WriteCompactUInt64(stream, 0);
        }));
    }

    private static FssHttpWriteObject ObjectBlobReference(OneNoteExtendedGuid blobElementId) {
        return Simple(0x1C, Data(stream => {
            FssHttpStreamObjectWriter.WriteCompactUInt64(stream, 0);
            FssHttpStreamObjectWriter.WriteCompactUInt64(stream, 0);
            WriteExtended(stream, blobElementId);
        }));
    }

    private static FssHttpWriteObject Declaration(OneNoteExtendedGuid id, ulong partition, int length, int objectReferences, int cellReferences) {
        return Simple(0x18, Data(stream => {
            WriteExtended(stream, id);
            FssHttpStreamObjectWriter.WriteCompactUInt64(stream, partition);
            FssHttpStreamObjectWriter.WriteCompactUInt64(stream, (ulong)length);
            FssHttpStreamObjectWriter.WriteCompactUInt64(stream, (ulong)objectReferences);
            FssHttpStreamObjectWriter.WriteCompactUInt64(stream, (ulong)cellReferences);
        }));
    }

    private static FssHttpWriteObject ObjectData(IReadOnlyList<OneNoteExtendedGuid> objectReferences, IReadOnlyList<FssHttpCellId> cellReferences, byte[] payload) {
        return Simple(0x16, Data(stream => {
            FssHttpStreamObjectWriter.WriteCompactUInt64(stream, (ulong)objectReferences.Count);
            foreach (OneNoteExtendedGuid id in objectReferences) WriteExtended(stream, id);
            FssHttpStreamObjectWriter.WriteCompactUInt64(stream, (ulong)cellReferences.Count);
            foreach (FssHttpCellId id in cellReferences) WriteCellId(stream, id);
            FssHttpStreamObjectWriter.WriteCompactUInt64(stream, (ulong)payload.Length);
            stream.Write(payload, 0, payload.Length);
        }));
    }

    private static FssHttpWriteObject DataElement(OneNoteExtendedGuid id, ulong type, IEnumerable<FssHttpWriteObject> children) {
        return new FssHttpWriteObject(0x01, Data(stream => { WriteExtended(stream, id); stream.WriteByte(0); FssHttpStreamObjectWriter.WriteCompactUInt64(stream, type); }), children);
    }
    private static FssHttpWriteObject Simple(int type, byte[] data) => new FssHttpWriteObject(type, data);
    private static byte[] Data(Action<Stream> writer) { using (var stream = new MemoryStream()) { writer(stream); return stream.ToArray(); } }
    private static void WriteExtended(Stream stream, OneNoteExtendedGuid value) => FssHttpStreamObjectWriter.WriteExtendedGuid(stream, value);
    private static void WriteCellId(Stream stream, FssHttpCellId value) { WriteExtended(stream, value.First); WriteExtended(stream, value.Second); }
    private static OneNoteExtendedGuid Extended(string guid, uint value) => new OneNoteExtendedGuid(new Guid(guid), value, 17);
    private static OneNoteExtendedGuid Null() => new OneNoteExtendedGuid(Guid.Empty, 0, 1);
    private static OneNoteWriteProperty DataProperty(uint id, byte[] value) => new OneNoteWriteProperty(id, data: value);
    private static OneNoteWriteProperty ScalarProperty(uint id, uint value) => new OneNoteWriteProperty(id, scalar: value);
    private static Guid Xor(Guid left, Guid right) {
        byte[] result = left.ToByteArray();
        byte[] other = right.ToByteArray();
        for (int index = 0; index < result.Length; index++) result[index] ^= other[index];
        return new Guid(result);
    }
}
