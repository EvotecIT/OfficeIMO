namespace OfficeIMO.OneNote;

internal static partial class OneNotePackageStoreReader {
    private static void ReadObjectGroup(
        Stream stream,
        PackageDataElement group,
        PackageRevisionElement revision,
        IReadOnlyDictionary<PackageGuidKey, PackageDataElement> blobs,
        PackageGraph graph,
        HashSet<PackageGuidKey> seenObjects,
        OneNoteReaderOptions options) {
        FssHttpStreamObject declarationsNode = RequireExactlyOne(
            group.Node.Children,
            item => item.Type == 0x1D,
            "ONENOTE_PACKAGE_OBJECT_DECLARATIONS",
            "An object group must contain exactly one declarations container.",
            group.Node.HeaderOffset);
        FssHttpStreamObject dataNode = RequireExactlyOne(
            group.Node.Children,
            item => item.Type == 0x1E,
            "ONENOTE_PACKAGE_OBJECT_DATA",
            "An object group must contain exactly one data container.",
            group.Node.HeaderOffset);
        PackageObjectDeclaration[] declarations = declarationsNode.Children
            .Where(item => item.Type == 0x18 || item.Type == 0x05)
            .Select(item => ReadDeclaration(stream, item, options))
            .ToArray();
        FssHttpStreamObject[] dataItems = dataNode.Children
            .Where(item => item.Type == 0x16 || item.Type == 0x1C)
            .ToArray();
        if (declarations.Length != dataItems.Length) {
            throw new OneNoteFormatException("ONENOTE_PACKAGE_OBJECT_PAIRING", "An object group has unequal declaration and data counts.", group.Node.HeaderOffset);
        }

        var objects = new Dictionary<PackageGuidKey, PackageObjectAccumulator>();
        var order = new List<PackageObjectAccumulator>();
        for (int index = 0; index < declarations.Length; index++) {
            PackageObjectDeclaration declaration = declarations[index];
            FssHttpStreamObject dataItem = dataItems[index];
            PackageObjectData data = ReadObjectData(stream, dataItem, options);
            if (declaration.ObjectReferences != (ulong)data.ObjectReferences.Count || declaration.CellReferences != (ulong)data.CellReferences.Count) {
                throw new OneNoteFormatException("ONENOTE_PACKAGE_REFERENCE_COUNT", "An object declaration does not match its data reference counts.", declaration.Offset);
            }
            if (declaration.DataSize.HasValue && declaration.DataSize.Value != (ulong)(data.Data?.Length ?? 0)) {
                throw new OneNoteFormatException("ONENOTE_PACKAGE_OBJECT_SIZE", "An object declaration does not match its object-data size.", declaration.Offset);
            }
            PackageGuidKey key = new PackageGuidKey(declaration.ObjectId);
            if (!objects.TryGetValue(key, out PackageObjectAccumulator? accumulator)) {
                accumulator = new PackageObjectAccumulator(declaration.ObjectId, declaration.Offset);
                objects.Add(key, accumulator);
                order.Add(accumulator);
            }
            accumulator.ReferenceCount = Math.Max(accumulator.ReferenceCount, ToReferenceCount(declaration.ObjectReferences, declaration.CellReferences));
            if (declaration.PartitionId == 4) {
                if (data.Data == null || data.Data.Length != 4) {
                    throw new OneNoteFormatException("ONENOTE_PACKAGE_JCID", "Static object metadata does not contain a four-byte JCID.", dataItem.DataOffset);
                }
                accumulator.Jcid = OneNoteBinary.ReadUInt32(data.Data, 0);
            } else if (declaration.PartitionId == 1) {
                accumulator.PropertyData = data.Data;
                accumulator.ObjectReferences = data.ObjectReferences;
                accumulator.CellReferences = data.CellReferences;
            } else if (declaration.PartitionId == 2) {
                accumulator.BlobElementId = data.BlobElementId ?? declaration.BlobElementId;
            }
        }

        foreach (PackageObjectAccumulator accumulator in order) {
            if (graph.Objects.Count >= options.MaxObjects) {
                throw new OneNoteFormatException("ONENOTE_OBJECT_LIMIT", "The object declaration limit was exceeded.", accumulator.Offset);
            }
            var jcid = new OneNoteJcid(accumulator.Jcid);
            var declarationNode = new OneNoteFileNode(0, 0, 0, 0, OneNoteFileNodeBaseType.Inline, accumulator.Offset, null, Array.Empty<byte>());
            var record = new OneNoteRevisionStoreObject(accumulator.Id, jcid, declarationNode) {
                ReferenceCount = accumulator.ReferenceCount,
                RevisionId = revision.Manifest.Id,
                IsRevision = !seenObjects.Add(new PackageGuidKey(accumulator.Id))
            };
            // File-data objects carry an ObjectSpaceObjectPropSet in FSSHTTP even
            // though their JCID does not set the IsPropertySet flag.
            if (accumulator.PropertyData != null) {
                if (!revision.CellId.HasValue) {
                    throw new OneNoteFormatException("ONENOTE_PACKAGE_OBJECT_SPACE", "A property-set object is not associated with a storage cell.", accumulator.Offset);
                }
                Dictionary<uint, Guid> mappings = BuildGlobalIdMappings(
                    accumulator.PropertyData,
                    accumulator.ObjectReferences,
                    accumulator.CellReferences,
                    revision.CellId.Value,
                    options,
                    accumulator.Offset);
                record.RawPropertyData = OneNoteBinaryPayload.FromBytes(accumulator.PropertyData);
                record.PropertySet = OneNotePropertySetReader.Read(accumulator.PropertyData, mappings, options, (ulong)Math.Max(0, accumulator.Offset));
                ApplyFileDataProperties(record);
            }
            graph.Objects.Add(record);

            if (jcid.IsFileData && accumulator.BlobElementId != null && record.FileDataReference != null &&
                TryParseFileDataGuid(record.FileDataReference, out Guid fileDataId) &&
                !graph.FileDataObjects.Any(item => item.Id == fileDataId) &&
                blobs.TryGetValue(new PackageGuidKey(accumulator.BlobElementId), out PackageDataElement? blobElement)) {
                FssHttpStreamObject blobNode = RequireExactlyOne(
                    blobElement.Node.Children,
                    item => item.Type == 0x02,
                    "ONENOTE_PACKAGE_BLOB",
                    "An object-data BLOB element must contain exactly one BLOB payload.",
                    blobElement.Node.HeaderOffset);
                if (blobNode.DataLength > (ulong)options.MaxAssetBytes ||
                    graph.TotalAssetBytes > options.MaxTotalAssetBytes - (long)blobNode.DataLength) {
                    throw new OneNoteFormatException("ONENOTE_ASSET_LIMIT", "An embedded OneNote asset exceeds the configured materialization limits.", blobNode.DataOffset);
                }
                byte[] bytes = FssHttpStreamObjectReader.ReadData(stream, blobNode, (ulong)options.MaxAssetBytes, "object-data BLOB");
                graph.TotalAssetBytes += bytes.LongLength;
                graph.FileDataObjects.Add(new OneNoteFileDataStoreObject(fileDataId, OneNoteBinaryPayload.FromBytes(bytes)));
            }
        }
    }

    private static PackageObjectDeclaration ReadDeclaration(Stream stream, FssHttpStreamObject item, OneNoteReaderOptions options) {
        byte[] data = FssHttpStreamObjectReader.ReadData(stream, item, 256, "object declaration");
        var cursor = new FssHttpDataCursor(data, item.DataOffset);
        OneNoteExtendedGuid objectId = cursor.ReadExtendedGuid();
        OneNoteExtendedGuid? blobId = item.Type == 0x05 ? cursor.ReadExtendedGuid() : null;
        ulong partition = cursor.ReadCompactUInt64();
        ulong? dataSize = item.Type == 0x18 ? cursor.ReadCompactUInt64() : (ulong?)null;
        ulong objectReferences = cursor.ReadCompactUInt64();
        ulong cellReferences = cursor.ReadCompactUInt64();
        cursor.EnsureEnd("object declaration");
        if (partition > uint.MaxValue || objectReferences > (ulong)options.MaxObjects || cellReferences > (ulong)options.MaxObjects) {
            throw new OneNoteFormatException("ONENOTE_PACKAGE_OBJECT_DECLARATION", "An object declaration contains an unsupported partition or excessive reference count.", item.DataOffset);
        }
        return new PackageObjectDeclaration(objectId, blobId, partition, dataSize, objectReferences, cellReferences, item.DataOffset);
    }

    private static PackageObjectData ReadObjectData(Stream stream, FssHttpStreamObject item, OneNoteReaderOptions options) {
        ulong max = options.MaxInputBytes.HasValue ? (ulong)Math.Min(options.MaxInputBytes.Value, int.MaxValue) : int.MaxValue;
        byte[] raw = FssHttpStreamObjectReader.ReadData(stream, item, max, "object data");
        var cursor = new FssHttpDataCursor(raw, item.DataOffset);
        IReadOnlyList<OneNoteExtendedGuid> objectReferences = cursor.ReadExtendedGuidArray(options.MaxObjects);
        IReadOnlyList<FssHttpCellId> cellReferences = cursor.ReadCellIdArray(options.MaxObjects);
        if (item.Type == 0x16) {
            byte[] data = cursor.ReadBinaryItem((long)max);
            cursor.EnsureEnd("object data");
            return new PackageObjectData(objectReferences, cellReferences, data, null);
        }
        OneNoteExtendedGuid blobId = cursor.ReadExtendedGuid();
        cursor.EnsureEnd("object-data BLOB reference");
        return new PackageObjectData(objectReferences, cellReferences, null, blobId);
    }

    private static Dictionary<uint, Guid> BuildGlobalIdMappings(
        byte[] propertyData,
        IReadOnlyList<OneNoteExtendedGuid> objectReferences,
        IReadOnlyList<FssHttpCellId> cellReferences,
        FssHttpCellId currentCell,
        OneNoteReaderOptions options,
        long offset) {
        int position = 0;
        PackageCompactStream oids = ReadCompactStream(propertyData, ref position, options, offset);
        PackageCompactStream osids = PackageCompactStream.Empty;
        PackageCompactStream contexts = PackageCompactStream.Empty;
        if (!oids.OsidStreamNotPresent) {
            osids = ReadCompactStream(propertyData, ref position, options, offset);
            if (osids.ExtendedStreamsPresent) contexts = ReadCompactStream(propertyData, ref position, options, offset);
        }

        OneNoteExtendedGuid[] osidReferences = cellReferences.Where(cell => cell.First.Equals(currentCell.First)).Select(cell => cell.Second).ToArray();
        OneNoteExtendedGuid[] contextReferences = cellReferences.Where(cell => !cell.First.Equals(currentCell.First)).Select(cell => cell.First).ToArray();
        var mappings = new Dictionary<uint, Guid>();
        AddMappings(mappings, oids.CompactIds, objectReferences, offset, "object");
        AddMappings(mappings, osids.CompactIds, osidReferences, offset, "object-space");
        AddMappings(mappings, contexts.CompactIds, contextReferences, offset, "context");
        return mappings;
    }

    private static PackageCompactStream ReadCompactStream(byte[] data, ref int position, OneNoteReaderOptions options, long offset) {
        if (position > data.Length - 4) throw new OneNoteFormatException("ONENOTE_PACKAGE_OBJECT_STREAM", "A property reference stream is truncated.", offset + position);
        uint header = OneNoteBinary.ReadUInt32(data, position);
        position += 4;
        int count = (int)(header & 0x00FFFFFFU);
        if ((header & 0x3F000000U) != 0 || count > options.MaxObjects || position > data.Length - checked(count * 4)) {
            throw new OneNoteFormatException("ONENOTE_PACKAGE_OBJECT_STREAM", "A property reference stream is invalid or exceeds configured limits.", offset + position - 4);
        }
        var ids = new uint[count];
        for (int index = 0; index < count; index++) {
            ids[index] = OneNoteBinary.ReadUInt32(data, position);
            position += 4;
        }
        return new PackageCompactStream(ids, (header & 0x40000000U) != 0, (header & 0x80000000U) != 0);
    }

    private static void AddMappings(
        Dictionary<uint, Guid> mappings,
        IReadOnlyList<uint> compactIds,
        IReadOnlyList<OneNoteExtendedGuid> extendedIds,
        long offset,
        string kind) {
        if (compactIds.Count != extendedIds.Count) {
            throw new OneNoteFormatException("ONENOTE_PACKAGE_MAPPING_COUNT", "The " + kind + " CompactID and extended-GUID arrays have different counts.", offset);
        }
        for (int index = 0; index < compactIds.Count; index++) {
            uint compact = compactIds[index];
            OneNoteExtendedGuid extended = extendedIds[index];
            if (compact == 0 && IsNull(extended)) continue;
            uint globalIndex = compact >> 8;
            uint value = compact & 0xFFU;
            if (globalIndex >= 0xFFFFFFU || extended.Identifier == Guid.Empty || extended.Value != value) {
                throw new OneNoteFormatException("ONENOTE_PACKAGE_MAPPING", "A " + kind + " mapping contains incompatible CompactID and extended-GUID values.", offset);
            }
            if (mappings.TryGetValue(globalIndex, out Guid existing) && existing != extended.Identifier) {
                throw new OneNoteFormatException("ONENOTE_PACKAGE_MAPPING", "A CompactID global index maps to multiple GUIDs.", offset);
            }
            mappings[globalIndex] = extended.Identifier;
        }
    }

    private static void ApplyFileDataProperties(OneNoteRevisionStoreObject record) {
        if (!record.Jcid.IsFileData || record.PropertySet == null) return;
        OneNotePropertyValue? guidProperty = record.PropertySet.Find(0x1C00343E);
        if (guidProperty?.Data != null) {
            byte[] bytes = guidProperty.Data.ToArray(64);
            if (bytes.Length == 16) record.FileDataReference = "<ifndf>" + new Guid(bytes).ToString("B");
        }
        OneNotePropertyValue? extensionProperty = record.PropertySet.Find(0x1C003424);
        if (extensionProperty?.Data != null) {
            byte[] bytes = extensionProperty.Data.ToArray(4096);
            if ((bytes.Length & 1) == 0) record.FileExtension = System.Text.Encoding.Unicode.GetString(bytes).TrimEnd('\0');
        }
    }

    private static bool TryParseFileDataGuid(string value, out Guid id) {
        string text = value.StartsWith("<ifndf>", StringComparison.OrdinalIgnoreCase) ? value.Substring(7) : value;
        return Guid.TryParse(text.Trim().TrimEnd('\0'), out id);
    }

    private static uint ToReferenceCount(ulong objectReferences, ulong cellReferences) {
        ulong total = objectReferences + cellReferences;
        return total > uint.MaxValue ? uint.MaxValue : (uint)total;
    }

    private sealed class PackageObjectDeclaration {
        internal PackageObjectDeclaration(OneNoteExtendedGuid objectId, OneNoteExtendedGuid? blobElementId, ulong partitionId, ulong? dataSize, ulong objectReferences, ulong cellReferences, long offset) {
            ObjectId = objectId; BlobElementId = blobElementId; PartitionId = partitionId; DataSize = dataSize;
            ObjectReferences = objectReferences; CellReferences = cellReferences; Offset = offset;
        }
        internal OneNoteExtendedGuid ObjectId { get; }
        internal OneNoteExtendedGuid? BlobElementId { get; }
        internal ulong PartitionId { get; }
        internal ulong? DataSize { get; }
        internal ulong ObjectReferences { get; }
        internal ulong CellReferences { get; }
        internal long Offset { get; }
    }

    private sealed class PackageObjectData {
        internal PackageObjectData(IReadOnlyList<OneNoteExtendedGuid> objectReferences, IReadOnlyList<FssHttpCellId> cellReferences, byte[]? data, OneNoteExtendedGuid? blobElementId) {
            ObjectReferences = objectReferences; CellReferences = cellReferences; Data = data; BlobElementId = blobElementId;
        }
        internal IReadOnlyList<OneNoteExtendedGuid> ObjectReferences { get; }
        internal IReadOnlyList<FssHttpCellId> CellReferences { get; }
        internal byte[]? Data { get; }
        internal OneNoteExtendedGuid? BlobElementId { get; }
    }

    private sealed class PackageObjectAccumulator {
        internal PackageObjectAccumulator(OneNoteExtendedGuid id, long offset) { Id = id; Offset = offset; }
        internal OneNoteExtendedGuid Id { get; }
        internal long Offset { get; }
        internal uint Jcid { get; set; }
        internal uint ReferenceCount { get; set; }
        internal byte[]? PropertyData { get; set; }
        internal IReadOnlyList<OneNoteExtendedGuid> ObjectReferences { get; set; } = Array.Empty<OneNoteExtendedGuid>();
        internal IReadOnlyList<FssHttpCellId> CellReferences { get; set; } = Array.Empty<FssHttpCellId>();
        internal OneNoteExtendedGuid? BlobElementId { get; set; }
    }

    private sealed class PackageCompactStream {
        internal static readonly PackageCompactStream Empty = new PackageCompactStream(Array.Empty<uint>(), false, true);
        internal PackageCompactStream(IReadOnlyList<uint> compactIds, bool extendedStreamsPresent, bool osidStreamNotPresent) {
            CompactIds = compactIds; ExtendedStreamsPresent = extendedStreamsPresent; OsidStreamNotPresent = osidStreamNotPresent;
        }
        internal IReadOnlyList<uint> CompactIds { get; }
        internal bool ExtendedStreamsPresent { get; }
        internal bool OsidStreamNotPresent { get; }
    }
}
