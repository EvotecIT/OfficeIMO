namespace OfficeIMO.OneNote;

internal sealed class OneNoteEncodedPropertySet {
    internal OneNoteEncodedPropertySet(byte[] data, IReadOnlyList<OneNoteExtendedGuid> objectReferences, IReadOnlyList<FssHttpCellId> cellReferences) {
        Data = data; ObjectReferences = objectReferences; CellReferences = cellReferences;
    }
    internal byte[] Data { get; }
    internal IReadOnlyList<OneNoteExtendedGuid> ObjectReferences { get; }
    internal IReadOnlyList<FssHttpCellId> CellReferences { get; }
}

internal static class OneNotePropertySetWriter {
    internal static OneNoteEncodedPropertySet Write(
        IReadOnlyList<OneNoteWriteProperty> properties,
        OneNoteExtendedGuid currentObjectSpaceId,
        OneNoteExtendedGuid currentContextId) {
        OneNoteWriteProperty[] flattened = EnumerateProperties(properties).ToArray();
        var objectReferences = flattened.Where(item => item.ReferenceKind == OneNoteWriteReferenceKind.Object).SelectMany(item => item.References).ToArray();
        var objectSpaceReferences = flattened.Where(item => item.ReferenceKind == OneNoteWriteReferenceKind.ObjectSpace).SelectMany(item => item.References).ToArray();
        var contextReferences = flattened.Where(item => item.ReferenceKind == OneNoteWriteReferenceKind.Context).SelectMany(item => item.References).ToArray();
        var cellReferences = objectSpaceReferences.Select(id => new FssHttpCellId(currentContextId, id))
            .Concat(contextReferences.Select(id => new FssHttpCellId(id, currentObjectSpaceId)))
            .ToArray();
        using (var stream = new MemoryStream()) {
            bool hasObjectSpaceStream = objectSpaceReferences.Length > 0 || contextReferences.Length > 0;
            WriteReferenceStream(stream, objectReferences, 0, false, !hasObjectSpaceStream);
            if (hasObjectSpaceStream) {
                WriteReferenceStream(stream, objectSpaceReferences, objectReferences.Length, contextReferences.Length > 0, false);
            }
            if (contextReferences.Length > 0) {
                WriteReferenceStream(stream, contextReferences, objectReferences.Length + objectSpaceReferences.Length, false, false);
            }
            WritePropertySet(stream, properties, 0);
            return new OneNoteEncodedPropertySet(stream.ToArray(), objectReferences, cellReferences);
        }
    }

    internal static byte[] WriteDesktop(IReadOnlyList<OneNoteWriteProperty> properties, IReadOnlyDictionary<Guid, uint> globalIds) {
        if (globalIds == null) throw new ArgumentNullException(nameof(globalIds));
        OneNoteWriteProperty[] flattened = EnumerateProperties(properties).ToArray();
        OneNoteExtendedGuid[] objectReferences = flattened.Where(item => item.ReferenceKind == OneNoteWriteReferenceKind.Object).SelectMany(item => item.References).ToArray();
        OneNoteExtendedGuid[] objectSpaceReferences = flattened.Where(item => item.ReferenceKind == OneNoteWriteReferenceKind.ObjectSpace).SelectMany(item => item.References).ToArray();
        OneNoteExtendedGuid[] contextReferences = flattened.Where(item => item.ReferenceKind == OneNoteWriteReferenceKind.Context).SelectMany(item => item.References).ToArray();
        using (var stream = new MemoryStream()) {
            bool hasObjectSpaceStream = objectSpaceReferences.Length > 0 || contextReferences.Length > 0;
            WriteDesktopReferenceStream(stream, objectReferences, globalIds, false, !hasObjectSpaceStream);
            if (hasObjectSpaceStream) {
                WriteDesktopReferenceStream(stream, objectSpaceReferences, globalIds, contextReferences.Length > 0, false);
            }
            if (contextReferences.Length > 0) {
                WriteDesktopReferenceStream(stream, contextReferences, globalIds, false, false);
            }
            WritePropertySet(stream, properties, 0);
            return stream.ToArray();
        }
    }

    private static void WriteDesktopReferenceStream(
        Stream stream,
        IReadOnlyList<OneNoteExtendedGuid> references,
        IReadOnlyDictionary<Guid, uint> globalIds,
        bool extendedStreamsPresent,
        bool osidStreamNotPresent) {
        uint header = (uint)references.Count |
            (extendedStreamsPresent ? 0x40000000U : 0U) |
            (osidStreamNotPresent ? 0x80000000U : 0U);
        FssHttpStreamObjectWriter.WriteUInt32(stream, header);
        foreach (OneNoteExtendedGuid id in references) {
            if (!globalIds.TryGetValue(id.Identifier, out uint index) || index >= 0xFFFFFFU || id.Value > byte.MaxValue) {
                throw new OneNoteFormatException("ONENOTE_WRITE_COMPACT_ID", "A property reference is not present in the desktop global-identification table.");
            }
            FssHttpStreamObjectWriter.WriteUInt32(stream, (index << 8) | id.Value);
        }
    }

    private static void WriteReferenceStream(
        Stream stream,
        IReadOnlyList<OneNoteExtendedGuid> references,
        int globalIndexOffset,
        bool extendedStreamsPresent,
        bool osidStreamNotPresent) {
        uint header = (uint)references.Count |
            (extendedStreamsPresent ? 0x40000000U : 0U) |
            (osidStreamNotPresent ? 0x80000000U : 0U);
        FssHttpStreamObjectWriter.WriteUInt32(stream, header);
        for (int index = 0; index < references.Count; index++) {
            OneNoteExtendedGuid id = references[index];
            if (id.Value > byte.MaxValue || globalIndexOffset + index >= 0xFFFFFF) {
                throw new OneNoteFormatException("ONENOTE_WRITE_COMPACT_ID", "A writer-generated object identity cannot be represented as a CompactID.");
            }
            uint compact = ((uint)(globalIndexOffset + index) << 8) | id.Value;
            FssHttpStreamObjectWriter.WriteUInt32(stream, compact);
        }
    }

    private static void WritePropertySet(Stream stream, IReadOnlyList<OneNoteWriteProperty> properties, int depth) {
        if (depth >= 64) throw new OneNoteFormatException("ONENOTE_WRITE_PROPERTY_DEPTH", "A nested property-set depth exceeds the writer limit.");
        if (properties.Count > ushort.MaxValue) throw new OneNoteFormatException("ONENOTE_WRITE_PROPERTY_LIMIT", "An object contains too many properties to serialize.");
        FssHttpStreamObjectWriter.WriteUInt16(stream, (ushort)properties.Count);
        foreach (OneNoteWriteProperty property in properties) FssHttpStreamObjectWriter.WriteUInt32(stream, property.RawId);
        foreach (OneNoteWriteProperty property in properties) WriteValue(stream, property, depth);
    }

    private static void WriteValue(Stream stream, OneNoteWriteProperty property, int depth) {
        var type = (OneNotePropertyType)((property.RawId >> 26) & 0x1FU);
        switch (type) {
            case OneNotePropertyType.NoData:
            case OneNotePropertyType.Boolean:
            case OneNotePropertyType.ObjectId:
            case OneNotePropertyType.ObjectSpaceId:
            case OneNotePropertyType.ContextId:
                return;
            case OneNotePropertyType.ObjectIdArray:
            case OneNotePropertyType.ObjectSpaceIdArray:
            case OneNotePropertyType.ContextIdArray:
                FssHttpStreamObjectWriter.WriteUInt32(stream, (uint)property.References.Count);
                return;
            case OneNotePropertyType.Byte:
                stream.WriteByte((byte)(property.Scalar ?? 0));
                return;
            case OneNotePropertyType.UInt16:
                FssHttpStreamObjectWriter.WriteUInt16(stream, (ushort)(property.Scalar ?? 0));
                return;
            case OneNotePropertyType.UInt32:
                FssHttpStreamObjectWriter.WriteUInt32(stream, (uint)(property.Scalar ?? 0));
                return;
            case OneNotePropertyType.UInt64:
                FssHttpStreamObjectWriter.WriteUInt64(stream, property.Scalar ?? 0);
                return;
            case OneNotePropertyType.LengthPrefixedData: {
                byte[] data = property.Data ?? Array.Empty<byte>();
                FssHttpStreamObjectWriter.WriteUInt32(stream, (uint)data.Length);
                stream.Write(data, 0, data.Length);
                return;
            }
            case OneNotePropertyType.PropertySet:
                if (property.ChildPropertySets.Count != 1) {
                    throw new OneNoteFormatException("ONENOTE_WRITE_CHILD_PROPERTY_SET", "A PropertySet value must contain exactly one child property set.");
                }
                WritePropertySet(stream, property.ChildPropertySets[0], depth + 1);
                return;
            case OneNotePropertyType.PropertySetArray:
                FssHttpStreamObjectWriter.WriteUInt32(stream, (uint)property.ChildPropertySets.Count);
                if (property.ChildPropertySets.Count == 0) return;
                if (!property.ChildPropertyId.HasValue || ((property.ChildPropertyId.Value >> 26) & 0x1FU) != (uint)OneNotePropertyType.PropertySet) {
                    throw new OneNoteFormatException("ONENOTE_WRITE_CHILD_PROPERTY_ID", "A property-set array requires a PropertySet element identifier.");
                }
                FssHttpStreamObjectWriter.WriteUInt32(stream, property.ChildPropertyId.Value);
                foreach (IReadOnlyList<OneNoteWriteProperty> child in property.ChildPropertySets) WritePropertySet(stream, child, depth + 1);
                return;
            default:
                throw new OneNoteFormatException("ONENOTE_WRITE_PROPERTY_TYPE", "The writer does not support property representation " + type + ".");
        }
    }

    internal static IEnumerable<OneNoteWriteProperty> EnumerateProperties(IReadOnlyList<OneNoteWriteProperty> properties) {
        foreach (OneNoteWriteProperty property in properties) {
            yield return property;
            foreach (IReadOnlyList<OneNoteWriteProperty> child in property.ChildPropertySets) {
                foreach (OneNoteWriteProperty nested in EnumerateProperties(child)) yield return nested;
            }
        }
    }
}
