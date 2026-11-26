namespace OfficeIMO.OneNote;

internal static class OneNotePropertySetReader {
    public static OneNotePropertySet Read(
        byte[] data,
        Dictionary<uint, Guid> globalIds,
        OneNoteLoadOptions options,
        ulong absoluteOffset) => Read(new Cursor(data, globalIds, null, options, absoluteOffset));

    // Package CompactIDs are opaque mapping keys; the mapped value can exceed one byte.
    internal static OneNotePropertySet Read(
        byte[] data,
        Dictionary<uint, OneNoteExtendedGuid> compactIds,
        OneNoteLoadOptions options,
        ulong absoluteOffset) => Read(new Cursor(data, null, compactIds, options, absoluteOffset));

    // Checks a property set the way Read decodes it, raising the same errors, without allocating the decoded values.
    public static void Validate(
        byte[] data,
        Dictionary<uint, Guid> globalIds,
        OneNoteLoadOptions options,
        ulong absoluteOffset) {
        var cursor = new Cursor(data, globalIds, null, options, absoluteOffset);
        var counts = new ReferenceCounts { Oids = cursor.ScanReferenceStream(out _, out bool osidStreamNotPresent) };
        if (!osidStreamNotPresent) {
            counts.Osids = cursor.ScanReferenceStream(out bool osidExtendedStreamsPresent, out _);
            if (osidExtendedStreamsPresent) counts.Contexts = cursor.ScanReferenceStream(out _, out _);
        }
        cursor.ScanPropertySet(ref counts, 0);
    }

    private static OneNotePropertySet Read(Cursor cursor) {
        ReferenceStream oids = cursor.ReadReferenceStream();
        ReferenceStream osids = ReferenceStream.Empty;
        ReferenceStream contexts = ReferenceStream.Empty;
        if (!oids.OsidStreamNotPresent) {
            osids = cursor.ReadReferenceStream();
            if (osids.ExtendedStreamsPresent) contexts = cursor.ReadReferenceStream();
        }
        var counters = new ReferenceCounters(oids.Ids, osids.Ids, contexts.Ids);
        return cursor.ReadPropertySet(counters, 0);
    }

    private sealed class Cursor {
        private readonly byte[] _data;
        private readonly Dictionary<uint, Guid>? _globalIds;
        private readonly Dictionary<uint, OneNoteExtendedGuid>? _compactIds;
        private readonly OneNoteLoadOptions _options;
        private readonly ulong _absoluteOffset;
        private int _position;
        private int _totalProperties;

        public Cursor(byte[] data, Dictionary<uint, Guid>? globalIds, Dictionary<uint, OneNoteExtendedGuid>? compactIds, OneNoteLoadOptions options, ulong absoluteOffset) {
            _data = data;
            _globalIds = globalIds;
            _compactIds = compactIds;
            _options = options;
            _absoluteOffset = absoluteOffset;
        }

        public ReferenceStream ReadReferenceStream() {
            uint header = ReadUInt32();
            int count = (int)(header & 0x00FFFFFFU);
            if ((header & 0x3F000000U) != 0 || count > _options.MaxObjects) {
                throw Error("ONENOTE_OBJECT_STREAM_HEADER", "An object-reference stream header is invalid or exceeds the object limit.");
            }
            var ids = new List<OneNoteExtendedGuid>(count);
            for (int index = 0; index < count; index++) ids.Add(ReadCompactId());
            return new ReferenceStream(ids.AsReadOnly(), (header & 0x40000000U) != 0, (header & 0x80000000U) != 0);
        }

        public int ScanReferenceStream(out bool extendedStreamsPresent, out bool osidStreamNotPresent) {
            uint header = ReadUInt32();
            int count = (int)(header & 0x00FFFFFFU);
            if ((header & 0x3F000000U) != 0 || count > _options.MaxObjects) {
                throw Error("ONENOTE_OBJECT_STREAM_HEADER", "An object-reference stream header is invalid or exceeds the object limit.");
            }
            for (int index = 0; index < count; index++) ScanCompactId();
            extendedStreamsPresent = (header & 0x40000000U) != 0;
            osidStreamNotPresent = (header & 0x80000000U) != 0;
            return count;
        }

        // Mirrors ReadPropertySet step for step so positions, limits and errors match.
        public void ScanPropertySet(ref ReferenceCounts counts, int depth) {
            if (depth >= _options.MaxPropertySetDepth) {
                throw Error("ONENOTE_PROPERTY_DEPTH", "The nested property-set depth limit was exceeded.");
            }
            ushort count = ReadUInt16();
            if (count > _options.MaxPropertiesPerObject || _totalProperties > _options.MaxPropertiesPerObject - count) {
                throw Error("ONENOTE_PROPERTY_LIMIT", "The property count exceeds the configured per-object limit.");
            }
            _totalProperties += count;
            int rawIdStart = _position;
            for (int index = 0; index < count; index++) ReadUInt32();

            for (int index = 0; index < count; index++) {
                uint rawId = OneNoteBinary.ReadUInt32(_data, rawIdStart + index * 4);
                byte type = (byte)((rawId >> 26) & 0x1FU);
                switch (type) {
                    case 0x01:
                    case 0x02:
                        break;
                    case 0x03:
                        Skip(1);
                        break;
                    case 0x04:
                        Skip(2);
                        break;
                    case 0x05:
                        Skip(4);
                        break;
                    case 0x06:
                        Skip(8);
                        break;
                    case 0x07: {
                        uint length = ReadUInt32();
                        if (length >= 0x40000000U || length > int.MaxValue) {
                            throw Error("ONENOTE_PROPERTY_DATA_LENGTH", "A length-prefixed property value is too large.");
                        }
                        Skip((int)length);
                        break;
                    }
                    case 0x08:
                        ReferenceCounts.Take(counts.Oids, ref counts.OidIndex, 1, this, "object");
                        break;
                    case 0x09:
                        ReferenceCounts.Take(counts.Oids, ref counts.OidIndex, ReadReferenceCount(), this, "object");
                        break;
                    case 0x0A:
                        ReferenceCounts.Take(counts.Osids, ref counts.OsidIndex, 1, this, "object-space");
                        break;
                    case 0x0B:
                        ReferenceCounts.Take(counts.Osids, ref counts.OsidIndex, ReadReferenceCount(), this, "object-space");
                        break;
                    case 0x0C:
                        ReferenceCounts.Take(counts.Contexts, ref counts.ContextIndex, 1, this, "context");
                        break;
                    case 0x0D:
                        ReferenceCounts.Take(counts.Contexts, ref counts.ContextIndex, ReadReferenceCount(), this, "context");
                        break;
                    case 0x10: {
                        int childCount = ReadReferenceCount();
                        if (childCount == 0) break;
                        uint childPropertyId = ReadUInt32();
                        if (((childPropertyId >> 26) & 0x1FU) != 0x11U) {
                            throw Error("ONENOTE_PROPERTY_ARRAY_TYPE", "A property-set array does not declare PropertySet element values.");
                        }
                        for (int child = 0; child < childCount; child++) ScanPropertySet(ref counts, depth + 1);
                        break;
                    }
                    case 0x11:
                        ScanPropertySet(ref counts, depth + 1);
                        break;
                    default:
                        throw Error("ONENOTE_PROPERTY_TYPE", "The property set contains an unsupported representation type 0x" + type.ToString("X2", System.Globalization.CultureInfo.InvariantCulture) + ".");
                }
            }
        }

        private void Skip(int length) {
            Ensure(length);
            _position += length;
        }

        private void ScanCompactId() {
            uint compact = ReadUInt32();
            if (!_globalIds!.ContainsKey(compact >> 8)) {
                throw Error("ONENOTE_COMPACT_ID", "A property reference uses a missing global-identification table entry.");
            }
        }

        public OneNotePropertySet ReadPropertySet(ReferenceCounters counters, int depth) {
            if (depth >= _options.MaxPropertySetDepth) {
                throw Error("ONENOTE_PROPERTY_DEPTH", "The nested property-set depth limit was exceeded.");
            }
            int start = _position;
            ushort count = ReadUInt16();
            if (count > _options.MaxPropertiesPerObject || _totalProperties > _options.MaxPropertiesPerObject - count) {
                throw Error("ONENOTE_PROPERTY_LIMIT", "The property count exceeds the configured per-object limit.");
            }
            _totalProperties += count;
            var rawIds = new uint[count];
            for (int index = 0; index < rawIds.Length; index++) rawIds[index] = ReadUInt32();

            var properties = new OneNotePropertyValue[count];
            for (int index = 0; index < rawIds.Length; index++) {
                uint rawId = rawIds[index];
                var property = new OneNotePropertyValue(rawId, index);
                byte type = (byte)((rawId >> 26) & 0x1FU);
                switch (type) {
                    case 0x01:
                        break;
                    case 0x02:
                        property.BooleanValue = (rawId & 0x80000000U) != 0;
                        break;
                    case 0x03:
                        SetScalar(property, 1);
                        break;
                    case 0x04:
                        SetScalar(property, 2);
                        break;
                    case 0x05:
                        SetScalar(property, 4);
                        break;
                    case 0x06:
                        SetScalar(property, 8);
                        break;
                    case 0x07: {
                        uint length = ReadUInt32();
                        if (length >= 0x40000000U || length > int.MaxValue) {
                            throw Error("ONENOTE_PROPERTY_DATA_LENGTH", "A length-prefixed property value is too large.");
                        }
                        SetDataRange(property, (int)length);
                        break;
                    }
                    case 0x08:
                        property.ReferencedIds = counters.TakeOids(1, this);
                        break;
                    case 0x09:
                        property.ReferencedIds = counters.TakeOids(ReadReferenceCount(), this);
                        break;
                    case 0x0A:
                        property.ReferencedIds = counters.TakeOsids(1, this);
                        break;
                    case 0x0B:
                        property.ReferencedIds = counters.TakeOsids(ReadReferenceCount(), this);
                        break;
                    case 0x0C:
                        property.ReferencedIds = counters.TakeContexts(1, this);
                        break;
                    case 0x0D:
                        property.ReferencedIds = counters.TakeContexts(ReadReferenceCount(), this);
                        break;
                    case 0x10: {
                        int childCount = ReadReferenceCount();
                        if (childCount == 0) {
                            property.ChildPropertySets = Array.Empty<OneNotePropertySet>();
                            break;
                        }
                        uint childPropertyId = ReadUInt32();
                        if (((childPropertyId >> 26) & 0x1FU) != 0x11U) {
                            throw Error("ONENOTE_PROPERTY_ARRAY_TYPE", "A property-set array does not declare PropertySet element values.");
                        }
                        property.ChildPropertyId = childPropertyId;
                        var children = new List<OneNotePropertySet>(childCount);
                        for (int child = 0; child < childCount; child++) children.Add(ReadPropertySet(counters, depth + 1));
                        property.ChildPropertySets = children.AsReadOnly();
                        break;
                    }
                    case 0x11:
                        property.ChildPropertySets = new[] { ReadPropertySet(counters, depth + 1) };
                        break;
                    default:
                        throw Error("ONENOTE_PROPERTY_TYPE", "The property set contains an unsupported representation type 0x" + type.ToString("X2", System.Globalization.CultureInfo.InvariantCulture) + ".");
                }
                properties[index] = property;
            }
            return new OneNotePropertySet(new System.Collections.ObjectModel.ReadOnlyCollection<OneNotePropertyValue>(properties), _position - start);
        }

        public OneNoteFormatException Error(string code, string message) {
            ulong position = _absoluteOffset + (ulong)_position;
            long offset = position > long.MaxValue ? long.MaxValue : (long)position;
            return new OneNoteFormatException(code, message, offset);
        }

        private void SetScalar(OneNotePropertyValue property, int byteCount) {
            Ensure(byteCount);
            ulong value = 0;
            for (int index = 0; index < byteCount; index++) value |= (ulong)_data[_position + index] << (index * 8);
            property.ScalarValue = value;
            SetDataRange(property, byteCount);
        }

        private int ReadReferenceCount() {
            uint count = ReadUInt32();
            if (count > _options.MaxObjects || count > int.MaxValue) {
                throw Error("ONENOTE_REFERENCE_COUNT", "A property reference count exceeds the configured object limit.");
            }
            return (int)count;
        }

        private OneNoteExtendedGuid ReadCompactId() {
            uint compact = ReadUInt32();
            if (_compactIds != null) {
                if (_compactIds.TryGetValue(compact, out OneNoteExtendedGuid? mapped)) {
                    return new OneNoteExtendedGuid(mapped.Identifier, mapped.Value, 4);
                }
                if (compact == 0) return new OneNoteExtendedGuid(Guid.Empty, 0, 4);
                throw Error("ONENOTE_COMPACT_ID", "A property reference uses a missing package mapping entry.");
            }
            byte value = (byte)(compact & 0xFFU);
            uint index = compact >> 8;
            if (!_globalIds!.TryGetValue(index, out Guid identifier)) {
                throw Error("ONENOTE_COMPACT_ID", "A property reference uses a missing global-identification table entry.");
            }
            return new OneNoteExtendedGuid(identifier, value, 4);
        }

        private ushort ReadUInt16() {
            Ensure(2);
            ushort value = OneNoteBinary.ReadUInt16(_data, _position);
            _position += 2;
            return value;
        }

        private uint ReadUInt32() {
            Ensure(4);
            uint value = OneNoteBinary.ReadUInt32(_data, _position);
            _position += 4;
            return value;
        }

        // The object data buffer is owned by the reader and never modified, so values share it instead of copying.
        private void SetDataRange(OneNotePropertyValue property, int length) {
            Ensure(length);
            property.SetDataRange(_data, _position, length);
            _position += length;
        }

        private void Ensure(int length) {
            if (length < 0 || _position > _data.Length - length) {
                throw Error("ONENOTE_TRUNCATED_PROPERTY_SET", "The object data ended inside a property set.");
            }
        }
    }

    private struct ReferenceCounts {
        public int Oids;
        public int Osids;
        public int Contexts;
        public int OidIndex;
        public int OsidIndex;
        public int ContextIndex;

        public static void Take(int available, ref int index, int count, Cursor cursor, string kind) {
            if (count < 0 || index > available - count) {
                throw cursor.Error("ONENOTE_REFERENCE_STREAM", "A property consumes more " + kind + " identifiers than its object stream contains.");
            }
            index += count;
        }
    }

    private sealed class ReferenceCounters {
        private readonly IReadOnlyList<OneNoteExtendedGuid> _oids;
        private readonly IReadOnlyList<OneNoteExtendedGuid> _osids;
        private readonly IReadOnlyList<OneNoteExtendedGuid> _contexts;
        private int _oidIndex;
        private int _osidIndex;
        private int _contextIndex;

        public ReferenceCounters(
            IReadOnlyList<OneNoteExtendedGuid> oids,
            IReadOnlyList<OneNoteExtendedGuid> osids,
            IReadOnlyList<OneNoteExtendedGuid> contexts) {
            _oids = oids;
            _osids = osids;
            _contexts = contexts;
        }

        public IReadOnlyList<OneNoteExtendedGuid> TakeOids(int count, Cursor cursor) => Take(_oids, ref _oidIndex, count, cursor, "object");
        public IReadOnlyList<OneNoteExtendedGuid> TakeOsids(int count, Cursor cursor) => Take(_osids, ref _osidIndex, count, cursor, "object-space");
        public IReadOnlyList<OneNoteExtendedGuid> TakeContexts(int count, Cursor cursor) => Take(_contexts, ref _contextIndex, count, cursor, "context");

        private static IReadOnlyList<OneNoteExtendedGuid> Take(
            IReadOnlyList<OneNoteExtendedGuid> source,
            ref int index,
            int count,
            Cursor cursor,
            string kind) {
            if (count < 0 || index > source.Count - count) {
                throw cursor.Error("ONENOTE_REFERENCE_STREAM", "A property consumes more " + kind + " identifiers than its object stream contains.");
            }
            var result = new OneNoteExtendedGuid[count];
            for (int item = 0; item < count; item++) result[item] = source[index++];
            return new System.Collections.ObjectModel.ReadOnlyCollection<OneNoteExtendedGuid>(result);
        }
    }

    private sealed class ReferenceStream {
        public static readonly ReferenceStream Empty = new ReferenceStream(Array.Empty<OneNoteExtendedGuid>(), false, true);

        public ReferenceStream(IReadOnlyList<OneNoteExtendedGuid> ids, bool extendedStreamsPresent, bool osidStreamNotPresent) {
            Ids = ids;
            ExtendedStreamsPresent = extendedStreamsPresent;
            OsidStreamNotPresent = osidStreamNotPresent;
        }

        public IReadOnlyList<OneNoteExtendedGuid> Ids { get; }
        public bool ExtendedStreamsPresent { get; }
        public bool OsidStreamNotPresent { get; }
    }
}

// Decodes the property sets declared under one global-identification table when they are first read.
internal sealed class OneNotePropertySetDecoder {
    private readonly Dictionary<uint, Guid> _globalIds;
    private readonly OneNotePropertySetDecodeScope _scope;

    internal OneNotePropertySetDecoder(Dictionary<uint, Guid> globalIds, OneNotePropertySetDecodeScope scope) {
        _globalIds = globalIds;
        _scope = scope;
    }

    internal OneNotePropertySet Decode(OneNoteRevisionStoreObject owner, byte[] data) {
        OneNotePropertySet propertySet = OneNotePropertySetReader.Read(data, _globalIds, _scope.Options, owner.DeclarationNode.ChunkOffset);
        _scope.Track(owner);
        return propertySet;
    }
}

// Shared by the lazily decoded objects of one store. While the loader maps pages it records which objects were
// decoded so it can drop their property sets once a page is done.
internal sealed class OneNotePropertySetDecodeScope {
    private List<OneNoteRevisionStoreObject>? _decoded;

    internal OneNotePropertySetDecodeScope(OneNoteLoadOptions options) {
        // Decoding can happen after the read returns, so keep the limits the objects were validated against
        // rather than the caller's instance, which may change later.
        Options = new OneNoteLoadOptions {
            MaxObjects = options.MaxObjects,
            MaxPropertiesPerObject = options.MaxPropertiesPerObject,
            MaxPropertySetDepth = options.MaxPropertySetDepth
        };
    }

    internal OneNoteLoadOptions Options { get; }

    internal void BeginTracking() => _decoded ??= new List<OneNoteRevisionStoreObject>();

    internal void Track(OneNoteRevisionStoreObject owner) => _decoded?.Add(owner);

    internal void ReleaseTracked() {
        if (_decoded == null) return;
        foreach (OneNoteRevisionStoreObject owner in _decoded) owner.ReleaseDecodedPropertySet();
        _decoded.Clear();
    }

    internal void EndTracking() {
        ReleaseTracked();
        _decoded = null;
    }
}
