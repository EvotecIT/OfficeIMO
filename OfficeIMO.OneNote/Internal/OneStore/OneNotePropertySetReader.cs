namespace OfficeIMO.OneNote;

internal static class OneNotePropertySetReader {
    public static OneNotePropertySet Read(
        byte[] data,
        Dictionary<uint, Guid> globalIds,
        OneNoteReaderOptions options,
        ulong absoluteOffset) {
        var cursor = new Cursor(data, globalIds, options, absoluteOffset);
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
        private readonly Dictionary<uint, Guid> _globalIds;
        private readonly OneNoteReaderOptions _options;
        private readonly ulong _absoluteOffset;
        private int _position;
        private int _totalProperties;

        public Cursor(byte[] data, Dictionary<uint, Guid> globalIds, OneNoteReaderOptions options, ulong absoluteOffset) {
            _data = data;
            _globalIds = globalIds;
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

            var properties = new List<OneNotePropertyValue>(count);
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
                        property.Data = OneNoteBinaryPayload.FromBytes(ReadBytes((int)length));
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
                properties.Add(property);
            }
            return new OneNotePropertySet(properties.AsReadOnly(), _position - start);
        }

        public OneNoteFormatException Error(string code, string message) {
            ulong position = _absoluteOffset + (ulong)_position;
            long offset = position > long.MaxValue ? long.MaxValue : (long)position;
            return new OneNoteFormatException(code, message, offset);
        }

        private void SetScalar(OneNotePropertyValue property, int byteCount) {
            byte[] bytes = ReadBytes(byteCount);
            ulong value = 0;
            for (int index = 0; index < bytes.Length; index++) value |= (ulong)bytes[index] << (index * 8);
            property.ScalarValue = value;
            property.Data = OneNoteBinaryPayload.FromBytes(bytes);
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
            byte value = (byte)(compact & 0xFFU);
            uint index = compact >> 8;
            if (!_globalIds.TryGetValue(index, out Guid identifier)) {
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

        private byte[] ReadBytes(int length) {
            Ensure(length);
            var value = new byte[length];
            if (length > 0) Buffer.BlockCopy(_data, _position, value, 0, length);
            _position += length;
            return value;
        }

        private void Ensure(int length) {
            if (length < 0 || _position > _data.Length - length) {
                throw Error("ONENOTE_TRUNCATED_PROPERTY_SET", "The object data ended inside a property set.");
            }
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
            var result = new List<OneNoteExtendedGuid>(count);
            for (int item = 0; item < count; item++) result.Add(source[index++]);
            return result.AsReadOnly();
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
