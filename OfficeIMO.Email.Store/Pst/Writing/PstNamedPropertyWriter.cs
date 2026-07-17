using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed class PstNamedPropertyWriter {
    private static readonly Guid PsMapi = new Guid("00020328-0000-0000-C000-000000000046");
    private static readonly Guid PsPublicStrings = new Guid("00020329-0000-0000-C000-000000000046");
    private static readonly Guid OfficeImoUnknownPropertySet =
        new Guid("E962B602-9F1E-4F76-BC29-4795CD1752F7");
    private const int BucketCount = 251;
    private readonly Dictionary<string, NamedEntry> _byIdentity =
        new Dictionary<string, NamedEntry>(StringComparer.Ordinal);
    private readonly List<NamedEntry> _entries = new List<NamedEntry>();

    internal PstNamedPropertyWriter() {
        // An empty binary entry stream is treated as a missing Name-to-ID map by
        // Outlook-compatible readers. Seed one harmless, standard public-string
        // mapping and one PS_COMMON mapping so every required map stream is
        // materialized even before the caller writes a named property.
        GetOrAdd(new MapiNamedProperty(PsPublicStrings, "Keywords"));
        GetOrAdd(new MapiNamedProperty(
            new Guid("00062008-0000-0000-C000-000000000046"), 0x8501));
    }

    internal IReadOnlyList<MapiProperty> Map(IEnumerable<MapiProperty> properties,
        Action<EmailStoreDiagnostic>? reportDiagnostic, string location) {
        var result = new List<MapiProperty>();
        foreach (MapiProperty property in properties) {
            ushort propertyId = property.PropertyId;
            MapiNamedProperty? name = property.Name;
            if (name == null && propertyId >= 0x8000) {
                name = new MapiNamedProperty(OfficeImoUnknownPropertySet, propertyId);
                reportDiagnostic?.Invoke(new EmailStoreDiagnostic(
                    "EMAIL_STORE_PST_WRITE_NAMED_PROPERTY_PLACEHOLDER",
                    string.Concat("Named property 0x", propertyId.ToString("X4",
                        CultureInfo.InvariantCulture),
                        " had no source Name-to-ID mapping and was preserved under an OfficeIMO placeholder property set."),
                    EmailStoreDiagnosticSeverity.Warning, location));
            }
            if (name != null) propertyId = GetOrAdd(name).PropertyId;
            result.Add(new MapiProperty(propertyId, property.PropertyType,
                property.Value, property.Flags, name) {
                RawData = property.RawData == null ? null : (byte[])property.RawData.Clone()
            });
        }
        return result;
    }

    internal IReadOnlyList<MapiProperty> BuildProperties() {
        var guidIndexes = new Dictionary<Guid, int>();
        var guidStream = new List<byte>();
        var stringStream = new List<byte>();
        var entryStream = new byte[_entries.Count * 8];
        var buckets = Enumerable.Range(0, BucketCount).Select(_ => new List<byte>()).ToArray();

        for (int index = 0; index < _entries.Count; index++) {
            NamedEntry entry = _entries[index];
            int guidIndex = GetGuidIndex(entry.Name.PropertySet, entry.Name.IsStringNamed,
                guidIndexes, guidStream);
            uint identifier;
            uint bucketIdentifier;
            if (entry.Name.IsStringNamed) {
                byte[] nameBytes = Encoding.Unicode.GetBytes(entry.Name.Name!);
                identifier = checked((uint)stringStream.Count);
                AddUInt32(stringStream, checked((uint)nameBytes.Length));
                stringStream.AddRange(nameBytes);
                while ((stringStream.Count & 3) != 0) stringStream.Add(0);
                bucketIdentifier = PstCrc32.Compute(nameBytes);
            } else {
                identifier = entry.Name.LocalId.GetValueOrDefault();
                bucketIdentifier = identifier;
            }
            ushort guidAndKind = checked((ushort)((guidIndex << 1) |
                (entry.Name.IsStringNamed ? 1 : 0)));
            ushort propertyIndex = checked((ushort)(entry.PropertyId - 0x8000));
            WriteNameId(entryStream, index * 8, identifier, guidAndKind, propertyIndex);

            uint firstDword = bucketIdentifier;
            uint secondDword = (uint)guidAndKind | ((uint)propertyIndex << 16);
            int bucket = checked((int)((firstDword ^ (secondDword & 0xFFFFU)) % BucketCount));
            byte[] bucketEntry = new byte[8];
            WriteNameId(bucketEntry, 0, bucketIdentifier, guidAndKind, propertyIndex);
            buckets[bucket].AddRange(bucketEntry);
        }

        var properties = new List<MapiProperty> {
            new MapiProperty(0x0001, MapiPropertyType.Integer32, BucketCount),
            new MapiProperty(0x0002, MapiPropertyType.Binary, guidStream.ToArray()),
            new MapiProperty(0x0003, MapiPropertyType.Binary, entryStream),
            new MapiProperty(0x0004, MapiPropertyType.Binary, stringStream.ToArray())
        };
        for (int index = 0; index < buckets.Length; index++) {
            if (buckets[index].Count == 0) continue;
            properties.Add(new MapiProperty(checked((ushort)(0x1000 + index)),
                MapiPropertyType.Binary, buckets[index].ToArray()));
        }
        return properties;
    }

    private NamedEntry GetOrAdd(MapiNamedProperty name) {
        string identity = name.IsStringNamed
            ? string.Concat(name.PropertySet.ToString("D"), ":S:", name.Name)
            : string.Concat(name.PropertySet.ToString("D"), ":N:",
                name.LocalId.GetValueOrDefault().ToString("X8", CultureInfo.InvariantCulture));
        if (_byIdentity.TryGetValue(identity, out NamedEntry? existing)) return existing;
        if (_entries.Count >= 0x8000) throw new NotSupportedException("The PST named-property map is full.");
        var created = new NamedEntry(checked((ushort)(0x8000 + _entries.Count)), name);
        _entries.Add(created);
        _byIdentity.Add(identity, created);
        return created;
    }

    private static int GetGuidIndex(Guid guid, bool stringNamed,
        IDictionary<Guid, int> indexes, ICollection<byte> stream) {
        if (stringNamed && guid == Guid.Empty) return 0;
        if (guid == PsMapi) return 1;
        if (guid == PsPublicStrings) return 2;
        if (indexes.TryGetValue(guid, out int existing)) return existing;
        int index = checked(3 + indexes.Count);
        indexes.Add(guid, index);
        foreach (byte value in guid.ToByteArray()) stream.Add(value);
        return index;
    }

    private static void WriteNameId(byte[] bytes, int offset, uint identifier,
        ushort guidAndKind, ushort propertyIndex) {
        PstBinary.WriteUInt32(bytes, offset, identifier);
        PstBinary.WriteUInt16(bytes, offset + 4, guidAndKind);
        PstBinary.WriteUInt16(bytes, offset + 6, propertyIndex);
    }

    private static void AddUInt32(ICollection<byte> bytes, uint value) {
        bytes.Add((byte)value);
        bytes.Add((byte)(value >> 8));
        bytes.Add((byte)(value >> 16));
        bytes.Add((byte)(value >> 24));
    }

    private sealed class NamedEntry {
        internal NamedEntry(ushort propertyId, MapiNamedProperty name) {
            PropertyId = propertyId;
            Name = name;
        }
        internal ushort PropertyId { get; }
        internal MapiNamedProperty Name { get; }
    }
}
