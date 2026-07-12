using OfficeIMO.Shared;

namespace OfficeIMO.Email;

internal sealed class MsgNamedPropertyWriter {
    private readonly Dictionary<string, Entry> _byName = new Dictionary<string, Entry>(StringComparer.Ordinal);
    private readonly List<Entry> _entries = new List<Entry>();

    internal ushort GetPropertyId(MapiNamedProperty name) {
        string key = CreateKey(name);
        Entry entry;
        if (_byName.TryGetValue(key, out entry!)) return entry.PropertyId;
        if (_entries.Count >= 0x7fff) throw new InvalidOperationException("MSG named-property count exceeds the MAPI mapped-ID range.");
        entry = new Entry(unchecked((ushort)(0x8000 + _entries.Count)), name);
        _entries.Add(entry);
        _byName[key] = entry;
        return entry.PropertyId;
    }

    internal void WriteStreams(IList<OfficeCompoundStream> streams) {
        var guidIndexes = new Dictionary<Guid, int>();
        var guidOrder = new List<Guid>();
        var lookupStreams = new Dictionary<string, MemoryStream>(StringComparer.Ordinal);
        using (MemoryStream entryStream = new MemoryStream())
        using (MemoryStream stringStream = new MemoryStream()) {
            foreach (Entry entry in _entries) {
                MapiNamedProperty name = entry.Name;
                int guidIndex;
                if (name.PropertySet == MsgNamedPropertyMap.PsMapi) {
                    guidIndex = 1;
                } else if (name.PropertySet == MsgNamedPropertyMap.PsPublicStrings) {
                    guidIndex = 2;
                } else if (!guidIndexes.TryGetValue(name.PropertySet, out guidIndex)) {
                    guidIndex = guidOrder.Count + 3;
                    guidIndexes[name.PropertySet] = guidIndex;
                    guidOrder.Add(name.PropertySet);
                }

                uint identifier;
                bool stringNamed = name.Name != null;
                if (stringNamed) {
                    identifier = checked((uint)stringStream.Length);
                    byte[] text = Encoding.Unicode.GetBytes(name.Name!);
                    WriteUInt32(stringStream, checked((uint)text.Length));
                    stringStream.Write(text, 0, text.Length);
                    while (stringStream.Length % 4 != 0) stringStream.WriteByte(0);
                } else {
                    identifier = name.LocalId.GetValueOrDefault();
                }

                WriteUInt32(entryStream, identifier);
                ushort propertyIndex = unchecked((ushort)(entry.PropertyId - 0x8000));
                ushort guidAndKind = checked((ushort)(guidIndex << 1));
                if (stringNamed) guidAndKind |= 0x0001;
                WriteUInt16(entryStream, guidAndKind);
                WriteUInt16(entryStream, propertyIndex);

                uint lookupIdentifier = stringNamed
                    ? ComputeNameCrc(name.PropertySet == MsgProjection.PsInternetHeaders
                        ? name.Name!.ToLowerInvariant()
                        : name.Name!)
                    : identifier;
                uint kind = stringNamed ? 1U : 0U;
                uint streamId = checked(0x1000U + ((lookupIdentifier ^ ((uint)guidIndex << 1 | kind)) % 0x1fU));
                string lookupName = string.Concat("__nameid_version1.0/__substg1.0_",
                    ((streamId << 16) | 0x0102U).ToString("X8", CultureInfo.InvariantCulture));
                if (!lookupStreams.TryGetValue(lookupName, out MemoryStream? lookupStream)) {
                    lookupStream = new MemoryStream();
                    lookupStreams[lookupName] = lookupStream;
                }
                WriteUInt32(lookupStream, lookupIdentifier);
                WriteUInt16(lookupStream, guidAndKind);
                WriteUInt16(lookupStream, propertyIndex);
            }

            using (MemoryStream guidStream = new MemoryStream()) {
                foreach (Guid guid in guidOrder) {
                    byte[] bytes = guid.ToByteArray();
                    guidStream.Write(bytes, 0, bytes.Length);
                }
                streams.Add(new OfficeCompoundStream("__nameid_version1.0/__substg1.0_00020102", guidStream.ToArray()));
            }
            streams.Add(new OfficeCompoundStream("__nameid_version1.0/__substg1.0_00030102", entryStream.ToArray()));
            streams.Add(new OfficeCompoundStream("__nameid_version1.0/__substg1.0_00040102", stringStream.ToArray()));
            foreach (KeyValuePair<string, MemoryStream> lookup in lookupStreams.OrderBy(item => item.Key, StringComparer.Ordinal)) {
                streams.Add(new OfficeCompoundStream(lookup.Key, lookup.Value.ToArray()));
                lookup.Value.Dispose();
            }
        }
    }

    private static uint ComputeNameCrc(string name) {
        byte[] bytes = Encoding.Unicode.GetBytes(name);
        uint crc = 0;
        for (int index = 0; index < bytes.Length; index++) {
            crc ^= bytes[index];
            for (int bit = 0; bit < 8; bit++) {
                crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320U : crc >> 1;
            }
        }
        return crc;
    }

    private static string CreateKey(MapiNamedProperty name) {
        return name.Name == null
            ? string.Concat(name.PropertySet.ToString("D"), ":L:", name.LocalId.GetValueOrDefault().ToString("X8", CultureInfo.InvariantCulture))
            : string.Concat(name.PropertySet.ToString("D"), ":N:", name.Name);
    }

    private static void WriteUInt32(Stream stream, uint value) {
        byte[] bytes = new byte[4];
        MsgBinary.WriteUInt32(bytes, 0, value);
        stream.Write(bytes, 0, bytes.Length);
    }

    private static void WriteUInt16(Stream stream, ushort value) {
        byte[] bytes = new byte[2];
        MsgBinary.WriteUInt16(bytes, 0, value);
        stream.Write(bytes, 0, bytes.Length);
    }

    private sealed class Entry {
        internal Entry(ushort propertyId, MapiNamedProperty name) {
            PropertyId = propertyId;
            Name = name;
        }

        internal ushort PropertyId { get; }

        internal MapiNamedProperty Name { get; }
    }
}
