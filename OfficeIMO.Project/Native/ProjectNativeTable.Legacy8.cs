using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project;

internal sealed partial class ProjectNativeTable {
    internal bool IsLegacy8 => _profile == ProjectNativeProfile.Mpp8;

    /// <summary>Reads the Project 98 table descriptor and its fixed records without interpreting process pointers.</summary>
    internal ProjectNativeTable(OfficeCompoundFile file, string table, ProjectNativeValue descriptor, int maxRecords, CancellationToken token) {
        _profile = ProjectNativeProfile.Mpp8; _prefix = _profile.DataRoot + "/" + table + "/";
        _streams = file.Streams.ToDictionary(s => s.Key, s => s.Value, StringComparer.Ordinal);
        var layout = ProjectNativeProperties.Read(descriptor.Copy(), token);
        if (!layout.TryGetValue(1, out var map) || !layout.TryGetValue(5, out var widthValue) || !layout.TryGetValue(7, out var presenceValue))
            throw new InvalidDataException("Incomplete Project 98 table descriptor.");
        int width = widthValue.Int32(), presence = presenceValue.Int32();
        if (map.Length % 24 != 0 || map.Length / 24 > 8192 || width < 4 || width > 1024 * 1024 ||
            presence < 0 || presence > width - (map.Length / 24 + 7) / 8)
            throw new InvalidDataException("Invalid Project 98 field map or fixed record width.");
        var deferred = new HashSet<uint>(); var indirect = new HashSet<uint>();
        for (int offset = 0; offset < map.Length; offset += 24) {
            token.ThrowIfCancellationRequested();
            uint id = map.UInt32(offset + 8); int type = map.UInt16(offset + 16), attributes = map.Int32(offset + 12);
            var field = new ProjectNativeField { Id = id, Position = offset / 24, Offset = map.UInt16(offset + 4),
                Size = map.UInt16(offset + 18), Type = type, Source = type == 100 ? 18 : (attributes & 128) != 0 ? 4 : 10,
                LegacyIndirect = (attributes & 160) == 128 };
            if (type == 100) {
                int bit = map.UInt16(offset + 20);
                if (bit > 31) throw new InvalidDataException("Invalid Project 98 boolean bit.");
                field.Mask = 1u << bit;
            }
            if (Fields.ContainsKey(id)) throw new InvalidDataException("Duplicate Project 98 field.");
            Fields.Add(id, field);
            if (field.Source == 4) deferred.Add(id);
            if ((attributes & 160) == 128) indirect.Add(id);
        }
        byte[] fixedBytes = Stream("FixFix   0");
        if (fixedBytes.Length % width != 0 || fixedBytes.Length / width > maxRecords)
            throw new InvalidDataException("Invalid Project 98 fixed table length.");
        byte[] variableBytes = OptionalStream("FixDeferFix   0") ?? Array.Empty<byte>();
        var blocks = new Dictionary<int, byte[]>(); long expandedBytes = 0;
        byte[] Block(int pointer) {
            if (blocks.TryGetValue(pointer, out var existing)) return existing;
            var result = ReadLegacy8Blocks(variableBytes, pointer, token);
            expandedBytes = checked(expandedBytes + result.Length);
            if (expandedBytes > variableBytes.Length) throw new InvalidDataException("Project 98 deferred expansion exceeds its stream budget.");
            blocks.Add(pointer, result); return result;
        }
        for (int offset = 0; offset < fixedBytes.Length; offset += width) {
            token.ThrowIfCancellationRequested();
            var data = new ProjectNativeValue(fixedBytes, offset, width);
            int uid = data.Int32();
            if (uid < 0) continue; // Reserved/deleted rows contain no semantic data.
            if (!layout.TryGetValue(6, out var recordFlags)) throw new InvalidDataException("Missing Project 98 record-state descriptor.");
            if ((data.UInt32(recordFlags.Int32()) & 6) != 0) continue;
            byte[] flags = new byte[8 + (Fields.Count + 7) / 8];
            data.Slice(presence, flags.Length - 8).Copy().CopyTo(flags, 8);
            var indexed = new Dictionary<int, ProjectNativeValue>();
            if (layout.TryGetValue(8, out var indexedOffset) && indexedOffset.Int32() >= 0) {
                int pointer = data.Int32(indexedOffset.Int32());
                if (pointer < 0) {
                    var blob = Block(~pointer);
                    var entries = new ProjectNativeValue(blob, 0, blob.Length);
                    for (int at = 0; at < entries.Length;) {
                        token.ThrowIfCancellationRequested();
                        int length = entries.Int32(at), position = entries.Int32(checked(at + 4));
                        if (position < 0 || position >= Fields.Count || indexed.ContainsKey(position)) throw new InvalidDataException("Invalid Project 98 indexed field.");
                        indexed.Add(position, entries.Slice(checked(at + 8), length));
                        at = checked(at + 8 + length + (length & 1));
                    }
                } else if (pointer != 0) throw new NotSupportedException("Project 98 external indexed storage is not yet qualified.");
            }
            foreach (uint id in deferred) {
                var field = Fields[id];
                if ((flags[8 + field.Position / 8] & (1 << (field.Position % 8))) == 0) continue;
                if (indexed.TryGetValue(field.Position, out var value)) {
                    if (indirect.Contains(id)) {
                        int pointer = value.Int32();
                        if (pointer == -1 || pointer == 0) continue;
                        if (pointer > 0) throw new NotSupportedException("Project 98 external variable streams are not yet qualified.");
                        var bytes = Block(~pointer); value = new ProjectNativeValue(bytes, 0, bytes.Length);
                    }
                } else {
                    int pointer = data.Int32(field.Offset);
                    if (pointer == -1) continue;
                    if (pointer >= 0) throw new NotSupportedException("Project 98 external variable storage is not yet qualified: " + table + "/" + uid + "/" + id.ToString("X8") + ".");
                    var bytes = Block(~pointer);
                    value = new ProjectNativeValue(bytes, 0, bytes.Length);
                }
                if (!Variable.TryGetValue(uid, out var values)) Variable.Add(uid, values = new Dictionary<uint, ProjectNativeValue>());
                values.Add(id, value);
            }
            Records.Add(new ProjectNativeRecord(this, offset / width, data, new ProjectNativeValue(flags, 0, flags.Length), null, null) { Uid = uid });
        }
    }

    private static byte[] ReadLegacy8Blocks(byte[] bytes, int start, CancellationToken token) {
        var stream = new ProjectNativeValue(bytes, 0, bytes.Length);
        if (start < 4 || (start - 4) % 36 != 0) throw new InvalidDataException("Unaligned Project 98 deferred block.");
        int length = stream.Int32(checked(start + 4));
        if (length < 0 || length > bytes.Length) throw new InvalidDataException("Project 98 deferred value exceeds its stream.");
        var result = new byte[length]; var visited = new HashSet<int>();
        int offset = start, written = 0; bool first = true;
        do {
            token.ThrowIfCancellationRequested();
            if (offset < 4 || (offset - 4) % 36 != 0 || !visited.Add(offset)) throw new InvalidDataException("Invalid or cyclic Project 98 deferred chain.");
            int next = stream.Int32(offset), header = first ? 8 : 4;
            int take = Math.Min(36 - header, length - written);
            stream.Slice(checked(offset + header), take).Copy().CopyTo(result, written);
            written += take; first = false;
            if (written == length) {
                if (next != -1) throw new InvalidDataException("Project 98 deferred chain exceeds its declared length.");
                break;
            }
            if (next == -1) throw new InvalidDataException("Truncated Project 98 deferred chain.");
            offset = next;
        } while (written < length);
        return result;
    }
}
