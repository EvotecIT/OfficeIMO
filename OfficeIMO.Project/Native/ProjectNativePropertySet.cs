namespace OfficeIMO.Project;

/// <summary>Lossless entry-level updates to an MPP property envelope, retaining the producer's trailer.</summary>
internal sealed class ProjectNativePropertySet {
    private readonly List<(uint Id, int Type, byte[] Bytes)> _entries = new List<(uint, int, byte[])>();
    private readonly byte[] _trailer;
    private readonly int _flags;
    private readonly ushort _entryFlags;
    private readonly bool _legacyExternal;
    internal ProjectNativePropertySet(bool legacyExternal = false) { _trailer = Array.Empty<byte>(); _legacyExternal = legacyExternal; }
    internal ProjectNativePropertySet(byte[] bytes, CancellationToken token, bool legacyExternal = false) {
        _legacyExternal = legacyExternal;
        var validated = ProjectNativeProperties.Read(bytes, token, legacyExternal);
        var input = new ProjectNativeValue(bytes, 0, bytes.Length);
        int end = checked(input.Int32() + 4), externalOffset = end; _flags = input.Int32(8);
        _entryFlags = input.UInt16(14);
        int offset = 16;
        foreach (var unused in validated) {
            token.ThrowIfCancellationRequested();
            int length = input.Int32(offset); uint id = input.UInt32(offset + 4); int type = input.Int32(offset + 8);
            if (legacyExternal && type == 0x10000) {
                _entries.Add((id, type, input.Slice(externalOffset, checked(length + 4)).Copy()));
                externalOffset = checked(externalOffset + length + 4); offset = checked(offset + 16);
            } else { _entries.Add((id, type, input.Slice(offset + 12, length).Copy())); offset = checked(offset + 12 + length + (length & 1)); }
        }
        _trailer = input.Slice(externalOffset, input.Length - externalOffset).Copy();
    }
    internal void Set(uint id, int type, byte[] bytes) {
        for (int index = 0; index < _entries.Count; index++) if (_entries[index].Id == id) { _entries[index] = (id, type, bytes); return; }
        _entries.Add((id, type, bytes));
    }
    internal void Remove(uint id) => _entries.RemoveAll(entry => entry.Id == id);
    internal byte[] Serialize(long maxBytes, CancellationToken token) {
        long length = 16, externalLength = 0;
        foreach (var entry in _entries) {
            if (_legacyExternal && entry.Type == 0x10000) {
                if (entry.Bytes.Length < 4) throw new InvalidDataException("Truncated external property value.");
                length = checked(length + 16); externalLength = checked(externalLength + entry.Bytes.Length);
            } else length = checked(length + 12 + entry.Bytes.Length + (entry.Bytes.Length & 1));
        }
        if (length + externalLength + _trailer.Length > maxBytes || length > int.MaxValue || _entries.Count > ushort.MaxValue) throw new InvalidDataException("Native property output exceeds its byte or entry budget.");
        using var output = new MemoryStream(checked((int)(length + externalLength + _trailer.Length))); using var writer = new BinaryWriter(output);
        writer.Write((int)length - 4); writer.Write((int)length - 4); writer.Write(_flags); writer.Write((ushort)_entries.Count); writer.Write(_entryFlags);
        foreach (var entry in _entries) {
            token.ThrowIfCancellationRequested();
            bool external = _legacyExternal && entry.Type == 0x10000;
            writer.Write(entry.Bytes.Length - (external ? 4 : 0)); writer.Write(entry.Id); writer.Write(entry.Type);
            if (external) writer.Write(0);
            else { writer.Write(entry.Bytes); if ((entry.Bytes.Length & 1) != 0) writer.Write((byte)0); }
        }
        foreach (var entry in _entries) if (_legacyExternal && entry.Type == 0x10000) { token.ThrowIfCancellationRequested(); writer.Write(entry.Bytes); }
        writer.Write(_trailer); return output.ToArray();
    }
}
