namespace OfficeIMO.Project;

/// <summary>Lossless entry-level updates to an MPP property envelope, retaining the producer's trailer.</summary>
internal sealed class ProjectNativePropertySet {
    private readonly List<(uint Id, int Type, byte[] Bytes)> _entries = new List<(uint, int, byte[])>();
    private readonly byte[] _trailer;
    private readonly int _flags;
    internal ProjectNativePropertySet() { _trailer = Array.Empty<byte>(); }
    internal ProjectNativePropertySet(byte[] bytes, CancellationToken token) {
        var validated = ProjectNativeProperties.Read(bytes, token);
        var input = new ProjectNativeValue(bytes, 0, bytes.Length);
        int end = checked(input.Int32() + 4); _trailer = input.Slice(end, input.Length - end).Copy(); _flags = input.Int32(8);
        int offset = 16;
        foreach (var unused in validated) {
            token.ThrowIfCancellationRequested();
            int length = input.Int32(offset); uint id = input.UInt32(offset + 4); int type = input.Int32(offset + 8);
            _entries.Add((id, type, input.Slice(offset + 12, length).Copy())); offset = checked(offset + 12 + length + (length & 1));
        }
    }
    internal void Set(uint id, int type, byte[] bytes) {
        for (int index = 0; index < _entries.Count; index++) if (_entries[index].Id == id) { _entries[index] = (id, type, bytes); return; }
        _entries.Add((id, type, bytes));
    }
    internal void Remove(uint id) => _entries.RemoveAll(entry => entry.Id == id);
    internal byte[] Serialize(long maxBytes, CancellationToken token) {
        long length = 16;
        foreach (var entry in _entries) length = checked(length + 12 + entry.Bytes.Length + (entry.Bytes.Length & 1));
        if (length + _trailer.Length > maxBytes || length > int.MaxValue) throw new InvalidDataException("Native property output exceeds its byte budget.");
        using var output = new MemoryStream(checked((int)(length + _trailer.Length))); using var writer = new BinaryWriter(output);
        writer.Write((int)length - 4); writer.Write((int)length - 4); writer.Write(_flags); writer.Write(_entries.Count);
        foreach (var entry in _entries) {
            token.ThrowIfCancellationRequested(); writer.Write(entry.Bytes.Length); writer.Write(entry.Id); writer.Write(entry.Type); writer.Write(entry.Bytes);
            if ((entry.Bytes.Length & 1) != 0) writer.Write((byte)0);
        }
        writer.Write(_trailer); return output.ToArray();
    }
}
