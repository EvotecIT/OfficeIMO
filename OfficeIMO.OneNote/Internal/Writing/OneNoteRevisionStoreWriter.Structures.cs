namespace OfficeIMO.OneNote;

internal interface IOneNoteDesktopReferenceTarget {
    ulong Offset { get; set; }
    uint Length { get; }
}

internal sealed class OneNoteDesktopDataChunk : IOneNoteDesktopReferenceTarget {
    internal OneNoteDesktopDataChunk(byte[] data) {
        Data = data ?? throw new ArgumentNullException(nameof(data));
        if (data.LongLength > uint.MaxValue) throw new OneNoteFormatException("ONENOTE_WRITE_CHUNK_SIZE", "A desktop OneNote data chunk exceeds the 32-bit reference range.");
    }

    internal byte[] Data { get; }
    public ulong Offset { get; set; }
    public uint Length => (uint)Data.Length;
}

internal sealed class OneNoteDesktopFileNodeList : IOneNoteDesktopReferenceTarget {
    private const int HeaderLength = 16;
    private const int TrailerLength = 20;

    internal OneNoteDesktopFileNodeList(uint id) {
        if (id < 0x10) throw new ArgumentOutOfRangeException(nameof(id));
        Id = id;
    }

    internal uint Id { get; }
    internal IList<OneNoteDesktopFileNode> Nodes { get; } = new List<OneNoteDesktopFileNode>();
    public ulong Offset { get; set; }

    public uint Length {
        get {
            long nodeBytes = Nodes.Sum(node => (long)node.Length);
            long length = OneNoteDesktopBinary.Align8(HeaderLength + nodeBytes + TrailerLength);
            if (length > uint.MaxValue) throw new OneNoteFormatException("ONENOTE_WRITE_FILE_NODE_LIST_SIZE", "A desktop OneNote file-node list exceeds the 32-bit reference range.");
            return (uint)length;
        }
    }

    internal byte[] Encode() {
        var data = new byte[checked((int)Length)];
        using (var stream = new MemoryStream(data, true)) {
            FssHttpStreamObjectWriter.WriteUInt64(stream, 0xA4567AB1F5F7F4C4UL);
            FssHttpStreamObjectWriter.WriteUInt32(stream, Id);
            FssHttpStreamObjectWriter.WriteUInt32(stream, 0);
            foreach (OneNoteDesktopFileNode node in Nodes) node.Write(stream);
            stream.Position = data.Length - TrailerLength;
            OneNoteDesktopBinary.WriteNilReference(stream);
            FssHttpStreamObjectWriter.WriteUInt64(stream, 0x8BC215C38233BA4BUL);
        }
        return data;
    }
}

internal sealed class OneNoteDesktopFileNode {
    private OneNoteDesktopFileNode(ushort id, OneNoteFileNodeBaseType baseType, byte[] body, IOneNoteDesktopReferenceTarget? target) {
        Id = id;
        BaseType = baseType;
        Body = body ?? Array.Empty<byte>();
        Target = target;
        if ((baseType == OneNoteFileNodeBaseType.Inline) != (target == null)) throw new ArgumentException("Inline nodes cannot have a target and reference nodes require one.");
        if (Length > 0x1FFF) throw new OneNoteFormatException("ONENOTE_WRITE_FILE_NODE_SIZE", "A desktop OneNote file node exceeds the 13-bit size field.");
    }

    internal ushort Id { get; }
    internal OneNoteFileNodeBaseType BaseType { get; }
    internal byte[] Body { get; }
    internal IOneNoteDesktopReferenceTarget? Target { get; }
    internal int Length => 4 + (Target == null ? 0 : 12) + Body.Length;

    internal static OneNoteDesktopFileNode Inline(OneNoteFileNodeId id, byte[]? body = null) =>
        new OneNoteDesktopFileNode((ushort)id, OneNoteFileNodeBaseType.Inline, body ?? Array.Empty<byte>(), null);

    internal static OneNoteDesktopFileNode Data(OneNoteFileNodeId id, IOneNoteDesktopReferenceTarget target, byte[]? body = null) =>
        new OneNoteDesktopFileNode((ushort)id, OneNoteFileNodeBaseType.DataReference, body ?? Array.Empty<byte>(), target);

    internal static OneNoteDesktopFileNode List(OneNoteFileNodeId id, IOneNoteDesktopReferenceTarget target, byte[]? body = null) =>
        new OneNoteDesktopFileNode((ushort)id, OneNoteFileNodeBaseType.FileNodeListReference, body ?? Array.Empty<byte>(), target);

    internal void Write(Stream stream) {
        uint header = Id |
            ((uint)Length << 10) |
            ((uint)BaseType << 27) |
            0x80000000U;
        FssHttpStreamObjectWriter.WriteUInt32(stream, header);
        if (Target != null) OneNoteDesktopBinary.WriteReference(stream, Target.Offset, Target.Length);
        stream.Write(Body, 0, Body.Length);
    }
}

internal static class OneNoteDesktopBinary {
    internal static long Align8(long value) => checked((value + 7L) & ~7L);

    internal static byte[] Data(Action<Stream> writer) {
        using (var stream = new MemoryStream()) {
            writer(stream);
            return stream.ToArray();
        }
    }

    internal static void WriteGuid(Stream stream, Guid value) => FssHttpStreamObjectWriter.WriteGuid(stream, value);

    internal static void WriteExtendedGuid(Stream stream, OneNoteExtendedGuid value) {
        WriteGuid(stream, value.Identifier);
        FssHttpStreamObjectWriter.WriteUInt32(stream, value.Value);
    }

    internal static void WriteReference(Stream stream, ulong offset, uint length) {
        FssHttpStreamObjectWriter.WriteUInt64(stream, offset);
        FssHttpStreamObjectWriter.WriteUInt32(stream, length);
    }

    internal static void WriteNilReference(Stream stream) => WriteReference(stream, ulong.MaxValue, 0);

    internal static uint CompactId(OneNoteExtendedGuid id, IReadOnlyDictionary<Guid, uint> globalIds) {
        if (!globalIds.TryGetValue(id.Identifier, out uint index) || index >= 0xFFFFFFU || id.Value > byte.MaxValue) {
            throw new OneNoteFormatException("ONENOTE_WRITE_COMPACT_ID", "A desktop OneNote identity is not present in the global-identification table.");
        }
        return (index << 8) | id.Value;
    }
}
