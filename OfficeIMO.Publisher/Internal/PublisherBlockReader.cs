namespace OfficeIMO.Publisher.Internal;

// Publisher's Contents and Quill style records use byte identifiers, typed values,
// and self-inclusive lengths for variable blocks. These are distinct from OfficeArt records.
internal sealed class PublisherBlockReader {
    private readonly PublisherBinaryData _data;
    private readonly PublisherParseContext _context;
    internal PublisherBlockReader(PublisherBinaryData data, PublisherParseContext context) { _data = data; _context = context; }
    internal IReadOnlyList<PublisherBlock> Read(int start, int end) {
        _data.Range(start, end - start);
        var result = new List<PublisherBlock>();
        for (int position = start; position < end;) {
            _context.Record(); _data.Range(position, 2, end);
            byte id = _data.U8(position), type = _data.U8(position + 1);
            int payload = position + 2;
            int length = FixedLength(type);
            bool variable = length < 0;
            if (variable) {
                _data.Range(payload, 4, end);
                length = _data.Offset(_data.U32(payload));
                if (length < 4) throw new InvalidDataException("Publisher variable block length is smaller than its length field.");
            }
            _data.Range(payload, length, end);
            uint value = !variable && length == 4 ? _data.U32(payload) : !variable && length == 2 ? _data.U16(payload) : 0U;
            result.Add(new PublisherBlock(id, type, payload, length, value));
            position = payload + length;
        }
        return result;
    }
    internal IReadOnlyList<PublisherBlock> Children(PublisherBlock block) {
        if (FixedLength(block.Type) >= 0 || block.Type == 0xC0) throw new InvalidDataException("Publisher scalar block cannot contain records.");
        return Read(block.Offset + 4, block.Offset + block.Length);
    }
    internal IReadOnlyList<PublisherBlock> Chunk(int offset, int? boundary = null) {
        int length = _data.Offset(_data.U32(offset));
        if (length < 4) throw new InvalidDataException("Publisher chunk length is smaller than its header.");
        _data.Range(offset, length, boundary);
        return Read(offset + 4, offset + length);
    }
    private static int FixedLength(byte type) => type switch {
        0x00 or 0x02 or 0x05 or 0x08 or 0x0A or 0x78 => 0,
        0x07 or 0x10 or 0x12 or 0x18 or 0x1A => 2,
        0x20 or 0x22 or 0x58 or 0x68 or 0x70 or 0xB8 => 4,
        0x28 => 8, 0x38 => 16, 0x48 => 24,
        0x80 or 0x82 or 0x88 or 0x8A or 0x90 or 0x98 or 0xA0 or 0xC0 => -1,
        _ => throw new NotSupportedException($"Unsupported Publisher block encoding 0x{type:X2}.")
    };
}

internal readonly struct PublisherBlock {
    internal PublisherBlock(byte id, byte type, int offset, int length, uint value) { Id = id; Type = type; Offset = offset; Length = length; Value = value; }
    internal byte Id { get; }
    internal byte Type { get; }
    internal int Offset { get; }
    internal int Length { get; }
    internal uint Value { get; }
}
