namespace OfficeIMO.OneNote;

internal sealed class FssHttpWriteObject {
    internal FssHttpWriteObject(int type, byte[]? data = null, IEnumerable<FssHttpWriteObject>? children = null) {
        Type = type;
        Data = data ?? Array.Empty<byte>();
        Children = children?.ToArray() ?? Array.Empty<FssHttpWriteObject>();
    }

    internal int Type { get; }
    internal byte[] Data { get; }
    internal IReadOnlyList<FssHttpWriteObject> Children { get; }
    internal bool Compound => Children.Count > 0;
}

internal static class FssHttpStreamObjectWriter {
    internal static byte[] Write(FssHttpWriteObject value) {
        long length = GetEncodedLength(value);
        if (length > int.MaxValue) throw new IOException("The FSSHTTP payload exceeds the supported in-memory size.");
        using (var stream = new MemoryStream((int)length)) {
            WriteObject(stream, value);
            return stream.ToArray();
        }
    }

    internal static long GetEncodedLength(FssHttpWriteObject value) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (value.Type < 0 || value.Type > 0x3FFF) throw new ArgumentOutOfRangeException(nameof(value));
        checked {
            long length = 4L + GetCompactUInt64Length((ulong)value.Data.LongLength) + value.Data.LongLength;
            foreach (FssHttpWriteObject child in value.Children) length += GetEncodedLength(child);
            if (value.Compound) length += value.Type <= 0x3F ? 1L : 2L;
            return length;
        }
    }

    internal static void WriteObject(Stream stream, FssHttpWriteObject value) {
        WriteStartHeader(stream, value.Type, value.Compound, (ulong)value.Data.LongLength);
        stream.Write(value.Data, 0, value.Data.Length);
        foreach (FssHttpWriteObject child in value.Children) WriteObject(stream, child);
        if (value.Compound) WriteEndHeader(stream, value.Type);
    }

    internal static void WriteCompactUInt64(Stream stream, ulong value) {
        if (value <= 0x7FUL) { stream.WriteByte((byte)((value << 1) | 1)); return; }
        for (int length = 2; length <= 7; length++) {
            int valueBits = length * 8 - length;
            if (value < (1UL << valueBits)) {
                ulong raw = (value << length) | (1UL << (length - 1));
                for (int index = 0; index < length; index++) stream.WriteByte((byte)(raw >> (index * 8)));
                return;
            }
        }
        stream.WriteByte(0x80);
        WriteUInt64(stream, value);
    }

    internal static void WriteExtendedGuid(Stream stream, OneNoteExtendedGuid value) {
        if (value.Identifier == Guid.Empty && value.Value == 0) { stream.WriteByte(0); return; }
        uint number = value.Value;
        if (number <= 0x1F) {
            stream.WriteByte((byte)((number << 3) | 0x04));
        } else if (number <= 0x3FF) {
            WriteUInt16(stream, (ushort)((number << 6) | 0x20));
        } else if (number <= 0x1FFFF) {
            uint raw = (number << 7) | 0x40;
            stream.WriteByte((byte)raw);
            stream.WriteByte((byte)(raw >> 8));
            stream.WriteByte((byte)(raw >> 16));
        } else {
            stream.WriteByte(0x80);
            WriteUInt32(stream, number);
        }
        byte[] guid = value.Identifier.ToByteArray();
        stream.Write(guid, 0, guid.Length);
    }

    internal static void WriteGuid(Stream stream, Guid value) {
        byte[] data = value.ToByteArray();
        stream.Write(data, 0, data.Length);
    }

    internal static void WriteUInt16(Stream stream, ushort value) {
        stream.WriteByte((byte)value);
        stream.WriteByte((byte)(value >> 8));
    }

    internal static void WriteUInt32(Stream stream, uint value) {
        stream.WriteByte((byte)value);
        stream.WriteByte((byte)(value >> 8));
        stream.WriteByte((byte)(value >> 16));
        stream.WriteByte((byte)(value >> 24));
    }

    internal static void WriteUInt64(Stream stream, ulong value) {
        WriteUInt32(stream, (uint)value);
        WriteUInt32(stream, (uint)(value >> 32));
    }

    private static int GetCompactUInt64Length(ulong value) {
        if (value < 0x7FFFUL) return 0;
        if (value <= 0x7FUL) return 1;
        for (int length = 2; length <= 7; length++) {
            int valueBits = length * 8 - length;
            if (value < (1UL << valueBits)) return length;
        }
        return 9;
    }

    private static void WriteStartHeader(Stream stream, int type, bool compound, ulong length) {
        if (type < 0 || type > 0x3FFF) throw new ArgumentOutOfRangeException(nameof(type));
        uint encodedLength = length < 0x7FFF ? (uint)length : 0x7FFFU;
        uint raw = 0x02U | (compound ? 0x04U : 0U) | ((uint)type << 3) | (encodedLength << 17);
        WriteUInt32(stream, raw);
        if (encodedLength == 0x7FFFU) WriteCompactUInt64(stream, length);
    }

    private static void WriteEndHeader(Stream stream, int type) {
        if (type <= 0x3F) stream.WriteByte((byte)((type << 2) | 0x01));
        else WriteUInt16(stream, (ushort)((type << 2) | 0x03));
    }
}
