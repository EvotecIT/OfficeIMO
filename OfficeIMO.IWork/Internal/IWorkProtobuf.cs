using System.Text;

namespace OfficeIMO.IWork.Internal;

internal enum IWorkWireKind {
    Varint,
    Fixed64,
    Bytes,
    Fixed32
}

internal sealed class IWorkWireValue {
    internal IWorkWireValue(ulong value, IWorkWireKind kind) {
        Unsigned = value;
        Kind = kind;
    }

    internal IWorkWireValue(byte[] bytes) {
        Bytes = bytes;
        Kind = IWorkWireKind.Bytes;
    }

    internal IWorkWireKind Kind { get; }
    internal ulong Unsigned { get; }
    internal byte[]? Bytes { get; }
}

internal sealed class IWorkWireMessage {
    private static readonly UTF8Encoding StrictUtf8 = new(false, true);
    private readonly Dictionary<int, List<IWorkWireValue>> _fields;
    private readonly IWorkReadOptions _options;
    private readonly int _depth;

    internal IWorkWireMessage(Dictionary<int, List<IWorkWireValue>> fields, IWorkReadOptions options, int depth) {
        _fields = fields;
        _options = options;
        _depth = depth;
    }

    internal ulong? GetUnsigned(int field) {
        IWorkWireValue? value = Values(field).FirstOrDefault(candidate => candidate.Kind == IWorkWireKind.Varint);
        return value?.Unsigned;
    }

    internal IReadOnlyList<ulong> GetRepeatedUnsigned(int field, bool packed = false) {
        var result = new List<ulong>();
        foreach (IWorkWireValue value in Values(field)) {
            if (value.Kind == IWorkWireKind.Varint) {
                result.Add(value.Unsigned);
            } else if (packed && value.Kind == IWorkWireKind.Bytes && value.Bytes != null) {
                int offset = 0;
                while (offset < value.Bytes.Length) {
                    if (result.Count >= _options.MaximumProtobufFieldCount) {
                        throw new InvalidDataException($"A packed protobuf field exceeds the configured value limit of {_options.MaximumProtobufFieldCount}.");
                    }
                    result.Add(IWorkProtobuf.ReadVarint(value.Bytes, ref offset));
                }
            }
            if (result.Count > _options.MaximumProtobufFieldCount) {
                throw new InvalidDataException($"A packed protobuf field exceeds the configured value limit of {_options.MaximumProtobufFieldCount}.");
            }
        }
        return result;
    }

    internal uint? GetFixed32(int field) {
        IWorkWireValue? value = Values(field).FirstOrDefault(candidate => candidate.Kind == IWorkWireKind.Fixed32);
        return value == null ? null : (uint)value.Unsigned;
    }

    internal ulong? GetFixed64(int field) {
        IWorkWireValue? value = Values(field).FirstOrDefault(candidate => candidate.Kind == IWorkWireKind.Fixed64);
        return value?.Unsigned;
    }

    internal byte[]? GetBytes(int field) =>
        Values(field).FirstOrDefault(candidate => candidate.Kind == IWorkWireKind.Bytes)?.Bytes;

    internal bool HasBytes(int field) =>
        Values(field).Any(candidate => candidate.Kind == IWorkWireKind.Bytes);

    internal IReadOnlyList<byte[]> GetRepeatedBytes(int field) => Values(field)
        .Where(candidate => candidate.Kind == IWorkWireKind.Bytes && candidate.Bytes != null)
        .Select(candidate => candidate.Bytes!)
        .ToArray();

    internal string? GetString(int field) {
        byte[]? bytes = GetBytes(field);
        if (bytes == null) return null;
        try {
            return StrictUtf8.GetString(bytes);
        } catch (DecoderFallbackException) {
            return null;
        }
    }

    internal IWorkWireMessage? GetMessage(int field) {
        byte[]? bytes = GetBytes(field);
        return bytes == null ? null : IWorkProtobuf.Parse(bytes, _options, _depth + 1);
    }

    internal IReadOnlyList<IWorkWireMessage> GetRepeatedMessages(int field) {
        var result = new List<IWorkWireMessage>();
        foreach (byte[] bytes in GetRepeatedBytes(field)) {
            result.Add(IWorkProtobuf.Parse(bytes, _options, _depth + 1));
        }
        return result;
    }

    private IReadOnlyList<IWorkWireValue> Values(int field) =>
        _fields.TryGetValue(field, out List<IWorkWireValue>? values)
            ? values
            : Array.Empty<IWorkWireValue>();
}

internal static class IWorkProtobuf {
    internal static IWorkWireMessage Parse(byte[] data, IWorkReadOptions options, int depth = 0) {
        if (depth > options.MaximumProtobufDepth) {
            throw new InvalidDataException($"Protobuf nesting exceeds the configured depth of {options.MaximumProtobufDepth}.");
        }

        var fields = new Dictionary<int, List<IWorkWireValue>>();
        int offset = 0;
        int fieldCount = 0;
        while (offset < data.Length) {
            int keyOffset = offset;
            ulong key = ReadVarint(data, ref offset);
            if (key >> 3 == 0 || key >> 3 > int.MaxValue) {
                throw new InvalidDataException($"Invalid protobuf field number at offset {keyOffset}.");
            }

            int field = (int)(key >> 3);
            int wire = (int)(key & 7);
            IWorkWireValue value;
            switch (wire) {
                case 0:
                    value = new IWorkWireValue(ReadVarint(data, ref offset), IWorkWireKind.Varint);
                    break;
                case 1:
                    EnsureAvailable(data, offset, 8, "fixed64");
                    value = new IWorkWireValue(ReadUInt64(data, offset), IWorkWireKind.Fixed64);
                    offset += 8;
                    break;
                case 2:
                    ulong rawLength = ReadVarint(data, ref offset);
                    if (rawLength > int.MaxValue) throw new InvalidDataException("A protobuf field exceeds the supported length.");
                    int length = (int)rawLength;
                    EnsureAvailable(data, offset, length, "length-delimited field");
                    var bytes = new byte[length];
                    Buffer.BlockCopy(data, offset, bytes, 0, length);
                    value = new IWorkWireValue(bytes);
                    offset += length;
                    break;
                case 5:
                    EnsureAvailable(data, offset, 4, "fixed32");
                    value = new IWorkWireValue(ReadUInt32(data, offset), IWorkWireKind.Fixed32);
                    offset += 4;
                    break;
                default:
                    throw new InvalidDataException($"Unsupported protobuf wire type {wire} at offset {keyOffset}.");
            }

            fieldCount++;
            if (fieldCount > options.MaximumProtobufFieldCount) {
                throw new InvalidDataException($"A protobuf message exceeds the configured field limit of {options.MaximumProtobufFieldCount}.");
            }
            if (!fields.TryGetValue(field, out List<IWorkWireValue>? values)) {
                values = new List<IWorkWireValue>();
                fields.Add(field, values);
            }
            values.Add(value);
        }
        return new IWorkWireMessage(fields, options, depth);
    }

    internal static ulong ReadVarint(byte[] data, ref int offset) {
        ulong result = 0;
        int shift = 0;
        while (shift < 64) {
            if (offset >= data.Length) throw new InvalidDataException($"Truncated varint at offset {offset}.");
            byte value = data[offset++];
            if (shift == 63 && (value & 0xfe) != 0) throw new InvalidDataException($"Varint overflows UInt64 at offset {offset - 1}.");
            result |= (ulong)(value & 0x7f) << shift;
            if ((value & 0x80) == 0) return result;
            shift += 7;
        }
        throw new InvalidDataException($"Varint exceeds ten bytes at offset {offset}.");
    }

    internal static uint ReadUInt32(byte[] data, int offset) =>
        (uint)(data[offset]
            | data[offset + 1] << 8
            | data[offset + 2] << 16
            | data[offset + 3] << 24);

    internal static ulong ReadUInt64(byte[] data, int offset) {
        ulong value = 0;
        for (int index = 0; index < 8; index++) value |= (ulong)data[offset + index] << (index * 8);
        return value;
    }

    private static void EnsureAvailable(byte[] data, int offset, int length, string label) {
        if (length < 0 || offset < 0 || offset > data.Length - length) {
            throw new InvalidDataException($"Truncated protobuf {label} at offset {offset}.");
        }
    }
}
