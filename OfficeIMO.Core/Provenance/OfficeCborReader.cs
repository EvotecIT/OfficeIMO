using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Provenance;

/// <summary>A CBOR tag and the item it wraps (for example COSE_Sign1, tag 18).</summary>
internal sealed class OfficeCborTag {
    internal OfficeCborTag(ulong tag, object? value) { Tag = tag; Value = value; }
    internal ulong Tag { get; }
    internal object? Value { get; }
}

/// <summary>
/// Minimal, bounded CBOR (RFC 8949) decoder for reading C2PA claims, assertions, and COSE headers.
/// Produces <see cref="Dictionary{TKey, TValue}"/> (string or long keys), <see cref="List{T}"/>, string, byte[],
/// long/ulong, double, bool, null, and <see cref="OfficeCborTag"/>. Returns false instead of throwing on malformed input.
/// </summary>
internal sealed class OfficeCborReader {
    private const int MaximumDepth = 32;
    private readonly byte[] _data;
    private readonly int _end;
    private int _position;
    private int _budget;

    private OfficeCborReader(byte[] data, int offset, int length, int maximumItems) {
        _data = data;
        _position = offset;
        _end = offset + length;
        _budget = maximumItems;
    }

    internal static bool TryDecode(byte[] data, int offset, int length, out object? value, int maximumItems = 50_000) {
        value = null;
        if (data == null || offset < 0 || length <= 0 || offset > data.Length - length) return false;
        var reader = new OfficeCborReader(data, offset, length, maximumItems);
        try {
            value = reader.ReadItem(0);
            return true;
        } catch (FormatException) {
            value = null;
            return false;
        }
    }

    private object? ReadItem(int depth) {
        if (depth > MaximumDepth) throw new FormatException("CBOR nesting is too deep.");
        if (--_budget < 0) throw new FormatException("CBOR item budget exceeded.");
        byte initial = ReadByte();
        int major = initial >> 5;
        int info = initial & 0x1F;
        switch (major) {
            case 0:
                return Narrow(ReadArgument(info));
            case 1: {
                ulong argument = ReadArgument(info);
                return argument <= long.MaxValue ? -1L - (long)argument : (object)double.NegativeInfinity;
            }
            case 2:
                return info == 31 ? ReadIndefiniteBytes(2) : ReadBytes(Length(ReadArgument(info)));
            case 3:
                return Utf8(info == 31 ? ReadIndefiniteBytes(3) : ReadBytes(Length(ReadArgument(info))));
            case 4: {
                var list = new List<object?>();
                if (info == 31) {
                    while (!TryReadBreak()) list.Add(ReadItem(depth + 1));
                } else {
                    int count = Count(ReadArgument(info));
                    for (int index = 0; index < count; index++) list.Add(ReadItem(depth + 1));
                }
                return list;
            }
            case 5: {
                var map = new Dictionary<object, object?>();
                if (info == 31) {
                    while (!TryReadBreak()) AddEntry(map, depth);
                } else {
                    int count = Count(ReadArgument(info));
                    for (int index = 0; index < count; index++) AddEntry(map, depth);
                }
                return map;
            }
            case 6: {
                ulong tag = ReadArgument(info);
                return new OfficeCborTag(tag, ReadItem(depth + 1));
            }
            default:
                return ReadSimple(info);
        }
    }

    private void AddEntry(Dictionary<object, object?> map, int depth) {
        object? key = ReadItem(depth + 1);
        object? value = ReadItem(depth + 1);
        if (key is string or long) map[key] = value; // other key types are not used by C2PA or COSE headers
    }

    private object? ReadSimple(int info) {
        switch (info) {
            case 20: return false;
            case 21: return true;
            case 22: case 23: return null;
            case 25: return HalfToDouble((ushort)ReadUnsigned(2));
            case 26: {
                byte[] bytes = ReadBytes(4);
                if (BitConverter.IsLittleEndian) Array.Reverse(bytes);
                return (double)BitConverter.ToSingle(bytes, 0);
            }
            case 27: {
                byte[] bytes = ReadBytes(8);
                if (BitConverter.IsLittleEndian) Array.Reverse(bytes);
                return BitConverter.ToDouble(bytes, 0);
            }
            default:
                if (info < 20) return null;
                if (info == 24) { ReadByte(); return null; }
                throw new FormatException("Unsupported CBOR simple value.");
        }
    }

    private static double HalfToDouble(ushort half) {
        int exponent = (half >> 10) & 0x1F;
        int mantissa = half & 0x3FF;
        double value = exponent == 0 ? mantissa * Math.Pow(2, -24)
            : exponent == 31 ? (mantissa == 0 ? double.PositiveInfinity : double.NaN)
            : (mantissa + 1024) * Math.Pow(2, exponent - 25);
        return (half & 0x8000) != 0 ? -value : value;
    }

    private byte[] ReadIndefiniteBytes(int major) {
        var chunks = new List<byte>();
        while (!TryReadBreak()) {
            byte initial = ReadByte();
            if (initial >> 5 != major || (initial & 0x1F) == 31) throw new FormatException("Invalid CBOR string chunk.");
            chunks.AddRange(ReadBytes(Length(ReadArgument(initial & 0x1F))));
        }
        return chunks.ToArray();
    }

    private bool TryReadBreak() {
        if (_position >= _end) throw new FormatException("Unterminated CBOR container.");
        if (_data[_position] != 0xFF) return false;
        _position++;
        return true;
    }

    private ulong ReadArgument(int info) {
        if (info < 24) return (ulong)info;
        return info switch {
            24 => ReadUnsigned(1),
            25 => ReadUnsigned(2),
            26 => ReadUnsigned(4),
            27 => ReadUnsigned(8),
            _ => throw new FormatException("Invalid CBOR argument.")
        };
    }

    private ulong ReadUnsigned(int size) {
        if (_end - _position < size) throw new FormatException("Truncated CBOR item.");
        ulong value = 0;
        for (int index = 0; index < size; index++) value = (value << 8) | _data[_position++];
        return value;
    }

    private byte ReadByte() {
        if (_position >= _end) throw new FormatException("Truncated CBOR item.");
        return _data[_position++];
    }

    private byte[] ReadBytes(int length) {
        if (_end - _position < length) throw new FormatException("Truncated CBOR string.");
        byte[] bytes = new byte[length];
        Buffer.BlockCopy(_data, _position, bytes, 0, length);
        _position += length;
        return bytes;
    }

    private int Length(ulong value) {
        if (value > (ulong)(_end - _position)) throw new FormatException("CBOR string exceeds the remaining data.");
        return (int)value;
    }

    // Every element takes at least one byte, so a count above the remaining bytes is malformed.
    private int Count(ulong value) {
        if (value > (ulong)(_end - _position)) throw new FormatException("CBOR container exceeds the remaining data.");
        return (int)value;
    }

    private static object Narrow(ulong value) => value <= long.MaxValue ? (long)value : (object)value;

    private static string Utf8(byte[] bytes) {
        try {
            return new UTF8Encoding(encoderShouldEmitUTF8Identifier: false, throwOnInvalidBytes: true).GetString(bytes);
        } catch (DecoderFallbackException) {
            throw new FormatException("Invalid UTF-8 in CBOR text.");
        }
    }
}
