using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

internal static partial class OfficeWoff2Decoder {
    private static uint ReadBase128(byte[] data, ref int offset) {
        if (offset >= data.Length) throw new InvalidDataException("The WOFF 2 UIntBase128 value is truncated.");
        uint result = 0;
        if (data[offset] == 0x80) throw new InvalidDataException("A WOFF 2 UIntBase128 value has a leading zero.");
        for (int index = 0; index < 5; index++) {
            EnsureAvailable(data, offset, 1, "The WOFF 2 UIntBase128 value is truncated.");
            byte value = data[offset++];
            if ((result & 0xFE000000U) != 0) throw new InvalidDataException("A WOFF 2 UIntBase128 value overflows UInt32.");
            result = (result << 7) | (uint)(value & 0x7F);
            if ((value & 0x80) == 0) return result;
        }
        throw new InvalidDataException("A WOFF 2 UIntBase128 value is longer than five bytes.");
    }

    private static int Read255UInt16(byte[] data, ref int offset) {
        EnsureAvailable(data, offset, 1, "A transformed WOFF 2 UInt16 value is truncated.");
        int code = data[offset++];
        if (code == 253) {
            EnsureAvailable(data, offset, 2, "A transformed WOFF 2 UInt16 value is truncated.");
            int value = ReadUInt16(data, offset);
            offset += 2;
            return value;
        }
        if (code == 254) {
            EnsureAvailable(data, offset, 1, "A transformed WOFF 2 UInt16 value is truncated.");
            return 506 + data[offset++];
        }
        if (code == 255) {
            EnsureAvailable(data, offset, 1, "A transformed WOFF 2 UInt16 value is truncated.");
            return 253 + data[offset++];
        }
        return code;
    }

    private static int Align2(int value) => checked((value + 1) & ~1);
    private static int Align4(int value) => checked((value + 3) & ~3);

    private static uint CalculateChecksum(byte[] data) {
        uint checksum = 0;
        for (int offset = 0; offset < data.Length; offset += 4) {
            uint value = (uint)data[offset] << 24;
            if (offset + 1 < data.Length) value |= (uint)data[offset + 1] << 16;
            if (offset + 2 < data.Length) value |= (uint)data[offset + 2] << 8;
            if (offset + 3 < data.Length) value |= data[offset + 3];
            checksum = unchecked(checksum + value);
        }
        return checksum;
    }

    private static void EnsureAvailable(byte[] data, int offset, int length, string message) {
        if (offset < 0 || length < 0 || offset > data.Length - length) throw new InvalidDataException(message);
    }

    private static ushort ReadUInt16(byte[] data, int offset) {
        EnsureAvailable(data, offset, 2, "Font data is truncated.");
        return unchecked((ushort)((data[offset] << 8) | data[offset + 1]));
    }

    private static short ReadInt16(byte[] data, int offset) => unchecked((short)ReadUInt16(data, offset));

    private static uint ReadUInt32(byte[] data, int offset) {
        EnsureAvailable(data, offset, 4, "Font data is truncated.");
        return unchecked(((uint)data[offset] << 24)
            | ((uint)data[offset + 1] << 16)
            | ((uint)data[offset + 2] << 8)
            | data[offset + 3]);
    }

    private static void WriteUInt16(byte[] data, int offset, ushort value) {
        EnsureAvailable(data, offset, 2, "Font output buffer is too small.");
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }

    private static void WriteInt16(byte[] data, int offset, short value) => WriteUInt16(data, offset, unchecked((ushort)value));

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        EnsureAvailable(data, offset, 4, "Font output buffer is too small.");
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private static uint Tag(string value) {
        if (value == null || value.Length != 4) throw new ArgumentException("OpenType tags must contain four characters.", nameof(value));
        return ((uint)value[0] << 24) | ((uint)value[1] << 16) | ((uint)value[2] << 8) | value[3];
    }

    private sealed class ByteBuilder {
        private readonly List<byte> _bytes;
        private readonly int _maximumCount;

        internal ByteBuilder(int maximumCount = int.MaxValue) {
            if (maximumCount < 0) throw new ArgumentOutOfRangeException(nameof(maximumCount));
            _maximumCount = maximumCount;
            _bytes = new List<byte>(Math.Min(maximumCount, 4096));
        }

        internal int Count => _bytes.Count;
        internal int RemainingCapacity => _maximumCount - _bytes.Count;

        internal void Add(byte value) {
            EnsureCapacity(1);
            _bytes.Add(value);
        }

        internal void Add(byte[] data) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            EnsureCapacity(data.Length);
            _bytes.AddRange(data);
        }

        internal void Add(byte[] data, int offset, int length) {
            EnsureAvailable(data, offset, length, "Font data is truncated.");
            EnsureCapacity(length);
            for (int index = 0; index < length; index++) _bytes.Add(data[offset + index]);
        }

        internal void AddUInt16(ushort value) {
            EnsureCapacity(2);
            _bytes.Add((byte)(value >> 8));
            _bytes.Add((byte)value);
        }

        internal void AddInt16(short value) => AddUInt16(unchecked((ushort)value));

        internal void AddUInt32(uint value) {
            EnsureCapacity(4);
            _bytes.Add((byte)(value >> 24));
            _bytes.Add((byte)(value >> 16));
            _bytes.Add((byte)(value >> 8));
            _bytes.Add((byte)value);
        }

        internal void PadToEven() {
            if ((_bytes.Count & 1) != 0) Add(0);
        }

        internal byte[] ToArray() => _bytes.ToArray();

        private void EnsureCapacity(int additionalCount) {
            if (additionalCount < 0 || _bytes.Count > _maximumCount - additionalCount) {
                throw new InvalidDataException("The reconstructed WOFF 2 table exceeds the configured byte limit.");
            }
        }
    }
}
