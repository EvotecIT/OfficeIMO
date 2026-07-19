namespace OfficeIMO.Email.Store;

internal static class PstBinary {
    internal static ushort UInt16(byte[] bytes, int offset) {
        Ensure(bytes, offset, 2);
        return (ushort)(bytes[offset] | (bytes[offset + 1] << 8));
    }

    internal static short Int16(byte[] bytes, int offset) => unchecked((short)UInt16(bytes, offset));

    internal static uint UInt32(byte[] bytes, int offset) {
        Ensure(bytes, offset, 4);
        return (uint)(bytes[offset] | (bytes[offset + 1] << 8) |
            (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24));
    }

    internal static int Int32(byte[] bytes, int offset) => unchecked((int)UInt32(bytes, offset));

    internal static ulong UInt64(byte[] bytes, int offset) {
        uint low = UInt32(bytes, offset);
        uint high = UInt32(bytes, offset + 4);
        return low | ((ulong)high << 32);
    }

    internal static long Int64(byte[] bytes, int offset) => unchecked((long)UInt64(bytes, offset));

    internal static void WriteUInt16(byte[] bytes, int offset, int value) {
        EnsureWrite(bytes, offset, 2);
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
    }

    internal static void WriteUInt32(byte[] bytes, int offset, uint value) {
        EnsureWrite(bytes, offset, 4);
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }

    internal static void WriteUInt64(byte[] bytes, int offset, ulong value) {
        WriteUInt32(bytes, offset, (uint)value);
        WriteUInt32(bytes, offset + 4, (uint)(value >> 32));
    }

    internal static byte[] ReadAt(Stream stream, long offset, int count) {
        if (offset < 0 || count < 0 || offset > stream.Length - count) {
            throw new InvalidDataException("A PST structure points outside the source stream.");
        }
        stream.Position = offset;
        var bytes = new byte[count];
        int total = 0;
        while (total < count) {
            int read = stream.Read(bytes, total, count - total);
            if (read == 0) throw new EndOfStreamException();
            total += read;
        }
        return bytes;
    }

    internal static void Ensure(byte[] bytes, int offset, int count) {
        if (offset < 0 || count < 0 || offset > bytes.Length - count) {
            throw new InvalidDataException("A PST structure is truncated.");
        }
    }

    private static void EnsureWrite(byte[] bytes, int offset, int count) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        if (offset < 0 || count < 0 || offset > bytes.Length - count) {
            throw new ArgumentOutOfRangeException(nameof(offset));
        }
    }

    internal static int Align64(int value) => checked((value + 63) & ~63);

    internal static int Align(int value, int alignment) {
        if (alignment <= 0 || (alignment & (alignment - 1)) != 0) {
            throw new ArgumentOutOfRangeException(nameof(alignment));
        }
        return checked((value + alignment - 1) & ~(alignment - 1));
    }

    internal static ulong NormalizeBid(ulong bid) => bid & ~1UL;
}
