namespace OfficeIMO.Email;

internal static class MsgBinary {
    internal static ushort ReadUInt16(byte[] bytes, int offset) {
        Ensure(bytes, offset, 2);
        return (ushort)(bytes[offset] | (bytes[offset + 1] << 8));
    }

    internal static short ReadInt16(byte[] bytes, int offset) => unchecked((short)ReadUInt16(bytes, offset));

    internal static uint ReadUInt32(byte[] bytes, int offset) {
        Ensure(bytes, offset, 4);
        return (uint)(bytes[offset] | (bytes[offset + 1] << 8) | (bytes[offset + 2] << 16) | (bytes[offset + 3] << 24));
    }

    internal static int ReadInt32(byte[] bytes, int offset) => unchecked((int)ReadUInt32(bytes, offset));

    internal static ulong ReadUInt64(byte[] bytes, int offset) {
        return ReadUInt32(bytes, offset) | ((ulong)ReadUInt32(bytes, offset + 4) << 32);
    }

    internal static long ReadInt64(byte[] bytes, int offset) => unchecked((long)ReadUInt64(bytes, offset));

    internal static float ReadSingle(byte[] bytes, int offset) {
        byte[] value = Slice(bytes, offset, 4);
        if (!BitConverter.IsLittleEndian) Array.Reverse(value);
        return BitConverter.ToSingle(value, 0);
    }

    internal static double ReadDouble(byte[] bytes, int offset) {
        byte[] value = Slice(bytes, offset, 8);
        if (!BitConverter.IsLittleEndian) Array.Reverse(value);
        return BitConverter.ToDouble(value, 0);
    }

    internal static byte[] Slice(byte[] bytes, int offset, int count) {
        Ensure(bytes, offset, count);
        byte[] result = new byte[count];
        Buffer.BlockCopy(bytes, offset, result, 0, count);
        return result;
    }

    internal static void WriteUInt16(byte[] bytes, int offset, ushort value) {
        Ensure(bytes, offset, 2);
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
    }

    internal static void WriteUInt32(byte[] bytes, int offset, uint value) {
        Ensure(bytes, offset, 4);
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
        bytes[offset + 2] = (byte)(value >> 16);
        bytes[offset + 3] = (byte)(value >> 24);
    }

    internal static void WriteUInt64(byte[] bytes, int offset, ulong value) {
        WriteUInt32(bytes, offset, unchecked((uint)value));
        WriteUInt32(bytes, offset + 4, unchecked((uint)(value >> 32)));
    }

    internal static string CombinePath(string prefix, string name) {
        return string.IsNullOrEmpty(prefix) ? name : string.Concat(prefix, "/", name);
    }

    private static void Ensure(byte[] bytes, int offset, int count) {
        if (offset < 0 || count < 0 || offset > bytes.Length - count) throw new InvalidDataException("Unexpected end of MSG data.");
    }
}
