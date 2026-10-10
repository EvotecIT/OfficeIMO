namespace OfficeIMO.Chm;

// All offsets are checked as unsigned values before converting to array indexes.
internal static class ChmBinary {
    internal static ChmReadException Error(string code, string message) => new ChmReadException("CHM_" + code, message);
    internal static void Range(byte[] data, long offset, long length) {
        if (offset < 0 || length < 0 || offset > data.LongLength || length > data.LongLength - offset)
            throw Error("TRUNCATED", "A CHM structure or entry extends beyond its containing stream.");
    }
    internal static ushort U16(byte[] data, int offset) {
        Range(data, offset, 2); return (ushort)(data[offset] | data[offset + 1] << 8);
    }
    internal static uint U32(byte[] data, int offset) {
        Range(data, offset, 4);
        return (uint)(data[offset] | data[offset + 1] << 8 | data[offset + 2] << 16 | data[offset + 3] << 24);
    }
    internal static ulong U64(byte[] data, int offset) => U32(data, offset) | (ulong)U32(data, checked(offset + 4)) << 32;
    internal static int Index(ulong value) {
        if (value > int.MaxValue) throw Error("BOUNDS", "A CHM offset or length exceeds the supported in-memory range.");
        return (int)value;
    }
    internal static bool Signature(byte[] data, int offset, string value) {
        Range(data, offset, value.Length);
        for (int i = 0; i < value.Length; i++) if (data[offset + i] != value[i]) return false;
        return true;
    }
    internal static ulong EncInt(byte[] data, ref int position, int end) {
        ulong value = 0;
        for (int i = 0; i < 10; i++) {
            if (position >= end) throw Error("DIRECTORY", "A directory integer is truncated.");
            byte next = data[position++];
            if (value > (ulong.MaxValue >> 7)) throw Error("DIRECTORY", "A directory integer overflows 64 bits.");
            value = (value << 7) | (uint)(next & 127);
            if (next < 128) return value;
        }
        throw Error("DIRECTORY", "A directory integer is too long.");
    }
    internal static string CString(byte[] data, int offset, Encoding encoding, int maximumBytes) {
        Range(data, offset, 1);
        int end = offset;
        int limit = (int)Math.Min(data.LongLength, (long)offset + maximumBytes + 1);
        while (end < limit && data[end] != 0) end++;
        if (end == limit) throw Error("STRING", "A CHM string is unterminated or exceeds its limit.");
        return encoding.GetString(data, offset, end - offset);
    }
}
