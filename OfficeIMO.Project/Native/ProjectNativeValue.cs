using System.Text;

namespace OfficeIMO.Project;

/// <summary>A bounded view over one native record value. All integers are little endian.</summary>
internal readonly struct ProjectNativeValue {
    private readonly byte[] _bytes;
    private readonly int _offset;
    internal int Length { get; }
    internal ProjectNativeValue(byte[] bytes, int offset, int length) {
        if (offset < 0 || length < 0 || offset > bytes.Length - length) throw new InvalidDataException("Native value lies outside its stream.");
        _bytes = bytes; _offset = offset; Length = length;
    }
    internal ProjectNativeValue Slice(int offset, int length) {
        Require(offset, length); return new ProjectNativeValue(_bytes, checked(_offset + offset), length);
    }
    private void Require(int offset, int length) {
        if (offset < 0 || length < 0 || offset > Length - length) throw new InvalidDataException("Truncated native record value.");
    }
    internal byte Byte(int offset) { Require(offset, 1); return _bytes[_offset + offset]; }
    internal ushort UInt16(int offset = 0) { Require(offset, 2); return (ushort)(Byte(offset) | Byte(offset + 1) << 8); }
    internal short Int16(int offset = 0) => unchecked((short)UInt16(offset));
    internal uint UInt32(int offset = 0) { Require(offset, 4); return (uint)(UInt16(offset) | UInt16(offset + 2) << 16); }
    internal int Int32(int offset = 0) => unchecked((int)UInt32(offset));
    internal double Double(int offset = 0) {
        Require(offset, 8);
        long bits = unchecked((long)((ulong)UInt32(offset) | (ulong)UInt32(offset + 4) << 32));
        return BitConverter.Int64BitsToDouble(bits);
    }
    internal byte[] Copy() { var bytes = new byte[Length]; Buffer.BlockCopy(_bytes, _offset, bytes, 0, Length); return bytes; }
    internal string Unicode() {
        if (Length % 2 != 0) throw new InvalidDataException("Native UTF-16 value has an odd byte length.");
        return new UnicodeEncoding(false, false, true).GetString(_bytes, _offset, Length).TrimEnd('\0');
    }
    internal DateTime? Date() {
        uint value = UInt32();
        if (value == uint.MaxValue || value == 0) return null;
        int tenths = (int)(value & 65535);
        if (tenths > 14400) throw new InvalidDataException("Native clock value exceeds one day.");
        return new DateTime(1983, 12, 31).AddDays(value >> 16).AddTicks(tenths * (TimeSpan.TicksPerMinute / 10));
    }
}

/// <summary>Length-prefixed MPP property entries; separate from OLE property sets.</summary>
internal static class ProjectNativeProperties {
    internal static Dictionary<uint, ProjectNativeValue> Read(byte[] bytes, CancellationToken token) {
        var input = new ProjectNativeValue(bytes, 0, bytes.Length);
        if (bytes.Length < 16) throw new InvalidDataException("Truncated MPP property header.");
        int declaredLength = checked(input.Int32() + 4);
        if (declaredLength < 16 || declaredLength > bytes.Length || input.Int32() != input.Int32(4))
            throw new InvalidDataException("MPP property envelope length is invalid.");
        input = input.Slice(0, declaredLength);
        int count = input.Int32(12);
        if (count < 0 || count > (bytes.Length - 16) / 12) throw new InvalidDataException("Invalid MPP property count.");
        var result = new Dictionary<uint, ProjectNativeValue>();
        int offset = 16;
        for (int i = 0; i < count; i++) {
            token.ThrowIfCancellationRequested();
            int length = input.Int32(offset);
            uint key = input.UInt32(offset + 4);
            var value = input.Slice(checked(offset + 12), length);
            if (result.ContainsKey(key)) throw new InvalidDataException("Duplicate MPP property identifier.");
            result.Add(key, value);
            offset = checked(offset + 12 + length + (length & 1));
        }
        if (offset != declaredLength) throw new InvalidDataException("MPP property count does not match its envelope.");
        return result;
    }
}
