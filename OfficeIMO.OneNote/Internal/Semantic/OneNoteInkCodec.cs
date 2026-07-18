using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote;

internal static class OneNoteInkCodec {
    internal const double NativeUnitsPerHalfInch = 1270D;
    internal static readonly Guid XDimension = new Guid("598a6a8f-52c0-4ba0-93af-af357411a561");
    internal static readonly Guid YDimension = new Guid("b53f9f75-04e0-4498-a7ee-c30dbb5a9011");
    internal static readonly Guid PressureDimension = new Guid("2d500773-f4f9-4e18-b3f2-2ce1b1a3610c");

    internal sealed class Dimension {
        internal Dimension(Guid id, int lower, int upper) { Id = id; Lower = lower; Upper = upper; }
        internal Guid Id { get; }
        internal int Lower { get; }
        internal int Upper { get; }
    }

    internal static IReadOnlyList<Dimension> DecodeDimensions(byte[]? data) {
        if (data == null || data.Length == 0) return Array.Empty<Dimension>();
        var dimensions = new List<Dimension>();
        for (int offset = 0; offset + 32 <= data.Length; offset += 32) {
            var guidBytes = new byte[16];
            Buffer.BlockCopy(data, offset, guidBytes, 0, 16);
            dimensions.Add(new Dimension(new Guid(guidBytes), BitConverter.ToInt32(data, offset + 16), BitConverter.ToInt32(data, offset + 20)));
        }
        return dimensions;
    }

    internal static byte[] EncodeDimensions(bool pressure) {
        var dimensions = new List<byte>();
        AddDimension(dimensions, XDimension, int.MinValue, int.MaxValue, 2U, 1000F);
        AddDimension(dimensions, YDimension, int.MinValue, int.MaxValue, 2U, 1000F);
        if (pressure) AddDimension(dimensions, PressureDimension, 0, 32767, 0U, 1F);
        return dimensions.ToArray();
    }

    internal static IReadOnlyList<long> DecodeSignedVector(byte[]? data, int maximumValues) {
        if (data == null || data.Length == 0) return Array.Empty<long>();
        int offset = 0;
        ulong encodedCount = ReadVarUInt(data, ref offset);
        ulong count = encodedCount >> 1;
        if (count > (ulong)maximumValues || count > int.MaxValue) {
            throw new OneNoteFormatException("ONENOTE_INK_PATH_LIMIT", "The ink path exceeds the configured property value limit.");
        }
        var values = new long[(int)count];
        for (int index = 0; index < values.Length; index++) {
            if (offset >= data.Length) throw new OneNoteFormatException("ONENOTE_INK_PATH_TRUNCATED", "The ink path ends before all declared coordinates were decoded.");
            ulong encoded = ReadVarUInt(data, ref offset);
            if ((encoded >> 1) > long.MaxValue) throw new OneNoteFormatException("ONENOTE_INK_PATH_VALUE", "An ink coordinate exceeds the supported signed range.");
            long magnitude = (long)(encoded >> 1);
            values[index] = (encoded & 1UL) == 0UL ? magnitude : -magnitude;
        }
        return values;
    }

    internal static byte[] EncodeSignedVector(IReadOnlyList<long> values) {
        if (values == null) throw new ArgumentNullException(nameof(values));
        var data = new List<byte>();
        WriteVarUInt(data, checked((ulong)values.Count << 1));
        for (int index = 0; index < values.Count; index++) {
            long value = values[index];
            if (value == long.MinValue) throw new ArgumentOutOfRangeException(nameof(values));
            ulong magnitude = (ulong)Math.Abs(value);
            WriteVarUInt(data, checked((magnitude << 1) | (value < 0 ? 1UL : 0UL)));
        }
        return data.ToArray();
    }

    // Each InkPath dimension starts with an absolute packet followed by signed deltas.
    internal static IReadOnlyList<long> DecodePacketValues(IReadOnlyList<long> encoded, int start, int count) {
        if (encoded == null) throw new ArgumentNullException(nameof(encoded));
        if (start < 0 || count < 0 || start > encoded.Count - count) throw new ArgumentOutOfRangeException(nameof(count));
        var values = new long[count];
        if (count == 0) return values;
        try {
            long value = encoded[start];
            values[0] = value;
            for (int packet = 1; packet < count; packet++) {
                value = checked(value + encoded[start + packet]);
                values[packet] = value;
            }
            return values;
        } catch (OverflowException exception) {
            throw new OneNoteFormatException(
                "ONENOTE_INK_PATH_VALUE",
                "An ink packet cannot be reconstructed within the supported signed range.",
                null,
                exception);
        }
    }

    internal static IReadOnlyList<long> EncodePacketValues(IReadOnlyList<long> values) {
        if (values == null) throw new ArgumentNullException(nameof(values));
        if (values.Count < 2) return values.ToArray();
        var encoded = new long[values.Count];
        encoded[0] = values[0];
        try {
            for (int index = 1; index < values.Count; index++) encoded[index] = checked(values[index] - values[index - 1]);
            return encoded;
        } catch (OverflowException exception) {
            throw new OneNoteFormatException(
                "ONENOTE_WRITE_INK_COORDINATE",
                "Ink packet differences exceed the native OneNote signed range.",
                null,
                exception);
        }
    }

    internal static int ToNativeCoordinate(double value) {
        double scaled = value * NativeUnitsPerHalfInch;
        if (double.IsNaN(scaled) || double.IsInfinity(scaled) || scaled < int.MinValue || scaled > int.MaxValue) {
            throw new OneNoteFormatException("ONENOTE_WRITE_INK_COORDINATE", "An ink coordinate is outside the native OneNote range.");
        }
        return (int)Math.Round(scaled);
    }

    internal static OfficeColor DecodeColor(uint? color) {
        if (!color.HasValue) return OfficeColor.Black;
        uint value = color.Value;
        return OfficeColor.FromRgb((byte)value, (byte)(value >> 8), (byte)(value >> 16));
    }

    internal static uint EncodeColor(OfficeColor color) => (uint)(color.R | color.G << 8 | color.B << 16);

    private static void AddDimension(List<byte> output, Guid id, int lower, int upper, uint unit, float resolution) {
        output.AddRange(id.ToByteArray());
        output.AddRange(BitConverter.GetBytes(lower));
        output.AddRange(BitConverter.GetBytes(upper));
        output.AddRange(BitConverter.GetBytes(unit));
        output.AddRange(BitConverter.GetBytes(resolution));
    }

    private static ulong ReadVarUInt(byte[] data, ref int offset) {
        ulong value = 0UL;
        for (int byteIndex = 0; byteIndex < 10; byteIndex++) {
            if (offset >= data.Length) throw new OneNoteFormatException("ONENOTE_INK_VARINT", "The ink path contains a truncated multi-byte integer.");
            byte current = data[offset++];
            if (byteIndex == 9 && (current & 0xFE) != 0) throw new OneNoteFormatException("ONENOTE_INK_VARINT", "The ink path contains a multi-byte integer wider than 64 bits.");
            value |= (ulong)(current & 0x7F) << (byteIndex * 7);
            if ((current & 0x80) == 0) return value;
        }
        throw new OneNoteFormatException("ONENOTE_INK_VARINT", "The ink path contains an invalid multi-byte integer.");
    }

    private static void WriteVarUInt(List<byte> output, ulong value) {
        do {
            byte current = (byte)(value & 0x7F);
            value >>= 7;
            if (value != 0) current |= 0x80;
            output.Add(current);
        } while (value != 0);
    }
}
