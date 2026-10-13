using System.Globalization;

namespace OfficeIMO.IWork.Internal;

/// <summary>Encodes a finite double's round-tripping decimal representation in native iWork Decimal128 storage.</summary>
internal static class IWorkDecimal128Writer {
    internal static byte[] Encode(double value) {
        if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(nameof(value));
        bool negative = BitConverter.DoubleToInt64Bits(value) < 0;
        string text = Math.Abs(value).ToString("R", CultureInfo.InvariantCulture);
        int exponent = 0;
        int marker = text.IndexOf('E');
        if (marker < 0) marker = text.IndexOf('e');
        if (marker >= 0) {
            exponent = int.Parse(text.Substring(marker + 1), NumberStyles.Integer, CultureInfo.InvariantCulture);
            text = text.Substring(0, marker);
        }
        int point = text.IndexOf('.');
        if (point >= 0) { exponent -= text.Length - point - 1; text = text.Remove(point, 1); }
        // A round-tripping double needs at most seventeen significant decimal digits.
        // That coefficient fits in UInt64, while Decimal128 reserves113 bits for it.
        ulong coefficient = ulong.Parse(text, NumberStyles.None, CultureInfo.InvariantCulture);
        var bytes = new byte[16];
        for (int index = 0; index < 8; index++) bytes[index] = (byte)(coefficient >> (index * 8));
        int biasedExponent = checked(exponent + 0x1820);
        bytes[14] = (byte)(biasedExponent << 1);
        bytes[15] = (byte)((biasedExponent >> 7) | (negative ? 0x80 : 0));
        return bytes;
    }
}
