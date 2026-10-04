using System.Globalization;
using System.Numerics;

namespace OfficeIMO.IWork.Internal;

internal static partial class IWorkTableReader {
    private static readonly BigInteger Decimal128CoefficientLimit = BigInteger.Pow(10, 34);
    private static double? ReadDecimal128(byte[] buffer, int offset,
        out string? sourceText, out bool approximate) {
        sourceText = null;
        approximate = false;
        if ((buffer[offset + 15] & 0x78) == 0x78) return null;
        int exponent = (((buffer[offset + 15] & 0x7f) << 7) | (buffer[offset + 14] >> 1)) - 0x1820;
        BigInteger coefficient = BigInteger.Zero;
        for (int index = 13; index >= 0; index--)
            coefficient = coefficient * 256 + buffer[offset + index];
        if ((buffer[offset + 14] & 1) != 0) coefficient += BigInteger.One << 112;
        // Decimal128 supports at most 34 digits. Do not assign a value to an
        // unsupported/noncanonical coefficient using floating-point rounding.
        if (coefficient >= Decimal128CoefficientLimit) return null;
        if (coefficient.IsZero) { sourceText = "0"; return 0d; }
        while (coefficient % 10 == 0) { coefficient /= 10; exponent++; }
        string digits = coefficient.ToString(CultureInfo.InvariantCulture);
        string text = ((buffer[offset + 15] & 0x80) != 0 ? "-" : "") + digits
            + "E" + exponent.ToString(CultureInfo.InvariantCulture);
        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture,
                out double value) || !IsFinite(value) || value == 0d) return null;
        sourceText = text;
        approximate = digits.Length > 15;
        return value;
    }
}
