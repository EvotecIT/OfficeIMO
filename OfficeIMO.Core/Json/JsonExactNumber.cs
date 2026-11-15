using System.Globalization;
using System.Numerics;
using System.Text;

namespace OfficeIMO.Json;

/// <summary>
/// An exact decimal number, <see cref="Mantissa"/> × 10^<see cref="Exponent"/>, normalized so equal values have equal
/// fields. It lets comparisons, equality and <c>multipleOf</c> work on numbers beyond <see cref="decimal"/> and
/// <see cref="double"/> precision without loss.
/// </summary>
internal readonly struct JsonExactNumber {
    /// <summary>Exponent alignments above this many digits fall back to binary floating point.</summary>
    private const int MaxAlignment = 10_000;
    private const int MaxExponentDigits = 9;

    private JsonExactNumber(BigInteger mantissa, int exponent, int digits) {
        Mantissa = mantissa;
        Exponent = exponent;
        Digits = digits;
    }

    public BigInteger Mantissa { get; }
    public int Exponent { get; }
    /// <summary>The number of decimal digits in the mantissa's magnitude; 1 for zero.</summary>
    public int Digits { get; }

    public bool IsInteger => Mantissa.IsZero || Exponent >= 0;

    /// <summary>Parses <c>[+-]digits[.digits][(e|E)[+-]digits]</c>; returns false for other text or exponents beyond nine digits.</summary>
    public static bool TryParse(string text, out JsonExactNumber number) {
        number = default;
        int i = 0;
        bool negative = false;
        if (i < text.Length && (text[i] == '-' || text[i] == '+')) negative = text[i++] == '-';
        var digits = new StringBuilder(text.Length);
        int fraction = 0;
        bool any = false;
        while (i < text.Length && IsDigit(text[i])) { digits.Append(text[i++]); any = true; }
        if (i < text.Length && text[i] == '.') {
            i++;
            while (i < text.Length && IsDigit(text[i])) { digits.Append(text[i++]); fraction++; any = true; }
        }
        if (!any) return false;
        long exponent = 0;
        if (i < text.Length && (text[i] == 'e' || text[i] == 'E')) {
            i++;
            bool negativeExponent = false;
            if (i < text.Length && (text[i] == '-' || text[i] == '+')) negativeExponent = text[i++] == '-';
            int start = i;
            while (i < text.Length && IsDigit(text[i])) {
                if (i - start >= MaxExponentDigits) return false;
                exponent = exponent * 10 + (text[i++] - '0');
            }
            if (i == start) return false;
            if (negativeExponent) exponent = -exponent;
        }
        if (i != text.Length) return false;
        exponent -= fraction;
        // Normalize: no leading zeros, no trailing zeros (they move into the exponent).
        int first = 0;
        while (first < digits.Length - 1 && digits[first] == '0') first++;
        int last = digits.Length - 1;
        while (last > first && digits[last] == '0') {
            last--;
            exponent++;
        }
        string significant = digits.ToString(first, last - first + 1);
        if (significant == "0") {
            number = new JsonExactNumber(BigInteger.Zero, 0, 1);
            return true;
        }
        if (exponent > int.MaxValue / 2 || exponent < int.MinValue / 2) return false;
        BigInteger mantissa = BigInteger.Parse(significant, NumberStyles.None, CultureInfo.InvariantCulture);
        number = new JsonExactNumber(negative ? -mantissa : mantissa, (int)exponent, significant.Length);
        return true;
    }

    public static JsonExactNumber FromInt64(long value) {
        TryParse(value.ToString(CultureInfo.InvariantCulture), out JsonExactNumber number);
        return number;
    }

    public static int Compare(JsonExactNumber left, JsonExactNumber right) {
        int leftSign = left.Mantissa.Sign, rightSign = right.Mantissa.Sign;
        if (leftSign != rightSign) return leftSign.CompareTo(rightSign);
        if (leftSign == 0) return 0;
        // Same sign: compare orders of magnitude first, so huge exponents never need aligning.
        long leftMagnitude = (long)left.Digits + left.Exponent, rightMagnitude = (long)right.Digits + right.Exponent;
        if (leftMagnitude != rightMagnitude) return leftMagnitude.CompareTo(rightMagnitude) * leftSign;
        int exponent = Math.Min(left.Exponent, right.Exponent);
        BigInteger a = left.Mantissa * BigInteger.Pow(10, left.Exponent - exponent);
        BigInteger b = right.Mantissa * BigInteger.Pow(10, right.Exponent - exponent);
        return a.CompareTo(b);
    }

    /// <summary>Returns whether the value is an integer multiple of the divisor, or null when aligning them would be too large.</summary>
    public static bool? IsMultipleOf(JsonExactNumber value, JsonExactNumber divisor) {
        if (divisor.Mantissa.IsZero) return null;
        if (value.Mantissa.IsZero) return true;
        int exponent = Math.Min(value.Exponent, divisor.Exponent);
        long valueShift = (long)value.Exponent - exponent, divisorShift = (long)divisor.Exponent - exponent;
        if (valueShift > MaxAlignment || divisorShift > MaxAlignment) return null;
        BigInteger a = value.Mantissa * BigInteger.Pow(10, (int)valueShift);
        BigInteger b = divisor.Mantissa * BigInteger.Pow(10, (int)divisorShift);
        return (a % b).IsZero;
    }

    public bool TryToDecimal(out decimal value) {
        value = 0;
        if (Digits > 29 || Exponent > 28 || Exponent < -28 - Digits) return false;
        if (!decimal.TryParse(ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out value)) return false;
        return TryParse(value.ToString(CultureInfo.InvariantCulture), out JsonExactNumber roundTrip) && Compare(roundTrip, this) == 0;
    }

    public double ToDouble() {
        double.TryParse(ToString(), NumberStyles.Float, CultureInfo.InvariantCulture, out double value);
        return value;
    }

    /// <summary>Formats the number as JSON: plain digits for moderate exponents, otherwise scientific notation.</summary>
    public override string ToString() {
        if (Mantissa.IsZero) return "0";
        string digits = BigInteger.Abs(Mantissa).ToString(CultureInfo.InvariantCulture);
        string sign = Mantissa.Sign < 0 ? "-" : string.Empty;
        if (Exponent >= 0 && Exponent <= 21) return sign + digits + new string('0', Exponent);
        if (Exponent < 0 && -Exponent < digits.Length) return sign + digits.Substring(0, digits.Length + Exponent) + "." + digits.Substring(digits.Length + Exponent);
        if (Exponent < 0 && -Exponent - digits.Length < 6) return sign + "0." + new string('0', -Exponent - digits.Length) + digits;
        string mantissa = digits.Length == 1 ? digits : digits.Substring(0, 1) + "." + digits.Substring(1);
        long scientific = (long)Exponent + digits.Length - 1;
        return sign + mantissa + "e" + (scientific >= 0 ? "+" : string.Empty) + scientific.ToString(CultureInfo.InvariantCulture);
    }

    private static bool IsDigit(char c) => c >= '0' && c <= '9';
}
