using System.Numerics;

namespace OfficeIMO.Invoicing;

/// <summary>Keeps products and quotients exact until monetary rounding, avoiding decimal intermediate underflow.</summary>
internal static class InvoiceArithmetic {
    private static readonly BigInteger MaximumCoefficient = (BigInteger.One << 96) - 1;

    internal static decimal Add(decimal first, decimal second) => Sum(new[] { first, second });

    internal static decimal Sum(IEnumerable<decimal> values) {
        BigInteger total = BigInteger.Zero;
        int totalScale = 0;
        foreach (decimal value in values) {
            Parts(value, out BigInteger coefficient, out int scale);
            if (scale > totalScale) { total *= BigInteger.Pow(10, scale - totalScale); totalScale = scale; }
            total += coefficient * BigInteger.Pow(10, totalScale - scale);
        }
        return ExactDecimal(total, totalScale);
    }

    internal static decimal LineAmount(decimal quantity, decimal price, decimal basis, decimal adjustments, out decimal rounded) {
        ProductRatio(quantity, price, basis, out BigInteger numerator, out BigInteger denominator);
        Parts(adjustments, out BigInteger adjustment, out int scale);
        BigInteger factor = BigInteger.Pow(10, scale);
        numerator = numerator * factor + adjustment * denominator;
        denominator *= factor;
        rounded = Money(numerator, denominator);
        return DecimalApproximation(numerator, denominator);
    }

    internal static decimal RoundedProduct(decimal first, decimal second, decimal divisor) {
        ProductRatio(first, second, divisor, out BigInteger numerator, out BigInteger denominator);
        return Money(numerator, denominator);
    }

    private static void ProductRatio(decimal first, decimal second, decimal divisor, out BigInteger numerator, out BigInteger denominator) {
        if (divisor == 0m) throw new DivideByZeroException();
        Parts(first, out BigInteger a, out int aScale);
        Parts(second, out BigInteger b, out int bScale);
        Parts(divisor, out BigInteger c, out int cScale);
        numerator = a * b * BigInteger.Pow(10, cScale);
        denominator = c * BigInteger.Pow(10, aScale + bScale);
        if (denominator.Sign < 0) { numerator = -numerator; denominator = -denominator; }
    }

    private static decimal Money(BigInteger numerator, BigInteger denominator) {
        BigInteger cents = Quantize(numerator, denominator, 2, tiesToPositiveInfinity: true);
        return ExactDecimal(cents, 2);
    }

    private static decimal ExactDecimal(BigInteger coefficient, int scale) {
        while (BigInteger.Abs(coefficient) > MaximumCoefficient && scale > 0 && coefficient % 10 == 0) { coefficient /= 10; scale--; }
        if (BigInteger.Abs(coefficient) > MaximumCoefficient) throw new OverflowException("Invoice amount cannot be represented exactly within decimal capacity.");
        return Decimal(coefficient, scale);
    }

    private static decimal DecimalApproximation(BigInteger numerator, BigInteger denominator) {
        for (int scale = 28; scale >= 0; scale--) {
            BigInteger value = Quantize(numerator, denominator, scale, tiesToPositiveInfinity: false);
            if (BigInteger.Abs(value) <= MaximumCoefficient) return Decimal(value, scale);
        }
        throw new OverflowException("Invoice formula exceeds decimal capacity.");
    }

    private static BigInteger Quantize(BigInteger numerator, BigInteger denominator, int scale, bool tiesToPositiveInfinity) {
        BigInteger value = BigInteger.DivRem(BigInteger.Abs(numerator) * BigInteger.Pow(10, scale), denominator, out BigInteger remainder);
        int half = (remainder * 2).CompareTo(denominator);
        if (half > 0 || half == 0 && (tiesToPositiveInfinity ? numerator.Sign >= 0 : !value.IsEven)) value++;
        return numerator.Sign < 0 ? -value : value;
    }

    private static decimal Decimal(BigInteger coefficient, int scale) {
        BigInteger value = BigInteger.Abs(coefficient);
        return new decimal(unchecked((int)(uint)(value & uint.MaxValue)), unchecked((int)(uint)((value >> 32) & uint.MaxValue)),
            unchecked((int)(uint)((value >> 64) & uint.MaxValue)), coefficient.Sign < 0, (byte)scale);
    }

    private static void Parts(decimal value, out BigInteger coefficient, out int scale) {
        int[] bits = decimal.GetBits(value);
        coefficient = (BigInteger)(uint)bits[0] | (BigInteger)(uint)bits[1] << 32 | (BigInteger)(uint)bits[2] << 64;
        if (bits[3] < 0) coefficient = -coefficient;
        scale = (bits[3] >> 16) & 0xff;
    }
}
