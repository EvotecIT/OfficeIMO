using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace OfficeIMO.Drawing.Binary;

// MS-ODRAW 2.2.58 and 2.3.6.27 define SG records and their backward references.
// Integer rounding follows the corresponding VML formula contract; native
// producer rounding and final appearance remain an importer qualification gap.
internal static class OfficeArtGeometryGuides {
    internal static bool TryEvaluate(IReadOnlyList<OfficeArtProperty> properties,
        double width, double height, Action<int> accountItems, CancellationToken token,
        out int[] values, out OfficeArtCustomPathFailure failure) {
        values = Array.Empty<int>(); failure = OfficeArtCustomPathFailure.InvalidGuide;
        token.ThrowIfCancellationRequested();
        OfficeArtProperty? property = properties.LastOrDefault(item => item.PropertyId == 0x0156);
        if (property?.IsComplex != true || !property.HasCompleteComplexData
            || property.CopyComplexData() is not byte[] bytes || bytes.Length < 6) return false;
        int count = U16(bytes, 0);
        if (count is < 1 or > 128 || count > U16(bytes, 2) || U16(bytes, 4) != 8
            || 6 + count * 8 > bytes.Length) return false;
        accountItems(count);
        token.ThrowIfCancellationRequested();
        var calculated = new int[count];
        var context = new OfficeArtGeometryGuideContext(properties, width, height);
        for (int index = 0; index < count; index++) {
            token.ThrowIfCancellationRequested();
            int offset = 6 + index * 8;
            int operation = U16(bytes, offset), kind = operation & 0x1FFF;
            if (kind > 0x10) { failure = OfficeArtCustomPathFailure.GuideFormula; return false; }
            int arity = kind is 3 or 13 ? 1 : kind is 2 or 4 or 5 or 8 or 9 or 10 or 16 ? 2 : 3;
            int first = 0, second = 0, third = 0;
            OfficeArtCustomPathFailure parameterFailure = OfficeArtCustomPathFailure.InvalidGuide;
            if (!Parameter(0, out first) || arity >= 2 && !Parameter(1, out second)
                || arity == 3 && !Parameter(2, out third)) { failure = parameterFailure; return false; }
            if (!TryCalculate(kind, first, second, third, out calculated[index])) {
                failure = OfficeArtCustomPathFailure.InvalidGuide; return false;
            }

            bool Parameter(int parameter, out int value) {
                int raw = U16(bytes, offset + 2 + parameter * 2);
                value = raw;
                if ((operation & (0x2000 << parameter)) == 0) return true;
                if (raw is >= 0x0400 and <= 0x047F) {
                    int source = raw - 0x0400;
                    if (source >= index) { parameterFailure = OfficeArtCustomPathFailure.InvalidGuide; return false; }
                    value = calculated[source]; return true;
                }
                return context.TryParameter(raw, out value, out parameterFailure);
            }
            // A calculated operand outside the owned, device-independent
            // parameter set must not be interpreted as its numeric identifier.
            // Invalid backward references keep the malformed-data diagnosis.
        }
        token.ThrowIfCancellationRequested();
        values = calculated; failure = OfficeArtCustomPathFailure.None; return true;
    }

    private static bool TryCalculate(int kind, int first, int second, int third, out int value) {
        value = 0;
        if (kind == 1) return Product(first, second, third, out value);
        double result = kind switch {
            0 => (long)first + second - third,
            2 => ((long)first + second) / 2,
            3 => Math.Abs((long)first),
            4 => Math.Min(first, second),
            5 => Math.Max(first, second),
            6 => first > 0 ? second : third,
            7 => Math.Sqrt((double)first * first + (double)second * second + (double)third * third),
            8 => first == 0 && second == 0 ? double.NaN : Math.Atan2(second, first) * (180D * 65536 / Math.PI),
            9 => first * Sine(second),
            10 => first * Cosine(second),
            11 => first * (double)second / Math.Sqrt((double)second * second + (double)third * third),
            12 => first * (double)third / Math.Sqrt((double)second * second + (double)third * third),
            13 => Math.Sqrt(first),
            14 => (long)first + (long)second * 65536 + (long)third * 65536,
            15 => third * Math.Sqrt(1D - (double)first * first / ((double)second * second)),
            16 => first * Tangent(second),
            _ => double.NaN
        };
        // Exact operations already have integral results. Inexact operations
        // round toward negative infinity; reject undefined or overflowing math.
        result = Math.Floor(result);
        if (double.IsNaN(result) || result < int.MinValue || result > int.MaxValue) return false;
        value = (int)result; return true;
    }

    private static bool Product(int first, int second, int third, out int value) {
        value = 0;
        if (third == 0) return false;
        long numerator = (long)first * second, denominator = third;
        if (denominator < 0) { numerator = -numerator; denominator = -denominator; }
        long quotient = numerator / denominator, remainder = numerator % denominator;
        // Nearest integer, ties toward positive infinity, without double's
        // loss of low product bits at the signed 32-bit operand boundary.
        if (remainder > 0 && remainder * 2 >= denominator) quotient++;
        else if (remainder < 0 && -remainder * 2 > denominator) quotient--;
        if (quotient < int.MinValue || quotient > int.MaxValue) return false;
        value = (int)quotient; return true;
    }

    private static double Sine(int angle) {
        int turn = angle % (360 * 65536);
        if (turn % (90 * 65536) == 0) {
            int quadrant = (turn / (90 * 65536) + 4) % 4;
            return quadrant == 1 ? 1 : quadrant == 3 ? -1 : 0;
        }
        return Math.Sin(turn * (Math.PI / (180D * 65536)));
    }
    private static double Cosine(int angle) {
        int turn = angle % (360 * 65536);
        if (turn % (90 * 65536) == 0) {
            int quadrant = (turn / (90 * 65536) + 4) % 4;
            return quadrant == 0 ? 1 : quadrant == 2 ? -1 : 0;
        }
        return Math.Cos(turn * (Math.PI / (180D * 65536)));
    }
    private static double Tangent(int angle) {
        int turn = angle % (180 * 65536);
        if (Cosine(angle) == 0) return double.NaN;
        if (turn == 0) return 0;
        if (turn is 45 * 65536 or -135 * 65536) return 1;
        if (turn is -45 * 65536 or 135 * 65536) return -1;
        return Math.Tan(turn * (Math.PI / (180D * 65536)));
    }
    private static ushort U16(byte[] bytes, int offset) => unchecked((ushort)(bytes[offset] | bytes[offset + 1] << 8));
}
