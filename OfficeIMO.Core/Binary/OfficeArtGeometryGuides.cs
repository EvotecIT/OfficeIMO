using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Threading;

namespace OfficeIMO.Drawing.Binary;

// MS-ODRAW 2.2.58 and 2.3.6.27 define SG records and their backward references.
// Integer rounding follows the corresponding VML formula contract; native
// producer rounding and final appearance remain an importer qualification gap.
internal sealed class OfficeArtGeometryGuides {
    private readonly OfficeArtComplexData _bytes;
    private readonly int _storedSize;
    private readonly int[] _values;
    private readonly bool[] _evaluated;
    private readonly OfficeArtGeometryGuideContext _context;
    private readonly CancellationToken _token;

    private OfficeArtGeometryGuides(OfficeArtComplexData bytes, int count, int storedSize,
        OfficeArtGeometryGuideContext context, CancellationToken token) {
        _bytes = bytes; _storedSize = storedSize; _values = new int[count];
        _evaluated = new bool[count]; _context = context; _token = token;
    }

    internal int Count => _values.Length;

    /// <summary>Validates and charges the entire stored array without evaluating unused formulas.</summary>
    internal static bool TryCreate(IReadOnlyList<OfficeArtProperty> properties,
        double width, double height, Action<int> accountItems, CancellationToken token,
        out OfficeArtGeometryGuides? guides, out OfficeArtCustomPathFailure failure) {
        guides = null; failure = OfficeArtCustomPathFailure.InvalidGuide;
        token.ThrowIfCancellationRequested();
        OfficeArtProperty? property = properties.LastOrDefault(item => item.PropertyId == 0x0156);
        if (property?.IsComplex != true || !property.HasCompleteComplexData) return false;
        OfficeArtComplexData bytes = property.ComplexData;
        if (bytes.Length < 6) return false;
        int count = U16(bytes, 0);
        int storedSize = U16(bytes, 4) switch { 8 => 8, 0xFFF0 => 4, _ => 0 };
        if (count is < 1 or > 128 || count > U16(bytes, 2) || storedSize == 0
            || 6 + count * storedSize != bytes.Length) return false;
        accountItems(count);
        token.ThrowIfCancellationRequested();
        guides = new OfficeArtGeometryGuides(bytes, count, storedSize,
            new OfficeArtGeometryGuideContext(properties, width, height), token);
        failure = OfficeArtCustomPathFailure.None; return true;
    }

    /// <summary>Evaluates one referenced guide and its backward dependencies, memoized for this path.</summary>
    internal bool TryGetValue(int index, out int value, out OfficeArtCustomPathFailure failure) {
        value = 0; failure = OfficeArtCustomPathFailure.InvalidGuide;
        _token.ThrowIfCancellationRequested();
        if ((uint)index >= (uint)Count) return false;
        if (_evaluated[index]) { value = _values[index]; failure = OfficeArtCustomPathFailure.None; return true; }
        int offset = 6 + index * _storedSize;
        int operation = U16(_bytes, offset), kind = operation & 0x1FFF;
        if (kind > 0x10) { failure = OfficeArtCustomPathFailure.GuideFormula; return false; }
        int arity = kind is 3 or 13 ? 1 : kind is 2 or 4 or 5 or 8 or 9 or 10 or 16 ? 2 : 3;
        int first, second = 0, third = 0;
        OfficeArtCustomPathFailure parameterFailure = OfficeArtCustomPathFailure.InvalidGuide;
        if (!Parameter(0, out first) || arity >= 2 && !Parameter(1, out second)
            || arity == 3 && !Parameter(2, out third)) { failure = parameterFailure; return false; }
        if (!TryCalculate(kind, first, second, third, out value)) return false;
        _values[index] = value; _evaluated[index] = true;
        failure = OfficeArtCustomPathFailure.None; return true;

        bool Parameter(int parameter, out int result) {
            // IMsoArray's truncated SG stores the low four bytes and
            // implicitly zeroes the two remaining operands.
            int raw = 2 + parameter * 2 < _storedSize ? U16(_bytes, offset + 2 + parameter * 2) : 0;
            result = raw;
            if ((operation & (0x2000 << parameter)) == 0) return true;
            if (raw is >= 0x0400 and <= 0x047F) {
                int source = raw - 0x0400;
                // Strict backward references bound recursion by the 128-entry array.
                if (source >= index) { parameterFailure = OfficeArtCustomPathFailure.InvalidGuide; return false; }
                return TryGetValue(source, out result, out parameterFailure);
            }
            return _context.TryParameter(raw, out result, out parameterFailure);
        }
    }

    private static bool TryCalculate(int kind, int first, int second, int third, out int value) {
        value = 0;
        if (kind == 1) return Product(first, second, third, out value);
        if (kind is 8 or 9 or 10 or 16) return OfficeArtGuideTrigonometry.TryCalculate(kind, first, second, out value);
        if (kind == 7) {
            ulong magnitude = MagnitudeFloor(MagnitudeSquared(first, second, third));
            if (magnitude > int.MaxValue) return false;
            value = (int)magnitude; return true;
        }
        if (kind is 11 or 12) return NormalizedProduct(first, second, third, kind == 12, out value);
        if (kind == 15) {
            long minor = (long)first * first, major = (long)second * second;
            if (major == 0 || minor > major) return false;
            return RootRatioFloor((BigInteger)third * third * (major - minor), (ulong)major, third < 0, out value);
        }
        double result = kind switch {
            0 => (long)first + second - third,
            2 => ((long)first + second) / 2,
            3 => Math.Abs((long)first),
            4 => Math.Min(first, second),
            5 => Math.Max(first, second),
            6 => first > 0 ? second : third,
            13 => Math.Sqrt(first),
            14 => (long)first + (long)second * 65536 + (long)third * 65536,
            _ => double.NaN
        };
        // Exact operations already have integral results. Inexact operations
        // round toward negative infinity; reject undefined or overflowing math.
        result = Math.Floor(result);
        if (double.IsNaN(result) || result < int.MinValue || result > int.MaxValue) return false;
        value = (int)result; return true;
    }

    private static ulong MagnitudeSquared(int first, int second, int third) =>
        (ulong)((long)first * first) + (ulong)((long)second * second) + (ulong)((long)third * third);

    private static bool NormalizedProduct(int first, int second, int third, bool sine, out int value) {
        value = 0;
        ulong squared = MagnitudeSquared(0, second, third);
        if (squared == 0) return false;
        long product = (long)first * (sine ? third : second);
        return RootRatioFloor((BigInteger)product * product, squared, product < 0, out value);
    }

    private static bool RootRatioFloor(BigInteger numerator, ulong denominator, bool negative, out int value) {
        // Ratios from signed-int SG operands have magnitudes no greater than
        // Int32's magnitude. Integer division before the square root preserves
        // its floor; exact remainder comparison distinguishes negative ties.
        ulong root = MagnitudeFloor((ulong)(numerator / denominator));
        bool integral = (BigInteger)root * root * denominator == numerator;
        long result = negative ? -(long)root - (integral ? 0 : 1) : (long)root;
        value = 0;
        if (result < int.MinValue || result > int.MaxValue) return false;
        value = (int)result; return true;
    }

    private static ulong MagnitudeFloor(ulong squared) {
        // Three signed-int squares fit in UInt64. Verify the floating estimate
        // against the exact sum so neither side of an integer boundary is lost.
        ulong root = (ulong)Math.Sqrt(squared);
        while (root != 0 && root > squared / root) root--;
        while (root + 1 <= squared / (root + 1)) root++;
        return root;
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

    private static ushort U16(OfficeArtComplexData bytes, int offset) => unchecked((ushort)(bytes[offset] | bytes[offset + 1] << 8));
}
