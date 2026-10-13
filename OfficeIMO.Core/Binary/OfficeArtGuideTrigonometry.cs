using System;
using System.Numerics;

namespace OfficeIMO.Drawing.Binary;

/// <summary>
/// Evaluates the bounded, signed-integer SG trigonometric operands with directed
/// fixed-point intervals. An integer is returned only when both bounds have the
/// same floor; uncertainty follows the importer's existing invalid-guide path.
/// </summary>
internal static class OfficeArtGuideTrigonometry {
    private const int QuarterTurn = 90 * 65536;
    private const int FullTurn = 4 * QuarterTurn;
    private static readonly BigInteger Scale = BigInteger.Pow(10, 80);
    private static readonly Bounds One = Integer(1);
    // Machin's identity, using rational arguments and alternating-series bounds,
    // avoids treating a rounded binary64 PI constant as an exact input.
    private static readonly Bounds Pi = AtanSeries(Ratio(1, 5)).Times(16) - AtanSeries(Ratio(1, 239)).Times(4);

    internal static bool TryCalculate(int kind, int first, int second, out int value) {
        value = 0;
        if (kind == 8) return TryAtan2(first, second, out value);
        int turn = second % FullTurn;
        if (turn < 0) turn += FullTurn;
        if (kind == 16 && (turn == QuarterTurn || turn == 3 * QuarterTurn)) return false;
        if (first == 0) return true;
        if (turn % QuarterTurn == 0) {
            int quadrant = turn / QuarterTurn;
            int factor = kind == 9 ? (quadrant == 1 ? 1 : quadrant == 3 ? -1 : 0)
                : kind == 10 ? (quadrant == 0 ? 1 : quadrant == 2 ? -1 : 0) : 0;
            return TryFloor(Integer(first).Times(factor), out value);
        }
        if (kind == 9 && turn % (60 * 65536) == 30 * 65536)
            return TryFloor(Ratio((long)first * (turn < 2 * QuarterTurn ? 1 : -1), 2), out value);
        if (kind == 10 && turn % (60 * 65536) == 0)
            return TryFloor(Ratio((long)first * (turn < QuarterTurn || turn > 3 * QuarterTurn ? 1 : -1), 2), out value);
        if (kind == 16 && turn % QuarterTurn == QuarterTurn / 2)
            return TryFloor(Integer(first).Times((turn / QuarterTurn & 1) == 0 ? 1 : -1), out value);
        SinCos(turn, out Bounds sine, out Bounds cosine);
        Bounds coefficient = kind == 9 ? sine : kind == 10 ? cosine : Divide(sine, cosine);
        return TryFloor(coefficient.Times(first), out value);
    }

    private static bool TryAtan2(int x, int y, out int value) {
        value = 0;
        if (x == 0 && y == 0) return false;
        if (y == 0) { value = x < 0 ? 2 * QuarterTurn : 0; return true; }
        if (x == 0) { value = y < 0 ? -QuarterTurn : QuarterTurn; return true; }
        long ax = Math.Abs((long)x), ay = Math.Abs((long)y);
        if (ax == ay) {
            value = (x < 0 ? 3 : 1) * (QuarterTurn / 2) * (y < 0 ? -1 : 1);
            return true;
        }
        Bounds ratio = Ratio(Math.Min(ax, ay), Math.Max(ax, ay));
        // Half-angle reduction keeps the alternating series below tan(PI/8).
        Bounds reduced = Divide(ratio, One + Sqrt(One + Multiply(ratio, ratio)));
        Bounds angle = AtanSeries(reduced).Times(2);
        if (ay > ax) angle = Pi.Over(2) - angle;
        if (x < 0) angle = Pi - angle;
        if (y < 0) angle = -angle;
        return TryFloor(Divide(angle.Times(2 * QuarterTurn), Pi), out value);
    }

    private static void SinCos(int turn, out Bounds sine, out Bounds cosine) {
        int quadrant = turn / QuarterTurn, offset = turn % QuarterTurn;
        bool reflected = offset > QuarterTurn / 2;
        if (reflected) offset = QuarterTurn - offset;
        Bounds radians = Pi.Times(offset).Over(2 * QuarterTurn);
        Bounds squared = Multiply(radians, radians);
        Bounds sinTerm = radians, cosTerm = One;
        Bounds sin = sinTerm, cos = cosTerm;
        // On [0, PI/4], 48 terms leave a remainder below one fixed-point unit.
        // Rounded interval arithmetic and the final alternating-term envelope
        // also enclose all intermediate rounding error.
        for (int index = 1; index <= 48; index++) {
            sinTerm = Multiply(sinTerm, squared).Over(2 * index * (2 * index + 1));
            cosTerm = Multiply(cosTerm, squared).Over((2 * index - 1) * 2 * index);
            sin = (index & 1) == 0 ? sin + sinTerm : sin - sinTerm;
            cos = (index & 1) == 0 ? cos + cosTerm : cos - cosTerm;
        }
        sin = sin.Expand(sinTerm.Upper); cos = cos.Expand(cosTerm.Upper);
        if (reflected) { Bounds saved = sin; sin = cos; cos = saved; }
        sine = quadrant == 0 ? sin : quadrant == 1 ? cos : quadrant == 2 ? -sin : -cos;
        cosine = quadrant == 0 ? cos : quadrant == 1 ? -sin : quadrant == 2 ? -cos : sin;
    }

    private static Bounds AtanSeries(Bounds argument) {
        Bounds squared = Multiply(argument, argument), power = argument, sum = argument, term = argument;
        // Arguments are at most tan(PI/8); 128 terms bound the remainder well
        // below the 80-digit scale. PI construction uses still smaller ratios.
        for (int index = 1; index <= 128; index++) {
            power = Multiply(power, squared); term = power.Over(2 * index + 1);
            sum = (index & 1) == 0 ? sum + term : sum - term;
        }
        return sum.Expand(term.Upper);
    }

    private static bool TryFloor(Bounds bounds, out int value) {
        BigInteger lower = Floor(bounds.Lower, Scale), upper = Floor(bounds.Upper, Scale);
        value = 0;
        if (lower != upper || lower < int.MinValue || lower > int.MaxValue) return false;
        value = (int)lower; return true;
    }

    private static Bounds Integer(long value) => new Bounds(value * Scale, value * Scale);
    private static Bounds Ratio(long numerator, long denominator) => Integer(numerator).Over(denominator);

    private static Bounds Multiply(Bounds left, Bounds right) {
        BigInteger a = left.Lower * right.Lower, b = left.Lower * right.Upper;
        BigInteger c = left.Upper * right.Lower, d = left.Upper * right.Upper;
        return new Bounds(Floor(BigInteger.Min(BigInteger.Min(a, b), BigInteger.Min(c, d)), Scale),
            Ceiling(BigInteger.Max(BigInteger.Max(a, b), BigInteger.Max(c, d)), Scale));
    }

    private static Bounds Divide(Bounds numerator, Bounds denominator) {
        if (denominator.Upper < 0) return -Divide(numerator, -denominator);
        if (denominator.Lower <= 0) throw new InvalidOperationException("A trigonometric interval includes zero.");
        BigInteger a = numerator.Lower * Scale, b = numerator.Upper * Scale;
        return new Bounds(BigInteger.Min(Floor(a, denominator.Lower), Floor(a, denominator.Upper)),
            BigInteger.Max(Ceiling(b, denominator.Lower), Ceiling(b, denominator.Upper)));
    }

    private static Bounds Sqrt(Bounds value) {
        BigInteger lower = IntegerSqrt(value.Lower * Scale), argument = value.Upper * Scale;
        BigInteger upper = IntegerSqrt(argument);
        if (upper * upper != argument) upper++;
        return new Bounds(lower, upper);
    }

    private static BigInteger IntegerSqrt(BigInteger value) {
        if (value.IsZero) return BigInteger.Zero;
        BigInteger root = BigInteger.One << ((value.ToByteArray().Length * 8 + 1) / 2);
        while (true) {
            BigInteger next = (root + value / root) / 2;
            if (next >= root) return root;
            root = next;
        }
    }

    private static BigInteger Floor(BigInteger numerator, BigInteger denominator) {
        BigInteger quotient = BigInteger.DivRem(numerator, denominator, out BigInteger remainder);
        return remainder.Sign < 0 ? quotient - 1 : quotient;
    }
    private static BigInteger Ceiling(BigInteger numerator, BigInteger denominator) => -Floor(-numerator, denominator);

    private readonly struct Bounds {
        internal Bounds(BigInteger lower, BigInteger upper) { Lower = lower; Upper = upper; }
        internal BigInteger Lower { get; }
        internal BigInteger Upper { get; }
        internal Bounds Times(long value) => value < 0 ? new Bounds(Upper * value, Lower * value) : new Bounds(Lower * value, Upper * value);
        internal Bounds Over(long value) => new Bounds(Floor(Lower, value), Ceiling(Upper, value));
        internal Bounds Expand(BigInteger error) => new Bounds(Lower - error, Upper + error);
        public static Bounds operator +(Bounds a, Bounds b) => new Bounds(a.Lower + b.Lower, a.Upper + b.Upper);
        public static Bounds operator -(Bounds a, Bounds b) => new Bounds(a.Lower - b.Upper, a.Upper - b.Lower);
        public static Bounds operator -(Bounds value) => new Bounds(-value.Upper, -value.Lower);
    }
}
