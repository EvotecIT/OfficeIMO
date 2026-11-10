namespace OfficeIMO.Drawing;

/// <summary>
/// Sine, cosine and arctangent built from IEEE additions and multiplications only. <see cref="Math"/> forwards these to
/// the platform C library, whose last-bit results differ between Windows, Linux and macOS and between x64 and arm64; the
/// stroker uses these instead so round joins and caps produce the same geometry on every platform. Accuracy is within a
/// few units in the last place for the angles the stroker uses.
/// </summary>
internal static class OfficeDeterministicTrig {
    private const double HalfPi = 1.5707963267948966;
    // Cody-Waite split of pi/2 so the range reduction stays exact for the angles strokes need.
    private const double HalfPiHigh = 1.5707963267341256;
    private const double HalfPiLow = 6.0771005065061922e-11;
    private const double SixthPi = 0.52359877559829887;
    private const double Sqrt3 = 1.7320508075688772;
    private const double TanTwelfthPi = 0.26794919243112270;

    public static double Sin(double angle) {
        double reduced = Reduce(angle, out int quadrant);
        return quadrant switch {
            0 => SinKernel(reduced),
            1 => CosKernel(reduced),
            2 => -SinKernel(reduced),
            _ => -CosKernel(reduced)
        };
    }

    public static double Cos(double angle) {
        double reduced = Reduce(angle, out int quadrant);
        return quadrant switch {
            0 => CosKernel(reduced),
            1 => -SinKernel(reduced),
            2 => -CosKernel(reduced),
            _ => SinKernel(reduced)
        };
    }

    public static double Atan2(double y, double x) {
        if (double.IsNaN(x) || double.IsNaN(y)) return double.NaN;
        if (x == 0D && y == 0D) return 0D;
        double ax = Math.Abs(x), ay = Math.Abs(y);
        double angle = ay <= ax ? Atan(ay / ax) : HalfPi - Atan(ax / ay);
        if (x < 0D) angle = Math.PI - angle;
        return y < 0D ? -angle : angle;
    }

    /// <summary>Arctangent of a value in [0, 1].</summary>
    private static double Atan(double value) {
        double offset = 0D;
        if (value > TanTwelfthPi) {
            // atan(t) = pi/6 + atan((t*sqrt3 - 1) / (t + sqrt3)) keeps the series argument below tan(pi/12).
            value = (value * Sqrt3 - 1D) / (value + Sqrt3);
            offset = SixthPi;
        }
        double square = value * value;
        double sum = 0D;
        for (int term = 31; term >= 1; term -= 2) {
            sum = 1D / term - square * sum;
        }
        return offset + value * sum;
    }

    private static double Reduce(double angle, out int quadrant) {
        double turns = Math.Round(angle / HalfPi);
        double reduced = angle - turns * HalfPiHigh - turns * HalfPiLow;
        long whole = (long)turns;
        quadrant = (int)(((whole % 4) + 4) % 4);
        return reduced;
    }

    /// <summary>Taylor series of sine for |x| &lt;= pi/4, evaluated in a fixed order.</summary>
    private static double SinKernel(double x) {
        double square = x * x;
        double sum = 0D;
        for (int n = 17; n >= 3; n -= 2) {
            sum = (1D - square * sum) / ((n - 1) * n);
        }
        return x - x * square * sum;
    }

    /// <summary>Taylor series of cosine for |x| &lt;= pi/4, evaluated in a fixed order.</summary>
    private static double CosKernel(double x) {
        double square = x * x;
        double sum = 0D;
        for (int n = 18; n >= 2; n -= 2) {
            sum = (1D - square * sum) / ((n - 1) * n);
        }
        return 1D - square * sum;
    }
}
