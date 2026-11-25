namespace OfficeIMO.Excel {
    /// <summary>
    /// Normal, Student t and chi-squared distributions for the formula evaluator. Cumulative functions come from the
    /// regularized incomplete gamma and beta functions; inverses refine a starting point with Newton steps inside a
    /// shrinking bracket, so every inverse agrees with its own cumulative function to double precision.
    /// </summary>
    internal static class ExcelFormulaDistributions {
        private const double Epsilon = 1e-16;
        private const double Tiny = 1e-300;
        private const int MinimumIterations = 1000;

        /// <summary>Thrown when a series or continued fraction has not converged; the formula is then left uncalculated.</summary>
        internal sealed class NonConvergenceException : ArithmeticException {
            internal NonConvergenceException() : base("A distribution series did not converge.") { }
        }

        // Near the mean the gamma and beta expansions need on the order of √a terms, so the limit grows with the parameter.
        private static int IterationLimit(double parameter) =>
            (int)Math.Min(int.MaxValue / 2, MinimumIterations + 40 * Math.Sqrt(Math.Max(0, parameter)));
        private static readonly double SquareRootOfTwoPi = Math.Sqrt(2 * Math.PI);

        internal static double NormalDensity(double z) => Math.Exp(-0.5 * z * z) / SquareRootOfTwoPi;

        /// <summary>The standard normal distribution, accurate in both tails.</summary>
        internal static double NormalCumulative(double z) {
            if (double.IsNaN(z)) return double.NaN;
            // Φ(z) = Q(1/2, z²/2) / 2 for z ≤ 0, so the lower tail keeps its relative accuracy.
            double half = 0.5 * RegularizedGammaUpper(0.5, 0.5 * z * z);
            return z <= 0 ? half : 1 - half;
        }

        internal static double NormalInverse(double p) {
            if (p <= 0 || p >= 1 || double.IsNaN(p)) return double.NaN;
            if (p > 0.5) return -NormalInverse(1 - p);
            if (p == 0.5) return 0;
            // A rational first guess (error below 5e-4), then Halley steps on the accurate cumulative function.
            double t = Math.Sqrt(-2 * Math.Log(p));
            double z = -(t - (2.515517 + 0.802853 * t + 0.010328 * t * t)
                / (1 + 1.432788 * t + 0.189269 * t * t + 0.001308 * t * t * t));
            for (int iteration = 0; iteration < 8; iteration++) {
                double error = NormalCumulative(z) - p;
                double u = error / NormalDensity(z);
                double step = u / (1 + 0.5 * z * u);
                z -= step;
                if (Math.Abs(step) <= 1e-15 * Math.Max(1, Math.Abs(z))) break;
            }
            return z;
        }

        internal static double StudentDensity(double t, double degrees) {
            double logDensity = LogGammaHalfRatio(degrees / 2) - 0.5 * Math.Log(degrees * Math.PI)
                - (degrees + 1) / 2 * LogOnePlusSquare(t / Math.Sqrt(degrees));
            return Math.Exp(logDensity);
        }

        /// <summary>ln(1 + r²) without overflowing r² or rounding away a small r.</summary>
        private static double LogOnePlusSquare(double r) {
            r = Math.Abs(r);
            if (r <= 1) return LogOnePlus(r * r);
            double s = 1 / r;
            return 2 * Math.Log(r) + LogOnePlus(s * s);
        }

        /// <summary>ln(1 + v), accurate for small v.</summary>
        private static double LogOnePlus(double v) {
            double u = 1 + v;
            return u == 1 ? v : Math.Log(u) * v / (u - 1);
        }

        /// <summary>ln Γ(z + ½) − ln Γ(z); for large z the two terms round alike, so an asymptotic series is used there.</summary>
        private static double LogGammaHalfRatio(double z) {
            if (z < 100) return LogGamma(z + 0.5) - LogGamma(z);
            double inverse = 1 / z;
            double inverseSquared = inverse * inverse;
            return 0.5 * Math.Log(z) - inverse * (1.0 / 8 - inverseSquared * (1.0 / 192 + inverseSquared * (1.0 / 640
                - inverseSquared * (17.0 / 14336))));
        }

        /// <summary>The left-tailed Student t distribution.</summary>
        internal static double StudentCumulative(double t, double degrees) {
            double tail = StudentTail(t, degrees);
            return t > 0 ? 1 - tail : tail;
        }

        /// <summary>The probability beyond |t| in one tail.</summary>
        internal static double StudentTail(double t, double degrees) {
            if (t == 0) return 0.5;
            // The t tail differs from the normal one by a relative (t⁴ + 2t²)/(4ν) + O(ν⁻²), so at huge degrees,
            // where x below rounds to 1 and the beta form loses the distance, the normal tail is exact to double precision.
            double squared = t * t;
            if (degrees >= 1e15 && squared * squared + 2 * squared <= 4e-16 * degrees) return NormalCumulative(-Math.Abs(t));
            // I_x(ν/2, ½) with x = ν/(ν+t²). Writing r = |t|/√ν, x and 1 − x come from r or 1/r, whichever is at most 1,
            // so neither t² overflowing nor x rounding to 1 near the median loses the result.
            double r = Math.Abs(t) / Math.Sqrt(degrees);
            double x, complement, logX, logComplement;
            if (r <= 1) {
                double square = r * r;
                x = 1 / (1 + square);
                complement = square / (1 + square);
                logX = -LogOnePlus(square);
                logComplement = 2 * Math.Log(r) + logX;
            } else {
                double s = 1 / r;
                double square = s * s;
                complement = 1 / (1 + square);
                x = square / (1 + square);
                logComplement = -LogOnePlus(square);
                logX = 2 * Math.Log(s) + logComplement;
            }
            return 0.5 * RegularizedBeta(x, complement, logX, logComplement, degrees / 2, 0.5);
        }

        internal static double StudentInverse(double p, double degrees) {
            if (p <= 0 || p >= 1 || double.IsNaN(p)) return double.NaN;
            if (p == 0.5) return 0;
            if (p > 0.5) return -StudentInverse(1 - p, degrees);
            // On t < 0 the cumulative function is the one-tail probability, which keeps small p precise.
            double guess = Math.Min(-1e-8, NormalInverse(p));
            return SolveIncreasing(t => StudentTail(t, degrees), t => StudentDensity(t, degrees), p, guess, double.NegativeInfinity, 0);
        }

        internal static double ChiSquaredDensity(double x, double degrees) {
            if (x < 0) return 0;
            double k = degrees / 2;
            if (x == 0) return k == 1 ? 0.5 : k > 1 ? 0 : double.PositiveInfinity;
            // x^(k−1) e^(−x/2) / (2ᵏ Γ(k)) = P's prefix at (k, x/2) over x, which stays exact near the mean; the
            // division happens in logarithms so a tiny x cannot underflow the prefix first.
            return Math.Exp(LogGammaPrefix(k, x / 2) - Math.Log(x));
        }

        internal static double ChiSquaredCumulative(double x, double degrees) =>
            x <= 0 ? 0 : RegularizedGammaLower(degrees / 2, x / 2);

        internal static double ChiSquaredUpper(double x, double degrees) =>
            x <= 0 ? 1 : RegularizedGammaUpper(degrees / 2, x / 2);

        internal static double ChiSquaredInverse(double p, double degrees) {
            if (p < 0 || p > 1 || double.IsNaN(p)) return double.NaN;
            if (p == 0) return 0;
            if (p == 1) return double.PositiveInfinity;
            return SolveIncreasing(x => ChiSquaredCumulative(x, degrees), x => ChiSquaredDensity(x, degrees), p,
                ChiSquaredGuess(p, degrees), 0, double.PositiveInfinity);
        }

        internal static double ChiSquaredUpperInverse(double p, double degrees) {
            if (p < 0 || p > 1 || double.IsNaN(p)) return double.NaN;
            if (p == 1) return 0;
            if (p == 0) return double.PositiveInfinity;
            // Solve on the upper tail itself so small right-tail probabilities keep their precision.
            return SolveIncreasing(x => -ChiSquaredUpper(x, degrees), x => ChiSquaredDensity(x, degrees), -p,
                ChiSquaredGuess(1 - p, degrees), 0, double.PositiveInfinity);
        }

        // Wilson–Hilferty: the cube root of χ²/k is close to normal.
        private static double ChiSquaredGuess(double p, double degrees) {
            if (p <= 0 || p >= 1) return degrees;
            double z = NormalInverse(p);
            double h = 2 / (9 * degrees);
            return degrees * Math.Pow(Math.Max(1e-6, 1 - h + z * Math.Sqrt(h)), 3);
        }

        /// <summary>
        /// Finds x in (lower, upper) where an increasing <paramref name="function"/> equals <paramref name="target"/>.
        /// A Newton step that would leave the current bracket is replaced by bisection.
        /// </summary>
        private static double SolveIncreasing(Func<double, double> function, Func<double, double> derivative, double target,
            double guess, double lower, double upper) {
            double x = guess;
            if (double.IsNegativeInfinity(lower) || double.IsPositiveInfinity(upper)) {
                // Step away from the guess, doubling each time, until the root is bracketed.
                double step = Math.Max(1, Math.Abs(guess));
                for (int expand = 0; expand < 2000; expand++) {
                    double value = function(x) - target;
                    if (value == 0) return x;
                    if (value < 0) {
                        lower = x;
                        if (!double.IsPositiveInfinity(upper)) break;
                        x += step;
                    } else {
                        upper = x;
                        if (!double.IsNegativeInfinity(lower)) break;
                        x -= step;
                    }
                    step *= 2;
                }
                if (double.IsInfinity(lower) || double.IsInfinity(upper)) return double.NaN;
            }
            if (x <= lower || x >= upper) x = Bisect(lower, upper);

            for (int iteration = 0; iteration < 500; iteration++) {
                double value = function(x) - target;
                if (value == 0) return x;
                if (value < 0) lower = x; else upper = x;
                double slope = derivative(x);
                double next = slope > 0 && !double.IsInfinity(slope) ? x - value / slope : double.NaN;
                if (double.IsNaN(next) || next <= lower || next >= upper) next = Bisect(lower, upper);
                if (Math.Abs(next - x) <= 2 * Epsilon * Math.Abs(x) || upper - lower <= 2 * Epsilon * Math.Abs(next)) return next;
                x = next;
            }
            throw new NonConvergenceException();
        }

        /// <summary>
        /// Splits a bracket on one side of zero by magnitude when its ends differ by orders of magnitude, so a root in
        /// an extreme tail (a quantile near 1e-200, say) is reached in a bounded number of steps.
        /// </summary>
        private static double Bisect(double lower, double upper) {
            if (lower >= 0 && upper > 4 * lower) return lower == 0 ? upper / 64 : Math.Sqrt(lower) * Math.Sqrt(upper);
            if (upper <= 0 && lower < 4 * upper) return upper == 0 ? lower / 64 : -Math.Sqrt(-lower) * Math.Sqrt(-upper);
            return lower + (upper - lower) / 2;
        }

        /// <summary>ln Γ(x) for x &gt; 0: Stirling's series after shifting the argument to 15 or more.</summary>
        internal static double LogGamma(double x) {
            if (x <= 0 || double.IsNaN(x)) return double.NaN;
            double product = 1;
            while (x < 15) {
                product *= x;
                x += 1;
            }
            return (x - 0.5) * Math.Log(x) - x + 0.5 * Math.Log(2 * Math.PI) + StirlingCorrection(x) - Math.Log(product);
        }

        /// <summary>ln Γ(x) − ((x − ½) ln x − x + ½ ln 2π): B₂ₙ / (2n(2n−1) x^(2n−1)) for n = 1..7, for x ≥ 15.</summary>
        private static double StirlingCorrection(double x) {
            double inverse = 1 / x;
            double inverseSquared = inverse * inverse;
            return inverse * (1.0 / 12 - inverseSquared * (1.0 / 360 - inverseSquared * (1.0 / 1260
                - inverseSquared * (1.0 / 1680 - inverseSquared * (1.0 / 1188 - inverseSquared * (691.0 / 360360
                - inverseSquared * (1.0 / 156)))))));
        }

        /// <summary>The regularized lower incomplete gamma function P(a, x).</summary>
        internal static double RegularizedGammaLower(double a, double x) {
            if (x <= 0) return 0;
            if (double.IsPositiveInfinity(x)) return 1;
            return x < a + 1 ? GammaSeries(a, x) : 1 - GammaContinuedFraction(a, x);
        }

        /// <summary>The regularized upper incomplete gamma function Q(a, x).</summary>
        internal static double RegularizedGammaUpper(double a, double x) {
            if (x <= 0) return 1;
            if (double.IsPositiveInfinity(x)) return 0;
            return x < a + 1 ? 1 - GammaSeries(a, x) : GammaContinuedFraction(a, x);
        }

        /// <summary>
        /// xᵃ e⁻ˣ / Γ(a). For large a the terms a ln x, x and ln Γ(a) are huge and nearly cancel near the mean, so
        /// Stirling's form is used: √(a / 2π) · exp(a (ln u − u + 1) − correction(a)) with u = x / a.
        /// </summary>
        private static double GammaPrefix(double a, double x) => Math.Exp(LogGammaPrefix(a, x));

        /// <summary>ln(xᵃ e⁻ˣ / Γ(a)), so a caller can scale it before it underflows.</summary>
        private static double LogGammaPrefix(double a, double x) {
            if (a < 15) return a * Math.Log(x) - x - LogGamma(a);
            double d = (x - a) / a;
            // Away from the mean nothing cancels; taking ln x directly keeps a tiny x from rounding u to 0 and a
            // huge x from overflowing 1 + d.
            double exponent = Math.Abs(d) < 0.5
                ? a * LogOnePlusMinusSelf(d)
                : a * (Math.Log(x) - Math.Log(a)) - (x - a);
            return 0.5 * Math.Log(a / (2 * Math.PI)) + exponent - StirlingCorrection(a);
        }

        /// <summary>ln(1 + d) − d for |d| &lt; ½, summed as a series so the two terms do not cancel.</summary>
        private static double LogOnePlusMinusSelf(double d) {
            double power = d * d, sum = 0;
            for (int k = 2; k < 200; k++) {
                double term = power / k;
                sum += (k % 2 == 0) ? -term : term;
                if (Math.Abs(term) <= 1e-17 * Math.Abs(sum)) break;
                power *= d;
            }
            return sum;
        }

        // P(a, x) = xᵃ e⁻ˣ / Γ(a) · Σ xⁿ / (a(a+1)…(a+n)).
        private static double GammaSeries(double a, double x) {
            double term = 1 / a;
            double sum = term;
            int limit = IterationLimit(Math.Max(a, x));
            for (int n = 1; ; n++) {
                if (n >= limit) throw new NonConvergenceException();
                term *= x / (a + n);
                sum += term;
                if (Math.Abs(term) < Math.Abs(sum) * Epsilon) break;
            }
            return sum * GammaPrefix(a, x);
        }

        // Q(a, x) = xᵃ e⁻ˣ / Γ(a) · 1/(x+1−a− 1(1−a)/(x+3−a− 2(2−a)/(x+5−a− …))), by the modified Lentz method.
        private static double GammaContinuedFraction(double a, double x) {
            double b = x + 1 - a;
            double c = 1 / Tiny;
            double d = 1 / b;
            double result = d;
            int limit = IterationLimit(Math.Max(a, x));
            for (int n = 1; ; n++) {
                if (n >= limit) throw new NonConvergenceException();
                double an = -n * (n - a);
                b += 2;
                d = an * d + b;
                if (Math.Abs(d) < Tiny) d = Tiny;
                c = b + an / c;
                if (Math.Abs(c) < Tiny) c = Tiny;
                d = 1 / d;
                double delta = d * c;
                result *= delta;
                if (Math.Abs(delta - 1) < Epsilon) break;
            }
            return result * GammaPrefix(a, x);
        }

        /// <summary>The regularized incomplete beta function I_x(a, b).</summary>
        internal static double RegularizedBeta(double x, double a, double b) {
            if (x <= 0) return 0;
            if (x >= 1) return 1;
            return RegularizedBeta(x, 1 - x, Math.Log(x), Math.Log(1 - x), a, b);
        }

        /// <summary>
        /// I_x(a, b) from x, 1 − x and their logarithms, each computed by the caller without cancellation, so an
        /// argument that is tiny or within rounding of 1 keeps its precision.
        /// </summary>
        private static double RegularizedBeta(double x, double complement, double logX, double logComplement, double a, double b) {
            double prefix = Math.Exp(LogBetaNormalization(a, b) + a * logX + b * logComplement);
            // The continued fraction converges quickly below the mean; above it, use I_x(a,b) = 1 − I_(1−x)(b,a).
            // Compare on the side nearer 0, where neither x nor the mean has rounded to 1.
            bool belowMean = x <= 0.5 ? x < (a + 1) / (a + b + 2) : complement > (b + 1) / (a + b + 2);
            if (!belowMean) return 1 - prefix * BetaContinuedFraction(complement, b, a) / b;
            // With a huge a the fraction is sensitive to x at about a·ε; once x has rounded further than that from
            // 1 − complement (x = 1 itself, say) neither form is reliable, so the formula is left uncalculated.
            if (a * Math.Abs(1 - x - complement) > 1e-8) throw new NonConvergenceException();
            return prefix * BetaContinuedFraction(x, a, b) / a;
        }

        /// <summary>
        /// ln Γ(a+b) − ln Γ(a) − ln Γ(b). With one parameter ½ and the other huge, as in Student t at many degrees,
        /// a + b rounds to the larger one and the plain difference cancels, so that case uses the half-gamma ratio.
        /// </summary>
        private static double LogBetaNormalization(double a, double b) {
            if (b == 0.5) return LogGammaHalfRatio(a) - LogGamma(0.5);
            if (a == 0.5) return LogGammaHalfRatio(b) - LogGamma(0.5);
            return LogGamma(a + b) - LogGamma(a) - LogGamma(b);
        }

        // I_x(a, b)'s continued fraction, with d₂ₘ = m(b−m)x / ((a+2m−1)(a+2m)) and
        // d₂ₘ₊₁ = −(a+m)(a+b+m)x / ((a+2m)(a+2m+1)), by the modified Lentz method.
        private static double BetaContinuedFraction(double x, double a, double b) {
            double c = 1;
            double d = 1 - (a + b) * x / (a + 1);
            if (Math.Abs(d) < Tiny) d = Tiny;
            d = 1 / d;
            double result = d;
            int limit = IterationLimit(Math.Max(a, b));
            for (int m = 1; ; m++) {
                if (m >= limit) throw new NonConvergenceException();
                int twoM = 2 * m;
                double even = m * (b - m) / (a + twoM - 1) * x / (a + twoM);
                d = 1 + even * d;
                if (Math.Abs(d) < Tiny) d = Tiny;
                c = 1 + even / c;
                if (Math.Abs(c) < Tiny) c = Tiny;
                d = 1 / d;
                result *= d * c;
                double odd = -(a + m) / (a + twoM) * ((a + b + m) / (a + twoM + 1)) * x;
                d = 1 + odd * d;
                if (Math.Abs(d) < Tiny) d = Tiny;
                c = 1 + odd / c;
                if (Math.Abs(c) < Tiny) c = Tiny;
                d = 1 / d;
                double delta = d * c;
                result *= delta;
                // A non-finite state never settles, so stop rather than spend the whole iteration budget.
                if (double.IsNaN(result) || double.IsInfinity(result)) throw new NonConvergenceException();
                if (Math.Abs(delta - 1) < Epsilon) break;
            }
            return result;
        }
    }
}
