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
            double logDensity = LogGamma((degrees + 1) / 2) - LogGamma(degrees / 2)
                - 0.5 * Math.Log(degrees * Math.PI) - (degrees + 1) / 2 * Math.Log(1 + t * t / degrees);
            return Math.Exp(logDensity);
        }

        /// <summary>The left-tailed Student t distribution.</summary>
        internal static double StudentCumulative(double t, double degrees) {
            double tail = StudentTail(t, degrees);
            return t > 0 ? 1 - tail : tail;
        }

        /// <summary>The probability beyond |t| in one tail.</summary>
        internal static double StudentTail(double t, double degrees) {
            if (t == 0) return 0.5;
            double x = degrees / (degrees + t * t);
            return 0.5 * RegularizedBeta(x, degrees / 2, 0.5);
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
            return Math.Exp((k - 1) * Math.Log(x) - x / 2 - k * Math.Log(2) - LogGamma(k));
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
            if (x <= lower || x >= upper) x = lower + (upper - lower) / 2;

            for (int iteration = 0; iteration < 500; iteration++) {
                double value = function(x) - target;
                if (value == 0) return x;
                if (value < 0) lower = x; else upper = x;
                double slope = derivative(x);
                double next = slope > 0 && !double.IsInfinity(slope) ? x - value / slope : double.NaN;
                if (double.IsNaN(next) || next <= lower || next >= upper) next = lower + (upper - lower) / 2;
                if (Math.Abs(next - x) <= 2 * Epsilon * Math.Abs(x) || upper - lower <= 2 * Epsilon * Math.Abs(next)) return next;
                x = next;
            }
            return x;
        }

        /// <summary>ln Γ(x) for x &gt; 0: Stirling's series after shifting the argument to 15 or more.</summary>
        internal static double LogGamma(double x) {
            if (x <= 0 || double.IsNaN(x)) return double.NaN;
            double product = 1;
            while (x < 15) {
                product *= x;
                x += 1;
            }
            double inverse = 1 / x;
            double inverseSquared = inverse * inverse;
            // B₂ₙ / (2n(2n−1) x^(2n−1)) for n = 1..7.
            double series = inverse * (1.0 / 12 - inverseSquared * (1.0 / 360 - inverseSquared * (1.0 / 1260
                - inverseSquared * (1.0 / 1680 - inverseSquared * (1.0 / 1188 - inverseSquared * (691.0 / 360360
                - inverseSquared * (1.0 / 156)))))));
            return (x - 0.5) * Math.Log(x) - x + 0.5 * Math.Log(2 * Math.PI) + series - Math.Log(product);
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

        private static double GammaPrefix(double a, double x) => Math.Exp(a * Math.Log(x) - x - LogGamma(a));

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
            double prefix = Math.Exp(LogGamma(a + b) - LogGamma(a) - LogGamma(b) + a * Math.Log(x) + b * Math.Log(1 - x));
            // The continued fraction converges quickly below the mean; above it, use I_x(a,b) = 1 − I_(1−x)(b,a).
            return x < (a + 1) / (a + b + 2)
                ? prefix * BetaContinuedFraction(x, a, b) / a
                : 1 - prefix * BetaContinuedFraction(1 - x, b, a) / b;
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
                double even = m * (b - m) * x / ((a + twoM - 1) * (a + twoM));
                d = 1 + even * d;
                if (Math.Abs(d) < Tiny) d = Tiny;
                c = 1 + even / c;
                if (Math.Abs(c) < Tiny) c = Tiny;
                d = 1 / d;
                result *= d * c;
                double odd = -(a + m) * (a + b + m) * x / ((a + twoM) * (a + twoM + 1));
                d = 1 + odd * d;
                if (Math.Abs(d) < Tiny) d = Tiny;
                c = 1 + odd / c;
                if (Math.Abs(c) < Tiny) c = Tiny;
                d = 1 / d;
                double delta = d * c;
                result *= delta;
                if (Math.Abs(delta - 1) < Epsilon) break;
            }
            return result;
        }
    }
}
