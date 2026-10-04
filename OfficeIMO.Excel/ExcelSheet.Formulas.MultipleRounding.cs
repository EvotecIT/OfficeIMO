using System.Globalization;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Legacy CEILING/FLOOR use the sign of the significance to choose the
        // direction; CEILING.MATH/FLOOR.MATH have a separate contract.
        private static double RoundAtMultiple(double value, double significance, bool ceiling) {
            bool awayFromZero = (ceiling == (significance > 0)) == (value > 0);
            double factor = Math.Abs(significance);
            if (Math.Abs(value) < (double)decimal.MaxValue && factor < (double)decimal.MaxValue) {
                try {
                    decimal number = (decimal)value, multiple = (decimal)factor;
                    // Decimal has 28 fractional places: below 1E-14 it may
                    // lose more than the intended 15-significant-digit normalization.
                    bool numberKeepsPrecision = Math.Abs(value) >= 1E-14 || DecimalToDouble(number) == value;
                    bool factorKeepsPrecision = factor >= 1E-14 || DecimalToDouble(multiple) == factor;
                    if (number != 0 && multiple != 0 && numberKeepsPrecision && factorKeepsPrecision) {
                        // Remainders avoid both quotient overflow and decimal division
                        // rounding a near-integer quotient across its boundary.
                        decimal remainder = number % multiple;
                        decimal rounded = number - remainder;
                        if (awayFromZero && remainder != 0) rounded += Math.Sign(value) * multiple;
                        return DecimalToDouble(rounded);
                    }
                } catch (OverflowException) {
                    // The finite double result can exceed decimal's range.
                }
            }
            // Extremely small factors or large inputs can overflow value/factor
            // even when the rounded multiple is finite. The remainder stays bounded.
            double quotient = value / factor;
            if (!double.IsInfinity(quotient)) {
                double nearest = Math.Round(quotient);
                // Normalize division noise at a nonzero integral boundary using
                // the same 15-significant-digit precision as the decimal path.
                // Do not treat an underflowed quotient as an exact zero multiple.
                if (nearest != 0 && Math.Abs(quotient - nearest) <= Math.Abs(quotient) * 1E-15) return value;
            }
            double rest = value % factor;
            double result = value - rest;
            if (awayFromZero && rest != 0) result += Math.Sign(value) * factor;
            return result;
        }

        // Parse the exact decimal text once: decimal-to-double multiplication
        // can introduce an extra rounding step for values such as 2E-28.
        private static double DecimalToDouble(decimal value) =>
            double.Parse(value.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
    }
}
