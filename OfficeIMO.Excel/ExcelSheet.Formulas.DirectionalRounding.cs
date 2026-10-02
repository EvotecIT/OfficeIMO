namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static double RoundDirectionalAtDigits(double value, int digits, bool awayFromZero) {
            // Integral doubles need no positive-digit scaling, which can overflow.
            if (digits >= 0 && value == Math.Truncate(value)) return value;
            double factor = Math.Pow(10, digits);
            // Use the same decimal normalization as ROUND to avoid binary artifacts
            // moving an exact decimal boundary (for example, 0.29 at two places).
            if (Math.Abs(value) < (double)decimal.MaxValue) {
                try {
                    decimal number = (decimal)value;
                    // Preserve nonzero inputs below decimal's precision for ROUNDUP.
                    if (number != 0 || value == 0) {
                        decimal scale = (decimal)factor;
                        decimal shifted = number * scale;
                        if (awayFromZero && shifted == 0 && number != 0) return Math.Sign(value) * Math.Pow(10, -digits);
                        decimal rounded = awayFromZero
                            ? Math.Sign(shifted) * decimal.Ceiling(Math.Abs(shifted)) : decimal.Truncate(shifted);
                        return (double)(rounded / scale);
                    }
                } catch (OverflowException) {
                    // Decimal scaling can exceed its range; doubles still represent it.
                }
            }
            double scaled = Math.Abs(value) * factor;
            if (awayFromZero && scaled == 0 && value != 0) return Math.Sign(value) * Math.Pow(10, -digits);
            double integral = awayFromZero ? Math.Ceiling(scaled) : Math.Truncate(scaled);
            return Math.Sign(value) * integral / factor;
        }
    }
}
