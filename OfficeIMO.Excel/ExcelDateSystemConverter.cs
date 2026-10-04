namespace OfficeIMO.Excel {
    /// <summary>
    /// Converts dates to and from Excel serial values using the workbook's configured date system.
    /// </summary>
    public static class ExcelDateSystemConverter {
        private static readonly DateTime Early1900Epoch = new DateTime(1899, 12, 31);
        /// <summary>
        /// Number of days between the Excel 1900 and 1904 date-system epochs.
        /// </summary>
        public const double Date1904OffsetDays = 1462d;

        /// <summary>
        /// Converts a date to the serial value used by the selected Excel date system.
        /// </summary>
        public static double ToSerial(DateTime value, ExcelDateSystem dateSystem) {
            if (dateSystem == ExcelDateSystem.NineteenHundred && value < new DateTime(1900, 3, 1))
                return (value - Early1900Epoch).TotalDays;
            double serial = value.ToOADate();
            return dateSystem == ExcelDateSystem.NineteenFour ? serial - Date1904OffsetDays : serial;
        }

        /// <summary>
        /// Converts an Excel serial value from the selected date system to a date.
        /// In the 1900 system, serial 60 denotes Excel's fictitious February 29 and
        /// maps to February 28 because <see cref="DateTime"/> cannot represent it.
        /// Negative 1900-system serials extend the December 31, 1899 epoch backwards,
        /// with fractional values representing elapsed portions of a day.
        /// </summary>
        public static DateTime FromSerial(double serial, ExcelDateSystem dateSystem) {
            if (dateSystem == ExcelDateSystem.NineteenHundred && serial < 60d)
                return Early1900Epoch.AddDays(serial);
            double oa = dateSystem == ExcelDateSystem.NineteenFour ? serial + Date1904OffsetDays : serial;
            return DateTime.FromOADate(oa);
        }

        /// <summary>Resolves a calendar value's format section without discarding its submillisecond ticks.</summary>
        internal static double ToSerialForDisplay(DateTime date, ExcelDateSystem dateSystem) =>
            (date - (dateSystem == ExcelDateSystem.NineteenFour ? new DateTime(1904, 1, 1)
                : date < new DateTime(1900, 3, 1) ? Early1900Epoch : new DateTime(1899, 12, 30))).TotalDays;

        /// <summary>Resolves a display value without millisecond rounding of the source serial.</summary>
        internal static DateTime FromSerialForDisplay(double serial, ExcelDateSystem dateSystem) {
            // Retain the public converter's range checks and epoch semantics. Typed
            // reads keep their established DateTime contract; display keeps raw precision.
            _ = FromSerial(serial, dateSystem);
            if (dateSystem == ExcelDateSystem.NineteenHundred && serial < 60d) {
                return Early1900Epoch.AddTicks((long)Math.Round(serial * TimeSpan.TicksPerDay));
            }
            double whole = Math.Truncate(serial);
            double fraction = serial - whole;
            if (dateSystem == ExcelDateSystem.NineteenFour) whole += Date1904OffsetDays;
            // An integral epoch shift can change the sign of an OA serial. Normalize
            // before applying OA's absolute time-of-day rule for negative serials.
            if (whole > 0 && fraction < 0) { whole--; fraction++; }
            else if (whole < 0 && fraction > 0) { whole++; fraction--; }
            long ticks = (long)Math.Round(Math.Abs(fraction) * TimeSpan.TicksPerDay);
            // DateTime-authored millisecond boundaries can land a few ticks below
            // the intended clock value when encoded in a double day serial. Recover
            // only boundaries within half one representable serial step; retain
            // genuinely submillisecond values rather than rounding all inputs to ms.
            double magnitude = Math.Abs(serial);
            double next = BitConverter.Int64BitsToDouble(BitConverter.DoubleToInt64Bits(magnitude) + 1);
            double tolerance = (next - magnitude) * TimeSpan.TicksPerDay / 2;
            long millisecondTicks = (long)Math.Round(ticks / (double)TimeSpan.TicksPerMillisecond) * TimeSpan.TicksPerMillisecond;
            if (Math.Abs(ticks - millisecondTicks) <= tolerance) ticks = millisecondTicks;
            return DateTime.FromOADate(whole).AddTicks(ticks);
        }
    }
}
