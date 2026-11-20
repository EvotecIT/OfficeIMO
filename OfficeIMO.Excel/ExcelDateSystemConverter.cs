using System.Numerics;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Converts dates to and from Excel serial values using the workbook's configured date system.
    /// </summary>
    /// <remarks>
    /// A serial counts days, and its whole part is the day Excel reports for it. Serials hold whole milliseconds:
    /// values are stored truncated to the millisecond, so a time just before midnight stays on its own day, and
    /// serials read back rounded to the nearest millisecond, never past the last millisecond of the serial's day.
    /// The time always comes from the serial's own fraction; a date system's epoch moves only whole days, as
    /// integer arithmetic, so the 1904 offset never changes how a fraction rounds.
    /// </remarks>
    public static class ExcelDateSystemConverter {
        private static readonly DateTime Early1900Epoch = new DateTime(1899, 12, 31);
        private static readonly DateTime OleAutomationEpoch = new DateTime(1899, 12, 30);
        private static readonly DateTime Epoch1904 = new DateTime(1904, 1, 1);
        private const long MillisecondsPerDay = 86_400_000L;
        private const double MaximumDateTimeDays = 3_652_059d;
        /// <summary>
        /// Number of days between the Excel 1900 and 1904 date-system epochs.
        /// </summary>
        public const double Date1904OffsetDays = 1462d;

        /// <summary>
        /// Converts a date to the serial value used by the selected Excel date system. Dates before serial 0,
        /// December 31, 1899 in the 1900 system and January 1, 1904 in the 1904 system, extend the epoch backwards
        /// as negative serials in both systems. Excel shows those cells as <c>####</c>; OfficeIMO reads them back
        /// as the same date.
        /// </summary>
        public static double ToSerial(DateTime value, ExcelDateSystem dateSystem) {
            if (dateSystem == ExcelDateSystem.NineteenHundred && value < new DateTime(1900, 3, 1))
                return ToMillisecondSerial(FloorToMilliseconds(value.Ticks - Early1900Epoch.Ticks));
            DateTime epoch = dateSystem == ExcelDateSystem.NineteenFour ? Epoch1904 : OleAutomationEpoch;
            return ToMillisecondSerial(FloorToMilliseconds(value.Ticks - epoch.Ticks));
        }

        /// <summary>
        /// Converts a date to an OLE Automation serial shifted to the date system, as pivot caches store dates. Dates
        /// before December 30, 1899 use the OLE Automation encoding, whose negative serials hold the time as a positive
        /// fraction. The time is floored to the millisecond, so it never moves to a later day.
        /// </summary>
        internal static double ToOleAutomationSerial(DateTime value, ExcelDateSystem dateSystem) {
            long milliseconds = value >= OleAutomationEpoch
                ? FloorToMilliseconds(value.Ticks - OleAutomationEpoch.Ticks)
                : OleAutomationMilliseconds(value);
            if (dateSystem == ExcelDateSystem.NineteenFour) milliseconds -= (long)Date1904OffsetDays * MillisecondsPerDay;
            return ToMillisecondSerial(milliseconds);
        }

        /// <summary>
        /// Converts an Excel serial value from the selected date system to a date.
        /// In the 1900 system, serial 60 denotes Excel's fictitious February 29 and
        /// maps to February 28 because <see cref="DateTime"/> cannot represent it.
        /// Negative serials extend the epoch backwards in both systems, December 31, 1899
        /// or January 1, 1904, with fractional values representing elapsed portions of a day.
        /// </summary>
        public static DateTime FromSerial(double serial, ExcelDateSystem dateSystem) {
            if (dateSystem == ExcelDateSystem.NineteenFour) return FromLinearSerial(Epoch1904, serial);
            if (serial < 60d) return FromLinearSerial(Early1900Epoch, serial);
            return FromOleAutomationSerial(serial);
        }

        /// <summary>
        /// Converts a pivot-cache serial, an OLE Automation date shifted to the date system. Unlike worksheet serials,
        /// these use the OLE Automation encoding before December 30, 1899.
        /// </summary>
        internal static DateTime FromOleAutomationSerial(double serial, ExcelDateSystem dateSystem) {
            if (dateSystem == ExcelDateSystem.NineteenHundred) return FromOleAutomationSerial(serial);
            // From the OLE Automation epoch on, a 1904 serial counts days from January 1, 1904.
            if (serial >= -Date1904OffsetDays) return FromLinearSerial(Epoch1904, serial);
            // Before it, adding the offset moves the serial toward zero, which is exact.
            return FromOleAutomationSerial(serial + Date1904OffsetDays);
        }

        /// <summary>
        /// Moves a serial by whole days, keeping its time to the millisecond. The time is resolved from the serial's
        /// own fraction before the days are added, so it never rounds into another day.
        /// </summary>
        internal static double ShiftSerialDays(double serial, long days) {
            double whole = Math.Floor(serial);
            // Serials too large for whole milliseconds to be exact are not dates; shift them as numbers.
            if (!(Math.Abs(whole) <= MaximumDateTimeDays * 2)) return serial + days;
            long milliseconds = ((long)whole + days) * MillisecondsPerDay + RoundFractionToMilliseconds(serial - whole);
            return ToMillisecondSerial(milliseconds);
        }

        /// <summary>
        /// Converts a time of day or a duration to its serial, truncated toward zero to whole milliseconds.
        /// A duration of 24 hours or more spans that many days; a negative duration has a negative serial.
        /// </summary>
        internal static double ToSerial(TimeSpan value) => ToMillisecondSerial(value.Ticks / TimeSpan.TicksPerMillisecond);

        /// <summary>
        /// Converts a serial formatted as a time or elapsed time. These count days from December 30, 1899 in both
        /// date systems, so the result minus that date is the stored duration, including a negative one.
        /// </summary>
        internal static DateTime FromElapsedSerial(double serial) => FromLinearSerial(OleAutomationEpoch, serial);

        /// <summary>
        /// Converts an OLE Automation date as <see cref="DateTime.FromOADate"/> does, without rounding the time into
        /// the next day.
        /// </summary>
        internal static DateTime FromOleAutomationSerial(double value) {
            if (value >= 0d) return FromLinearSerial(OleAutomationEpoch, value);
            // A negative OLE Automation date counts whole days back from the epoch and adds its fraction as a time.
            DateTime date = DateTime.FromOADate(value);
            DateTime day = OleAutomationEpoch.AddDays(Math.Truncate(value));
            return date >= day && date < day.AddDays(1) ? date : day.AddTicks(TimeSpan.TicksPerDay - TimeSpan.TicksPerMillisecond);
        }

        private static DateTime FromLinearSerial(DateTime epoch, double serial) {
            double wholeDays = Math.Floor(serial);
            // Also rejects NaN and infinities before the day count becomes an integer.
            if (!(wholeDays >= -MaximumDateTimeDays && wholeDays <= MaximumDateTimeDays)) {
                throw new ArgumentOutOfRangeException(nameof(serial), serial, "The serial is outside the range of DateTime.");
            }
            long milliseconds = RoundFractionToMilliseconds(serial - wholeDays);
            return epoch.AddTicks((long)wholeDays * TimeSpan.TicksPerDay + milliseconds * TimeSpan.TicksPerMillisecond);
        }

        /// <summary>
        /// Rounds a day fraction in [0, 1) to milliseconds, halves away from zero, and keeps the result within the day.
        /// The double product is correctly rounded, so it can only land on a half when the exact product is close to
        /// one; that case is decided exactly.
        /// </summary>
        private static long RoundFractionToMilliseconds(double fraction) {
            double product = fraction * MillisecondsPerDay;
            double floor = Math.Floor(product);
            long milliseconds = (long)floor;
            double rest = product - floor;
            if (rest > 0.5d || rest == 0.5d && !IsExactProductBelowHalf(fraction, milliseconds)) milliseconds++;
            return Math.Min(milliseconds, MillisecondsPerDay - 1);
        }

        /// <summary>Whether fraction × milliseconds per day is exactly less than <paramref name="milliseconds"/> + 0.5.</summary>
        private static bool IsExactProductBelowHalf(double fraction, long milliseconds) {
            long bits = BitConverter.DoubleToInt64Bits(fraction);
            int exponentBits = (int)((bits >> 52) & 0x7FF);
            long mantissa = bits & 0xF_FFFF_FFFF_FFFFL;
            if (exponentBits != 0) mantissa |= 1L << 52;
            int scale = 1075 - (exponentBits == 0 ? 1 : exponentBits);
            // fraction = mantissa / 2^scale; compare 2 * mantissa * msPerDay with (2 * ms + 1) * 2^scale.
            BigInteger product = new BigInteger(mantissa) * (2 * MillisecondsPerDay);
            BigInteger half = new BigInteger(2 * milliseconds + 1) << scale;
            return product < half;
        }

        /// <summary>
        /// The OLE Automation date of a date before the epoch in milliseconds, encoded as <see cref="DateTime.ToOADate"/>
        /// does but from the floored millisecond, so the time stays on its day.
        /// </summary>
        private static long OleAutomationMilliseconds(DateTime value) {
            long milliseconds = FloorToMilliseconds(value.Ticks - OleAutomationEpoch.Ticks);
            if (milliseconds < 0) {
                long fraction = milliseconds % MillisecondsPerDay;
                if (fraction != 0) milliseconds -= (MillisecondsPerDay + fraction) * 2;
            }
            return milliseconds;
        }

        private static long FloorToMilliseconds(long ticks) {
            long milliseconds = ticks / TimeSpan.TicksPerMillisecond;
            return ticks % TimeSpan.TicksPerMillisecond < 0 ? milliseconds - 1 : milliseconds;
        }

        private static double ToMillisecondSerial(long milliseconds) => milliseconds / (double)MillisecondsPerDay;

        /// <summary>
        /// Converts a date-formatted serial, or returns false when no <see cref="DateTime"/> represents it, the case
        /// Excel displays as <c>####</c>. A non-calendar (elapsed time) style converts without the date-system shift.
        /// </summary>
        internal static bool TryFromSerial(double serial, ExcelDateSystem dateSystem, bool calendarStyle, out DateTime value) =>
            OfficeIMO.Excel.LegacyXls.Projection.LegacyXlsDateSerialConverter.TryConvert(
                serial, dateSystem == ExcelDateSystem.NineteenFour, out value, calendarStyle);

        /// <summary>
        /// Returns the date for a date-formatted serial, or the serial number itself when it has no date. Untyped
        /// reads use it so one out-of-range cell does not stop a whole read or conversion.
        /// </summary>
        internal static object FromSerialOrNumber(double serial, ExcelDateSystem dateSystem, bool calendarStyle) =>
            TryFromSerial(serial, dateSystem, calendarStyle, out DateTime value) ? value : serial;

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
