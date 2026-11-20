namespace OfficeIMO.Excel {
    /// <summary>
    /// Converts dates to and from Excel serial values using the workbook's configured date system.
    /// </summary>
    /// <remarks>
    /// A serial counts days, and its whole part is the day Excel reports for it. Serials hold whole milliseconds:
    /// values are stored truncated to the millisecond, so a time just before midnight stays on its own day, and
    /// serials read back rounded to the nearest millisecond, never past the last millisecond of the serial's day.
    /// </remarks>
    public static class ExcelDateSystemConverter {
        private static readonly DateTime Early1900Epoch = new DateTime(1899, 12, 31);
        private static readonly DateTime OleAutomationEpoch = new DateTime(1899, 12, 30);
        private const long MillisecondsPerDay = 86_400_000L;
        private const double MaximumDateTimeDays = 3_652_059d;
        /// <summary>
        /// Number of days between the Excel 1900 and 1904 date-system epochs.
        /// </summary>
        public const double Date1904OffsetDays = 1462d;

        /// <summary>
        /// Converts a date to the serial value used by the selected Excel date system.
        /// </summary>
        public static double ToSerial(DateTime value, ExcelDateSystem dateSystem) {
            if (dateSystem == ExcelDateSystem.NineteenHundred && value < new DateTime(1900, 3, 1))
                return ToMillisecondSerial(FloorToMilliseconds(value.Ticks - Early1900Epoch.Ticks));
            // ToOADate truncates to the millisecond, which keeps the time on the value's day.
            double serial = value.ToOADate();
            return dateSystem == ExcelDateSystem.NineteenFour ? ShiftSerialDays(serial, -Date1904OffsetDays) : serial;
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
                return FromLinearSerial(Early1900Epoch, serial);
            return FromOleAutomationSerial(dateSystem == ExcelDateSystem.NineteenFour ? ShiftSerialDays(serial, Date1904OffsetDays) : serial);
        }

        /// <summary>
        /// Moves a serial by whole days without changing its time. Adding the days directly can round a fraction just
        /// below one up into the next day, because the larger result has fewer fractional bits; the result then keeps
        /// the largest serial still on the shifted day.
        /// </summary>
        internal static double ShiftSerialDays(double serial, double days) {
            if (double.IsNaN(serial) || double.IsInfinity(serial)) return serial + days;
            double whole = Math.Floor(serial);
            double shiftedWhole = whole + days;
            double shifted = shiftedWhole + (serial - whole);
            if (shifted < shiftedWhole + 1d) return shifted;
            double nextDay = shiftedWhole + 1d;
            if (nextDay == 0d) return -double.Epsilon;
            return BitConverter.Int64BitsToDouble(BitConverter.DoubleToInt64Bits(nextDay) + (nextDay > 0d ? -1L : 1L));
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
            long milliseconds = (long)Math.Round((serial - wholeDays) * MillisecondsPerDay, MidpointRounding.AwayFromZero);
            if (milliseconds == MillisecondsPerDay) milliseconds--;
            return epoch.AddTicks((long)wholeDays * TimeSpan.TicksPerDay + milliseconds * TimeSpan.TicksPerMillisecond);
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
