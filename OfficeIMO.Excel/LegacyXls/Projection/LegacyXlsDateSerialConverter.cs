namespace OfficeIMO.Excel.LegacyXls.Projection {
    internal static class LegacyXlsDateSerialConverter {
        // Serials in [60, 2950000) convert in both date systems and in both styles:
        // the 1900 system and time style pass them to FromOADate directly, and the
        // 1904 offset keeps them below its exclusive 2958466 limit.
        private const double AlwaysConvertibleMinimum = 60d;
        private const double AlwaysConvertibleLimit = 2_950_000d;

        /// <summary>Whether <see cref="TryConvert"/> succeeds, without building the date for common serials.</summary>
        internal static bool CanConvert(double serial, bool uses1904DateSystem, bool calendarStyle = true) =>
            serial >= AlwaysConvertibleMinimum && serial < AlwaysConvertibleLimit
            || TryConvert(serial, uses1904DateSystem, out _, calendarStyle);

        internal static bool TryConvert(double serial, bool uses1904DateSystem, out DateTime value, bool calendarStyle = true) {
            // A whole-day serial in the common range is the OLE Automation epoch plus that many days,
            // which is exactly what FromOADate returns for it in either date system and style.
            if (serial >= AlwaysConvertibleMinimum && serial < AlwaysConvertibleLimit) {
                double days = calendarStyle && uses1904DateSystem
                    ? serial + ExcelDateSystemConverter.Date1904OffsetDays
                    : serial;
                long wholeDays = (long)days;
                if (wholeDays == days) {
                    value = new DateTime(OleAutomationEpochTicks + wholeDays * TimeSpan.TicksPerDay);
                    return true;
                }
            }
            return TryConvertExactly(serial, uses1904DateSystem, out value, calendarStyle);
        }

        // Ticks of December 30, 1899, OLE Automation date 0.
        private const long OleAutomationEpochTicks = 599_264_352_000_000_000L;

        private static bool TryConvertExactly(double serial, bool uses1904DateSystem, out DateTime value, bool calendarStyle) {
            try {
                value = calendarStyle
                    ? ExcelDateSystemConverter.FromSerial(serial, uses1904DateSystem ? ExcelDateSystem.NineteenFour : ExcelDateSystem.NineteenHundred)
                    : DateTime.FromOADate(serial);
                return true;
            } catch (ArgumentException) {
                value = default;
                return false;
            } catch (OverflowException) {
                value = default;
                return false;
            }
        }
    }
}
