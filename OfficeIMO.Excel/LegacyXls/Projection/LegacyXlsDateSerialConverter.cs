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
