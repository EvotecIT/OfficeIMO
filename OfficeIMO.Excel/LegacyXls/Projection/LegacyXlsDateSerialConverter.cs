namespace OfficeIMO.Excel.LegacyXls.Projection {
    internal static class LegacyXlsDateSerialConverter {
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
