namespace OfficeIMO.Excel.LegacyXls.Projection {
    internal static class LegacyXlsDateSerialConverter {
        internal static bool TryConvert(double serial, bool uses1904DateSystem, out DateTime value) {
            try {
                value = uses1904DateSystem
                    ? new DateTime(1904, 1, 1).AddDays(serial)
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
