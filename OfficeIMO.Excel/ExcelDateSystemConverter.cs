namespace OfficeIMO.Excel {
    internal static class ExcelDateSystemConverter {
        internal const double Date1904OffsetDays = 1462d;

        internal static double ToSerial(DateTime value, ExcelDateSystem dateSystem) {
            double serial = value.ToOADate();
            return dateSystem == ExcelDateSystem.NineteenFour ? serial - Date1904OffsetDays : serial;
        }

        internal static DateTime FromSerial(double serial, ExcelDateSystem dateSystem) {
            double oa = dateSystem == ExcelDateSystem.NineteenFour ? serial + Date1904OffsetDays : serial;
            return DateTime.FromOADate(oa);
        }
    }
}
