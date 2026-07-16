namespace OfficeIMO.Excel {
    /// <summary>
    /// Converts dates to and from Excel serial values using the workbook's configured date system.
    /// </summary>
    public static class ExcelDateSystemConverter {
        /// <summary>
        /// Number of days between the Excel 1900 and 1904 date-system epochs.
        /// </summary>
        public const double Date1904OffsetDays = 1462d;

        /// <summary>
        /// Converts a date to the serial value used by the selected Excel date system.
        /// </summary>
        public static double ToSerial(DateTime value, ExcelDateSystem dateSystem) {
            double serial = value.ToOADate();
            return dateSystem == ExcelDateSystem.NineteenFour ? serial - Date1904OffsetDays : serial;
        }

        /// <summary>
        /// Converts an Excel serial value from the selected date system to a date.
        /// </summary>
        public static DateTime FromSerial(double serial, ExcelDateSystem dateSystem) {
            double oa = dateSystem == ExcelDateSystem.NineteenFour ? serial + Date1904OffsetDays : serial;
            return DateTime.FromOADate(oa);
        }
    }
}
