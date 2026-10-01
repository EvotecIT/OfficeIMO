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
    }
}
