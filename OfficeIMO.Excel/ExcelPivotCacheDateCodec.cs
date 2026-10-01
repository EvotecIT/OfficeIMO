namespace OfficeIMO.Excel {
    /// <summary>
    /// Converts pivot-cache date items. Excel caches use OLE Automation dates,
    /// including before March 1900, rather than the worksheet's early-1900 epoch.
    /// Keep this format boundary separate from Gregorian worksheet date conversion.
    /// </summary>
    internal static class ExcelPivotCacheDateCodec {
        internal static DateTime FromSerial(double serial, ExcelDateSystem system) => DateTime.FromOADate(
            system == ExcelDateSystem.NineteenFour ? serial + ExcelDateSystemConverter.Date1904OffsetDays : serial);
        internal static double ToSerial(DateTime date, ExcelDateSystem system) => date.ToOADate()
            - (system == ExcelDateSystem.NineteenFour ? ExcelDateSystemConverter.Date1904OffsetDays : 0d);
    }
}
