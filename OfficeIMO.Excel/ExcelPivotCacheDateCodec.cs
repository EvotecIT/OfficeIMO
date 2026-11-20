namespace OfficeIMO.Excel {
    /// <summary>
    /// Converts pivot-cache date items. Excel caches use OLE Automation dates,
    /// including before March 1900, rather than the worksheet's early-1900 epoch.
    /// Keep this format boundary separate from Gregorian worksheet date conversion.
    /// </summary>
    internal static class ExcelPivotCacheDateCodec {
        internal static DateTime FromSerial(double serial, ExcelDateSystem system) =>
            ExcelDateSystemConverter.FromOleAutomationSerial(serial, system);
        // The 1900 converter differs only before March 1900, which pivot caches encode as OLE Automation dates.
        internal static double ToSerial(DateTime date, ExcelDateSystem system) => system == ExcelDateSystem.NineteenFour
            ? ExcelDateSystemConverter.ToSerial(date, system)
            : date.ToOADate();
    }
}
