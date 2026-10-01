namespace OfficeIMO.Excel;

/// <summary>Identifies whether a numeric cell uses a calendar epoch or an unshifted time carrier.</summary>
internal enum ExcelSerialDateStyle : byte {
    None,
    Calendar,
    Time
}
