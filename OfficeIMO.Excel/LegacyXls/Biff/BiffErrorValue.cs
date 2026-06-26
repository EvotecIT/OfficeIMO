namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffErrorValue {
        internal static string ToText(byte errorCode) {
            return errorCode switch {
                0x00 => "#NULL!",
                0x07 => "#DIV/0!",
                0x0f => "#VALUE!",
                0x17 => "#REF!",
                0x1d => "#NAME?",
                0x24 => "#NUM!",
                0x2a => "#N/A",
                0x2b => "#GETTING_DATA",
                _ => FormattableString.Invariant($"#ERR({errorCode})")
            };
        }
    }
}
