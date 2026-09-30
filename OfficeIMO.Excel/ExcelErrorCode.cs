namespace OfficeIMO.Excel {
    /// <summary>
    /// Maps Excel error text to the error ordinals shared by native workbook codecs and pivot ranking.
    /// </summary>
    internal static class ExcelErrorCode {
        private static readonly Dictionary<string, byte> Codes = new(StringComparer.OrdinalIgnoreCase) {
            ["#NULL!"] = 0x00,
            ["#DIV/0!"] = 0x07,
            ["#VALUE!"] = 0x0f,
            ["#REF!"] = 0x17,
            ["#NAME?"] = 0x1d,
            ["#NUM!"] = 0x24,
            ["#N/A"] = 0x2a,
            ["#GETTING_DATA"] = 0x2b
        };

        internal static bool TryGetCode(string text, out byte errorCode) {
            return Codes.TryGetValue((text ?? string.Empty).Trim(), out errorCode);
        }
    }
}
