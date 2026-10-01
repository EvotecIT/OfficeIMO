namespace OfficeIMO.Excel {
    internal static partial class ExcelNumberFormatDisplay {
        private static DecimalPlaceInfo GetDecimalPlaceInfo(string formatCode) {
            int dot = -1;
            for (int index = 0; index < formatCode.Length; index++) {
                if (formatCode[index] == LiteralPunctuationMarker && index + 1 < formatCode.Length) { index++; continue; }
                if (formatCode[index] == '.') { dot = index; break; }
            }
            if (dot < 0) return new DecimalPlaceInfo(0, 0);
            int required = 0;
            int maximum = 0;
            for (int index = dot + 1; index < formatCode.Length; index++) {
                char token = formatCode[index];
                if (token == '0') { maximum++; required = maximum; continue; }
                if (token is '#' or '?') { maximum++; continue; }
                break;
            }
            return new DecimalPlaceInfo(required, maximum);
        }

        private static string TrimOptionalDecimalPlaces(string text, int requiredDecimalPlaces) {
            int dot = text.IndexOf('.');
            if (dot < 0) return text;
            int end = text.Length - 1;
            while (end > dot + requiredDecimalPlaces && text[end] == '0') end--;
            if (end == dot) return text.Substring(0, dot);
            return end == text.Length - 1 ? text : text.Substring(0, end + 1);
        }

        private readonly struct DecimalPlaceInfo {
            internal DecimalPlaceInfo(int required, int maximum) {
                Required = required;
                Maximum = maximum;
            }
            internal int Required { get; }
            internal int Maximum { get; }
            internal int Optional => Maximum - Required;
        }
    }
}
