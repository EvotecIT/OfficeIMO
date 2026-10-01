namespace OfficeIMO.Excel {
    internal static partial class ExcelNumberFormatDisplay {
        /// <summary>Resolves the supported red format color from the same value-selected section as numeric text.</summary>
        internal static string? GetNumericFormatColor(double value, uint numberFormatId, string? formatCode) {
            if (numberFormatId == 0U) return null;
            string? resolved = ResolveFormatCode(numberFormatId, formatCode);
            if (string.IsNullOrWhiteSpace(resolved)) return null;
            string section = SelectNumberFormatSection(resolved!, value < 0 ? 1 : value == 0 ? 2 : 0, value, out _);
            bool inQuote = false;
            for (int index = 0; index < section.Length; index++) {
                char ch = section[index];
                if (ch == '"') { inQuote = !inQuote; continue; }
                if (inQuote) continue;
                if (ch is '\\' or '_' or '*') { index++; continue; }
                if (ch != '[') continue;
                int close = section.IndexOf(']', index + 1);
                if (close < 0) return null;
                if (string.Equals(section.Substring(index + 1, close - index - 1), "Red", StringComparison.OrdinalIgnoreCase))
                    return "FF0000";
                index = close;
            }
            return null;
        }

    }
}
