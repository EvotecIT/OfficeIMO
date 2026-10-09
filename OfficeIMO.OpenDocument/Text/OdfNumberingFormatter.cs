namespace OfficeIMO.OpenDocument;

/// <summary>Bounded ODF decimal, alphabetic and Roman numbering shared by list and field projection.</summary>
internal static class OdfNumberingFormatter {
    internal static string Format(long value, string format, bool sync) {
        if (format == "1") return value.ToString(CultureInfo.InvariantCulture);
        if (value <= 0) throw new NotSupportedException("Nonpositive counters require decimal numbering in this projection profile.");
        if (format is "a" or "A") {
            if (sync) {
                long count = (value - 1) / 26 + 1;
                if (count > 4096) throw new NotSupportedException("Numbering exceeds the label length limit.");
                return new string((char)((format == "a" ? 'a' : 'A') + (value - 1) % 26), (int)count);
            }
            string result = "";
            while (value > 0) { value--; result = (char)((format == "a" ? 'a' : 'A') + value % 26) + result; value /= 26; }
            return result;
        }
        if (format is "i" or "I") {
            if (value > 3999) throw new NotSupportedException("Roman numbering is supported through 3999.");
            var result = new StringBuilder();
            foreach (var pair in new[] { (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") })
                while (value >= pair.Item1) { result.Append(pair.Item2); value -= pair.Item1; }
            return format == "i" ? result.ToString().ToLowerInvariant() : result.ToString();
        }
        throw new NotSupportedException("Unsupported native numbering format.");
    }
}
