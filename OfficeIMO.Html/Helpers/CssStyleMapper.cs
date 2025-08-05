using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using OfficeIMO.Word;

namespace OfficeIMO.Html.Helpers {
    internal static class CssStyleMapper {
        public static WordParagraphStyles? MapParagraphStyle(string? style) {
            if (string.IsNullOrWhiteSpace(style)) {
                return null;
            }

            Dictionary<string, string> properties = Parse(style);
            if (properties.TryGetValue("font-weight", out string? weight) && weight.Equals("bold", StringComparison.OrdinalIgnoreCase)) {
                if (properties.TryGetValue("font-size", out string? sizeValue) && TryParseFontSize(sizeValue, out double size)) {
                    if (size >= 32) {
                        return WordParagraphStyles.Heading1;
                    }
                    if (size >= 24) {
                        return WordParagraphStyles.Heading2;
                    }
                    if (size >= 18) {
                        return WordParagraphStyles.Heading3;
                    }
                    if (size >= 16) {
                        return WordParagraphStyles.Heading4;
                    }
                    if (size >= 13) {
                        return WordParagraphStyles.Heading5;
                    }
                    if (size >= 12) {
                        return WordParagraphStyles.Heading6;
                    }
                }
            }

            return null;
        }

        private static Dictionary<string, string> Parse(string style) {
            Dictionary<string, string> dict = new(StringComparer.OrdinalIgnoreCase);
            foreach (string part in style.Split(';', StringSplitOptions.RemoveEmptyEntries)) {
                string[] pieces = part.Split(':', 2);
                if (pieces.Length == 2) {
                    dict[pieces[0].Trim()] = pieces[1].Trim();
                }
            }
            return dict;
        }

        private static bool TryParseFontSize(string value, out double size) {
            size = 0;
            value = value.Trim().ToLowerInvariant();

            string number = new(value.Where(c => char.IsDigit(c) || c == '.').ToArray());
            if (!double.TryParse(number, NumberStyles.Number, CultureInfo.InvariantCulture, out size)) {
                return false;
            }

            if (value.EndsWith("em", StringComparison.Ordinal)) {
                size *= 16; // approximate conversion
            }

            return size > 0;
        }
    }
}
