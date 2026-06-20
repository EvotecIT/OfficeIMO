namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private static void AddUnsupportedHeaderFooterFormatting(ExcelSheet.HeaderFooterSnapshot headerFooter, string sheetName, ref int count, List<string> details) {
            var sections = EnumerateHeaderFooterSections(headerFooter).ToArray();
            var stylesByLocation = new Dictionary<string, string?>(StringComparer.Ordinal);
            var hasTextByLocation = new Dictionary<string, bool>(StringComparer.Ordinal);
            foreach (var section in sections) {
                if (!TryGetHeaderFooterFormatting(section.Text, out string? styleSignature, out string reason)) {
                    continue;
                }

                hasTextByLocation[section.Location] = !string.IsNullOrWhiteSpace(section.Text);
                stylesByLocation[section.Location] = styleSignature;
                if (string.IsNullOrWhiteSpace(reason)) {
                    continue;
                }

                count++;
                details.Add($"{sheetName} {section.Location}: {reason} is simplified by the first-party PDF header/footer writer.");
            }

            AddMixedHeaderFooterStyle("header", new[] { "header left", "header center", "header right" }, stylesByLocation, hasTextByLocation, sheetName, ref count, details);
            AddMixedHeaderFooterStyle("footer", new[] { "footer left", "footer center", "footer right" }, stylesByLocation, hasTextByLocation, sheetName, ref count, details);
            AddMixedHeaderFooterStyle("first header", new[] { "first header left", "first header center", "first header right" }, stylesByLocation, hasTextByLocation, sheetName, ref count, details);
            AddMixedHeaderFooterStyle("first footer", new[] { "first footer left", "first footer center", "first footer right" }, stylesByLocation, hasTextByLocation, sheetName, ref count, details);
            AddMixedHeaderFooterStyle("even header", new[] { "even header left", "even header center", "even header right" }, stylesByLocation, hasTextByLocation, sheetName, ref count, details);
            AddMixedHeaderFooterStyle("even footer", new[] { "even footer left", "even footer center", "even footer right" }, stylesByLocation, hasTextByLocation, sheetName, ref count, details);
        }

        private static IEnumerable<(string Location, string Text)> EnumerateHeaderFooterSections(ExcelSheet.HeaderFooterSnapshot headerFooter) {
            yield return ("header left", headerFooter.HeaderLeft);
            yield return ("header center", headerFooter.HeaderCenter);
            yield return ("header right", headerFooter.HeaderRight);
            yield return ("footer left", headerFooter.FooterLeft);
            yield return ("footer center", headerFooter.FooterCenter);
            yield return ("footer right", headerFooter.FooterRight);
            yield return ("first header left", headerFooter.FirstHeaderLeft);
            yield return ("first header center", headerFooter.FirstHeaderCenter);
            yield return ("first header right", headerFooter.FirstHeaderRight);
            yield return ("first footer left", headerFooter.FirstFooterLeft);
            yield return ("first footer center", headerFooter.FirstFooterCenter);
            yield return ("first footer right", headerFooter.FirstFooterRight);
            yield return ("even header left", headerFooter.EvenHeaderLeft);
            yield return ("even header center", headerFooter.EvenHeaderCenter);
            yield return ("even header right", headerFooter.EvenHeaderRight);
            yield return ("even footer left", headerFooter.EvenFooterLeft);
            yield return ("even footer center", headerFooter.EvenFooterCenter);
            yield return ("even footer right", headerFooter.EvenFooterRight);
        }

        private static void AddMixedHeaderFooterStyle(
            string scope,
            IReadOnlyList<string> locations,
            IReadOnlyDictionary<string, string?> stylesByLocation,
            IReadOnlyDictionary<string, bool> hasTextByLocation,
            string sheetName,
            ref int count,
            List<string> details) {
            string? sharedStyle = null;
            bool hasStyle = false;
            bool hasUnstyledText = false;
            foreach (string location in locations) {
                if (!hasTextByLocation.TryGetValue(location, out bool hasText) || !hasText) {
                    continue;
                }

                string? style = stylesByLocation.TryGetValue(location, out string? value) ? value : null;
                if (string.IsNullOrWhiteSpace(style)) {
                    hasUnstyledText = true;
                    continue;
                }

                if (!hasStyle) {
                    sharedStyle = style;
                    hasStyle = true;
                    continue;
                }

                if (!string.Equals(sharedStyle, style, StringComparison.Ordinal)) {
                    count++;
                    details.Add($"{sheetName} {scope}: mixed header/footer formatting is simplified by the first-party PDF header/footer writer.");
                    return;
                }
            }

            if (hasStyle && hasUnstyledText) {
                count++;
                details.Add($"{sheetName} {scope}: mixed header/footer formatting is simplified by the first-party PDF header/footer writer.");
            }
        }

        private static bool TryGetHeaderFooterFormatting(string text, out string? styleSignature, out string reason) {
            styleSignature = null;
            reason = string.Empty;
            if (string.IsNullOrEmpty(text)) {
                return true;
            }

            bool hasVisibleContent = false;
            bool hasStyle = false;
            bool bold = false;
            bool italic = false;
            string? color = null;
            string? fontFamily = null;
            double? fontSize = null;
            for (int i = 0; i < text.Length; i++) {
                char current = text[i];
                if (current != '&') {
                    hasVisibleContent = true;
                    continue;
                }

                if (i + 1 >= text.Length) {
                    hasVisibleContent = true;
                    continue;
                }

                char token = text[++i];
                switch (char.ToUpperInvariant(token)) {
                    case '&':
                        hasVisibleContent = true;
                        break;
                    case 'P':
                    case 'N':
                    case 'D':
                    case 'T':
                    case 'A':
                    case 'F':
                    case 'Z':
                    case 'G':
                        hasVisibleContent = true;
                        break;
                    case 'U':
                        reason = "underline formatting";
                        return true;
                    case 'S':
                        reason = "strikethrough formatting";
                        return true;
                    case 'B':
                    case 'I':
                        if (hasVisibleContent) {
                            reason = token == 'B' ? "partial bold formatting" : "partial italic formatting";
                            return true;
                        }

                        hasStyle = true;
                        if (token == 'B') {
                            bold = !bold;
                        } else {
                            italic = !italic;
                        }
                        break;
                    case 'K':
                        if (!TryReadHeaderFooterColorToken(text, ref i, out color)) {
                            reason = "malformed color formatting";
                            return true;
                        }

                        if (hasVisibleContent) {
                            reason = "partial color formatting";
                            return true;
                        }

                        hasStyle = true;
                        break;
                    case '"':
                        if (!TryReadHeaderFooterFontToken(text, ref i, out string quotedToken)) {
                            reason = "malformed font formatting";
                            return true;
                        }

                        if (!TryMapSupportedPdfFontToken(quotedToken, out fontFamily, out bool tokenBold, out bool tokenItalic)) {
                            reason = "unsupported font formatting";
                            return true;
                        }

                        if (hasVisibleContent) {
                            reason = "partial font formatting";
                            return true;
                        }

                        hasStyle = true;
                        bold = tokenBold;
                        italic = tokenItalic;
                        break;
                    case 'L':
                    case 'C':
                    case 'R':
                        break;
                    default:
                        if (char.IsDigit(token)) {
                            if (!TryReadHeaderFooterFontSizeToken(text, token, ref i, out double parsedFontSize)) {
                                reason = "malformed font-size formatting";
                                return true;
                            }

                            if (hasVisibleContent) {
                                reason = "partial font-size formatting";
                                return true;
                            }

                            hasStyle = true;
                            fontSize = parsedFontSize;
                        } else {
                            hasVisibleContent = true;
                        }

                        break;
                }
            }

            if (hasStyle) {
                styleSignature = $"B={bold};I={italic};C={color ?? string.Empty};F={fontFamily ?? string.Empty};S={fontSize?.ToString(System.Globalization.CultureInfo.InvariantCulture) ?? string.Empty}";
            }

            return true;
        }

        private static bool TryReadHeaderFooterColorToken(string text, ref int index, out string color) {
            color = string.Empty;
            int start = index + 1;
            int length = 0;
            while (start + length < text.Length && length < 6 && IsHeaderFooterHexDigit(text[start + length])) {
                length++;
            }

            if (length != 6) {
                return false;
            }

            color = text.Substring(start, length).ToUpperInvariant();
            index = start + length - 1;
            return true;
        }

        private static bool TryReadHeaderFooterFontToken(string text, ref int index, out string token) {
            token = string.Empty;
            int closingQuote = text.IndexOf('"', index + 1);
            if (closingQuote < 0) {
                return false;
            }

            token = text.Substring(index + 1, closingQuote - index - 1);
            index = closingQuote;
            return true;
        }

        private static bool TryReadHeaderFooterFontSizeToken(string text, char firstDigit, ref int index, out double fontSize) {
            var builder = new System.Text.StringBuilder();
            builder.Append(firstDigit);
            while (index + 1 < text.Length && char.IsDigit(text[index + 1])) {
                index++;
                builder.Append(text[index]);
            }

            return double.TryParse(builder.ToString(), System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out fontSize)
                   && fontSize > 0D;
        }

        private static bool TryMapSupportedPdfFontToken(string token, out string fontFamily, out bool bold, out bool italic) {
            fontFamily = string.Empty;
            bold = false;
            italic = false;
            if (string.IsNullOrWhiteSpace(token)) {
                return false;
            }

            string[] parts = token.Split(new[] { ',' }, 2);
            fontFamily = NormalizePdfFontFamily(parts[0]);
            if (!IsSupportedPdfFontFamily(fontFamily)) {
                return false;
            }

            if (parts.Length > 1) {
                string fontStyle = parts[1];
                bold = fontStyle.IndexOf("bold", StringComparison.OrdinalIgnoreCase) >= 0;
                italic = fontStyle.IndexOf("italic", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         fontStyle.IndexOf("oblique", StringComparison.OrdinalIgnoreCase) >= 0;
            }

            return true;
        }

        private static string NormalizePdfFontFamily(string fontFamily) {
            var builder = new System.Text.StringBuilder(fontFamily.Length);
            foreach (char ch in fontFamily.Trim(' ', '\t', '"', '\'')) {
                if (char.IsLetterOrDigit(ch)) {
                    builder.Append(char.ToLowerInvariant(ch));
                }
            }

            return builder.ToString();
        }

        private static bool IsSupportedPdfFontFamily(string normalizedFamily) {
            switch (normalizedFamily) {
                case "timesnewroman":
                case "times":
                case "timesroman":
                case "georgia":
                case "cambria":
                case "serif":
                case "couriernew":
                case "courier":
                case "consolas":
                case "lucidaconsole":
                case "monospace":
                case "monaco":
                case "arial":
                case "helvetica":
                case "calibri":
                case "aptos":
                case "segoeui":
                case "tahoma":
                case "verdana":
                case "sans":
                case "sansserif":
                    return true;
                default:
                    return false;
            }
        }

        private static bool IsHeaderFooterHexDigit(char value) {
            return (value >= '0' && value <= '9')
                   || (value >= 'a' && value <= 'f')
                   || (value >= 'A' && value <= 'F');
        }
    }
}
