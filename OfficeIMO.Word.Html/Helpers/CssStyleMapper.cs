using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Word.Html {
    internal enum WhiteSpaceMode {
        Normal,
        Pre,
        PreWrap,
        NoWrap,
    }

    internal static class CssStyleMapper {
        internal class CssProperties {
            internal int? MarginLeft { get; set; }
            internal int? MarginRight { get; set; }
            internal int? MarginTop { get; set; }
            internal int? MarginBottom { get; set; }
            internal int? PaddingLeft { get; set; }
            internal int? PaddingRight { get; set; }
            internal int? PaddingTop { get; set; }
            internal int? PaddingBottom { get; set; }
            internal bool? Underline { get; set; }
            internal UnderlineValues? UnderlineStyle { get; set; }
            internal bool? Strike { get; set; }
            internal string? BackgroundColor { get; set; }
            internal int? LineHeight { get; set; }
            internal LineSpacingRuleValues? LineHeightRule { get; set; }
            internal WhiteSpaceMode? WhiteSpace { get; set; }
        }

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

        public static CssProperties ParseStyles(string? style, bool rightToLeft = false) {
            CssProperties result = new();
            if (string.IsNullOrWhiteSpace(style)) {
                return result;
            }

            Dictionary<string, string> properties = Parse(style);

            ApplyBoxPropertiesInDeclarationOrder(style!, rightToLeft, result);

            if (properties.TryGetValue("text-decoration", out string? deco)) {
                ApplyTextDecoration(deco, result);
            }
            if (properties.TryGetValue("text-decoration-line", out string? decoLine)) {
                ApplyTextDecorationLine(decoLine, result);
            }
            if (properties.TryGetValue("text-decoration-style", out string? decoStyle) &&
                TryMapTextDecorationStyle(decoStyle, out var underlineStyle)) {
                result.UnderlineStyle = underlineStyle;
            }

            if (properties.TryGetValue("background-color", out string? bg)) {
                result.BackgroundColor = NormalizeColor(bg);
            }

            if (properties.TryGetValue("line-height", out string? lh) && TryParseLineHeight(lh, out int line, out LineSpacingRuleValues rule)) {
                result.LineHeight = line;
                result.LineHeightRule = rule;
            }

            if (properties.TryGetValue("white-space", out string? ws)) {
                ws = ws.Trim().ToLowerInvariant();
                result.WhiteSpace = ws switch {
                    "normal" => WhiteSpaceMode.Normal,
                    "pre" => WhiteSpaceMode.Pre,
                    "pre-wrap" => WhiteSpaceMode.PreWrap,
                    "nowrap" => WhiteSpaceMode.NoWrap,
                    _ => null,
                };
            }

            return result;
        }

        private static void ApplyBoxPropertiesInDeclarationOrder(
            string style,
            bool rightToLeft,
            CssProperties result) {
            ApplyBoxPropertiesInDeclarationOrder(style, rightToLeft, result, important: false);
            ApplyBoxPropertiesInDeclarationOrder(style, rightToLeft, result, important: true);
        }

        private static void ApplyBoxPropertiesInDeclarationOrder(
            string style,
            bool rightToLeft,
            CssProperties result,
            bool important) {
            foreach (string part in style.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                if (!TryParseDeclaration(part, out string name, out string value, out bool declarationIsImportant) ||
                    declarationIsImportant != important) {
                    continue;
                }

                switch (name) {
                    case "margin":
                        ApplyMarginShorthand(value, result);
                        break;
                    case "margin-left":
                        if (TryParseLength(value, out int marginLeft)) result.MarginLeft = marginLeft;
                        break;
                    case "margin-right":
                        if (TryParseLength(value, out int marginRight)) result.MarginRight = marginRight;
                        break;
                    case "margin-top":
                        if (TryParseLength(value, out int marginTop)) result.MarginTop = marginTop;
                        break;
                    case "margin-bottom":
                        if (TryParseLength(value, out int marginBottom)) result.MarginBottom = marginBottom;
                        break;
                    case "padding":
                        ApplyPaddingShorthand(value, result);
                        break;
                    case "padding-left":
                        if (TryParseLength(value, out int paddingLeft)) result.PaddingLeft = paddingLeft;
                        break;
                    case "padding-right":
                        if (TryParseLength(value, out int paddingRight)) result.PaddingRight = paddingRight;
                        break;
                    case "padding-top":
                        if (TryParseLength(value, out int paddingTop)) result.PaddingTop = paddingTop;
                        break;
                    case "padding-bottom":
                        if (TryParseLength(value, out int paddingBottom)) result.PaddingBottom = paddingBottom;
                        break;
                    default:
                        if (name.StartsWith("margin-", StringComparison.Ordinal)) {
                            ApplySingleLogicalBoxProperty(name, value, "margin", rightToLeft, result);
                        } else if (name.StartsWith("padding-", StringComparison.Ordinal)) {
                            ApplySingleLogicalBoxProperty(name, value, "padding", rightToLeft, result);
                        }
                        break;
                }
            }
        }

        internal static bool TryParseDeclaration(
            string declaration,
            out string name,
            out string value,
            out bool important) {
            string[] pieces = declaration.Split(new[] { ':' }, 2);
            if (pieces.Length != 2) {
                name = string.Empty;
                value = string.Empty;
                important = false;
                return false;
            }

            name = pieces[0].Trim().ToLowerInvariant();
            value = pieces[1].Trim();
            important = TryRemoveImportantSuffix(ref value);
            return name.Length > 0 && value.Length > 0;
        }

        private static bool TryRemoveImportantSuffix(ref string value) {
            int end = value.Length;
            while (end > 0 && char.IsWhiteSpace(value[end - 1])) {
                end--;
            }

            const string importantKeyword = "important";
            int keywordStart = end - importantKeyword.Length;
            if (keywordStart < 0 ||
                !value.Substring(keywordStart, importantKeyword.Length)
                    .Equals(importantKeyword, StringComparison.OrdinalIgnoreCase)) {
                return false;
            }

            int bangIndex = keywordStart;
            while (bangIndex > 0 && char.IsWhiteSpace(value[bangIndex - 1])) {
                bangIndex--;
            }
            if (bangIndex == 0 || value[bangIndex - 1] != '!') {
                return false;
            }

            value = value.Substring(0, bangIndex - 1).TrimEnd();
            return true;
        }

        private static void ApplySingleLogicalBoxProperty(
            string name,
            string value,
            string prefix,
            bool rightToLeft,
            CssProperties result) {
            var property = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase) {
                [name] = value
            };
            ApplyLogicalBoxProperties(property, prefix, rightToLeft, result);
        }

        private static void ApplyLogicalBoxProperties(
            IReadOnlyDictionary<string, string> properties,
            string prefix,
            bool rightToLeft,
            CssProperties result) {
            int? inlineStart = null;
            int? inlineEnd = null;
            int? blockStart = null;
            int? blockEnd = null;

            if (properties.TryGetValue($"{prefix}-inline", out string? inline)) {
                ParseLogicalPair(inline, out inlineStart, out inlineEnd);
            }
            if (properties.TryGetValue($"{prefix}-block", out string? block)) {
                ParseLogicalPair(block, out blockStart, out blockEnd);
            }
            if (properties.TryGetValue($"{prefix}-inline-start", out string? inlineStartText) &&
                TryParseLength(inlineStartText, out int parsedInlineStart)) {
                inlineStart = parsedInlineStart;
            }
            if (properties.TryGetValue($"{prefix}-inline-end", out string? inlineEndText) &&
                TryParseLength(inlineEndText, out int parsedInlineEnd)) {
                inlineEnd = parsedInlineEnd;
            }
            if (properties.TryGetValue($"{prefix}-block-start", out string? blockStartText) &&
                TryParseLength(blockStartText, out int parsedBlockStart)) {
                blockStart = parsedBlockStart;
            }
            if (properties.TryGetValue($"{prefix}-block-end", out string? blockEndText) &&
                TryParseLength(blockEndText, out int parsedBlockEnd)) {
                blockEnd = parsedBlockEnd;
            }

            if (prefix.Equals("margin", StringComparison.OrdinalIgnoreCase)) {
                if (inlineStart.HasValue) {
                    if (rightToLeft) result.MarginRight = inlineStart;
                    else result.MarginLeft = inlineStart;
                }
                if (inlineEnd.HasValue) {
                    if (rightToLeft) result.MarginLeft = inlineEnd;
                    else result.MarginRight = inlineEnd;
                }
                if (blockStart.HasValue) result.MarginTop = blockStart;
                if (blockEnd.HasValue) result.MarginBottom = blockEnd;
            } else {
                if (inlineStart.HasValue) {
                    if (rightToLeft) result.PaddingRight = inlineStart;
                    else result.PaddingLeft = inlineStart;
                }
                if (inlineEnd.HasValue) {
                    if (rightToLeft) result.PaddingLeft = inlineEnd;
                    else result.PaddingRight = inlineEnd;
                }
                if (blockStart.HasValue) result.PaddingTop = blockStart;
                if (blockEnd.HasValue) result.PaddingBottom = blockEnd;
            }
        }

        private static void ParseLogicalPair(string value, out int? start, out int? end) {
            start = null;
            end = null;
            var parts = value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0 || parts.Length > 2 || !TryParseLength(parts[0], out int first)) {
                return;
            }

            if (parts.Length == 1) {
                start = first;
                end = first;
                return;
            }

            if (!TryParseLength(parts[1], out int second)) {
                return;
            }

            start = first;
            end = second;
        }

        private static Dictionary<string, string> Parse(string? style) {
            Dictionary<string, string> dict = new(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(style)) {
                return dict;
            }

            string styleText = style ?? string.Empty;
            for (int priorityPass = 0; priorityPass < 2; priorityPass++) {
                bool important = priorityPass == 1;
                foreach (string part in styleText.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                    if (TryParseDeclaration(part, out string name, out string value, out bool declarationIsImportant) &&
                        declarationIsImportant == important) {
                        dict[name] = value;
                    }
                }
            }

            return dict;
        }

        private static void ApplyTextDecoration(string value, CssProperties result) {
            foreach (var token in SplitCssTokens(value)) {
                if (ApplyTextDecorationLineToken(token, result)) {
                    continue;
                }
                if (TryMapTextDecorationStyle(token, out var underlineStyle)) {
                    result.UnderlineStyle = underlineStyle;
                }
            }
        }

        private static void ApplyTextDecorationLine(string value, CssProperties result) {
            foreach (var token in SplitCssTokens(value)) {
                ApplyTextDecorationLineToken(token, result);
            }
        }

        private static bool ApplyTextDecorationLineToken(string token, CssProperties result) {
            switch (token.Trim().ToLowerInvariant()) {
                case "none":
                    result.Underline = false;
                    result.Strike = false;
                    result.UnderlineStyle = null;
                    return true;
                case "underline":
                    result.Underline = true;
                    result.UnderlineStyle ??= UnderlineValues.Single;
                    return true;
                case "line-through":
                    result.Strike = true;
                    return true;
                default:
                    return false;
            }
        }

        private static bool TryMapTextDecorationStyle(string value, out UnderlineValues underlineStyle) {
            underlineStyle = UnderlineValues.Single;
            switch (value.Trim().ToLowerInvariant()) {
                case "solid":
                    underlineStyle = UnderlineValues.Single;
                    return true;
                case "double":
                    underlineStyle = UnderlineValues.Double;
                    return true;
                case "dotted":
                    underlineStyle = UnderlineValues.Dotted;
                    return true;
                case "dashed":
                    underlineStyle = UnderlineValues.Dash;
                    return true;
                case "wavy":
                    underlineStyle = UnderlineValues.Wave;
                    return true;
                default:
                    return false;
            }
        }

        private static IEnumerable<string> SplitCssTokens(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                yield break;
            }

            foreach (var token in value!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)) {
                yield return token;
            }
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

        private static bool TryParseLength(string value, out int twips) {
            twips = 0;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }
            value = value.Trim().ToLowerInvariant();
            if (value.EndsWith("pt") && double.TryParse(value.Substring(0, value.Length - 2), NumberStyles.Number, CultureInfo.InvariantCulture, out double pt)) {
                twips = (int)Math.Round(pt * 20);
                return true;
            }
            if (value.EndsWith("px") && double.TryParse(value.Substring(0, value.Length - 2), NumberStyles.Number, CultureInfo.InvariantCulture, out double px)) {
                twips = (int)Math.Round(px * 15);
                return true;
            }
            if (value.EndsWith("em") && double.TryParse(value.Substring(0, value.Length - 2), NumberStyles.Number, CultureInfo.InvariantCulture, out double em)) {
                twips = (int)Math.Round(em * 16 * 15);
                return true;
            }
            if (double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out double number)) {
                twips = (int)Math.Round(number * 15);
                return true;
            }
            return false;
        }

        private static void ApplyMarginShorthand(string margin, CssProperties result) {
            var parts = margin.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) {
                return;
            }
            int? top = null, right = null, bottom = null, left = null;
            if (parts.Length == 1) {
                if (TryParseLength(parts[0], out int all)) {
                    top = right = bottom = left = all;
                }
            } else if (parts.Length == 2) {
                if (TryParseLength(parts[0], out int tb) && TryParseLength(parts[1], out int lr)) {
                    top = bottom = tb;
                    left = right = lr;
                }
            } else if (parts.Length == 3) {
                if (TryParseLength(parts[0], out int t) && TryParseLength(parts[1], out int rl) && TryParseLength(parts[2], out int b)) {
                    top = t;
                    bottom = b;
                    left = right = rl;
                }
            } else {
                if (TryParseLength(parts[0], out int t) && TryParseLength(parts[1], out int r) && TryParseLength(parts[2], out int b) && TryParseLength(parts[3], out int l)) {
                    top = t;
                    right = r;
                    bottom = b;
                    left = l;
                }
            }
            if (top.HasValue) result.MarginTop = top;
            if (right.HasValue) result.MarginRight = right;
            if (bottom.HasValue) result.MarginBottom = bottom;
            if (left.HasValue) result.MarginLeft = left;
        }

        private static void ApplyPaddingShorthand(string padding, CssProperties result) {
            var parts = padding.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) {
                return;
            }
            int? top = null, right = null, bottom = null, left = null;
            if (parts.Length == 1) {
                if (TryParseLength(parts[0], out int all)) {
                    top = right = bottom = left = all;
                }
            } else if (parts.Length == 2) {
                if (TryParseLength(parts[0], out int tb) && TryParseLength(parts[1], out int lr)) {
                    top = bottom = tb;
                    left = right = lr;
                }
            } else if (parts.Length == 3) {
                if (TryParseLength(parts[0], out int t) && TryParseLength(parts[1], out int rl) && TryParseLength(parts[2], out int b)) {
                    top = t;
                    bottom = b;
                    left = right = rl;
                }
            } else {
                if (TryParseLength(parts[0], out int t) && TryParseLength(parts[1], out int r) && TryParseLength(parts[2], out int b) && TryParseLength(parts[3], out int l)) {
                    top = t;
                    right = r;
                    bottom = b;
                    left = l;
                }
            }
            if (top.HasValue) result.PaddingTop = top;
            if (right.HasValue) result.PaddingRight = right;
            if (bottom.HasValue) result.PaddingBottom = bottom;
            if (left.HasValue) result.PaddingLeft = left;
        }

        private static bool TryParseLineHeight(string value, out int twips, out LineSpacingRuleValues rule) {
            twips = 0;
            rule = LineSpacingRuleValues.Auto;
            if (string.IsNullOrWhiteSpace(value)) {
                return false;
            }
            value = value.Trim().ToLowerInvariant();
            if (value.EndsWith("pt") && double.TryParse(value.Substring(0, value.Length - 2), NumberStyles.Number, CultureInfo.InvariantCulture, out double pt)) {
                twips = (int)Math.Round(pt * 20);
                rule = LineSpacingRuleValues.Exact;
                return true;
            }
            if (value.EndsWith("px") && double.TryParse(value.Substring(0, value.Length - 2), NumberStyles.Number, CultureInfo.InvariantCulture, out double px)) {
                twips = (int)Math.Round(px * 15);
                rule = LineSpacingRuleValues.Exact;
                return true;
            }
            if (value.EndsWith("%") && double.TryParse(value.Substring(0, value.Length - 1), NumberStyles.Number, CultureInfo.InvariantCulture, out double percent)) {
                twips = (int)Math.Round(percent / 100d * 240d);
                rule = LineSpacingRuleValues.Auto;
                return true;
            }
            if (double.TryParse(value, NumberStyles.Number, CultureInfo.InvariantCulture, out double multiple)) {
                twips = (int)Math.Round(multiple * 240d);
                rule = LineSpacingRuleValues.Auto;
                return true;
            }
            return false;
        }

        private static string? NormalizeColor(string value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }
            value = value.Trim();
            if (value.StartsWith("hsl", StringComparison.OrdinalIgnoreCase)) {
                if (TryParseHsl(value, out byte hr, out byte hg, out byte hb)) {
                    var color = Color.FromRgb(hr, hg, hb);
                    return color.ToRgbHex();
                }
                return null;
            }
            if (value.StartsWith("rgb", StringComparison.OrdinalIgnoreCase)) {
                int start = value.IndexOf('(');
                int end = value.IndexOf(')');
                if (start >= 0 && end > start) {
                    var parts = value.Substring(start + 1, end - start - 1).Split(',');
                    if (parts.Length >= 3 &&
                        byte.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out byte r) &&
                        byte.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out byte g) &&
                        byte.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out byte b)) {
                        var color = Color.FromRgb(r, g, b);
                        return color.ToRgbHex();
                    }
                }
                return null;
            }
            try {
                var parsed = Color.Parse(value);
                return parsed.ToRgbHex();
            } catch {
                if (!value.StartsWith("#", StringComparison.Ordinal)) {
                    try {
                        var parsed = Color.Parse("#" + value);
                        return parsed.ToRgbHex();
                    } catch {
                        return null;
                    }
                }
                return null;
            }
        }

        private static bool TryParseHsl(string text, out byte r, out byte g, out byte b) {
            r = g = b = 0;
            int start = text.IndexOf('(');
            int end = text.LastIndexOf(')');
            if (start < 0 || end <= start) {
                return false;
            }
            var content = text.Substring(start + 1, end - start - 1);
            var slashIndex = content.IndexOf('/');
            if (slashIndex >= 0) {
                content = content.Substring(0, slashIndex);
            }
            var parts = content.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 3) {
                return false;
            }
            if (!double.TryParse(parts[0].Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out var h)) {
                return false;
            }
            if (!TryParsePercent(parts[1], out var s) || !TryParsePercent(parts[2], out var l)) {
                return false;
            }
            return HslToRgb(h, s, l, out r, out g, out b);
        }

        private static bool TryParsePercent(string text, out double value) {
            value = 0;
            var t = text.Trim();
            if (t.EndsWith("%", StringComparison.Ordinal)) {
                t = t.Substring(0, t.Length - 1);
            }
            if (!double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)) {
                return false;
            }
            value = parsed / 100d;
            return true;
        }

        private static bool HslToRgb(double h, double s, double l, out byte r, out byte g, out byte b) {
            r = g = b = 0;
            h = h % 360;
            if (h < 0) h += 360;
            s = s < 0 ? 0 : s > 1 ? 1 : s;
            l = l < 0 ? 0 : l > 1 ? 1 : l;

            double c = (1 - Math.Abs(2 * l - 1)) * s;
            double x = c * (1 - Math.Abs((h / 60d) % 2 - 1));
            double m = l - c / 2;

            double r1, g1, b1;
            if (h < 60) {
                r1 = c; g1 = x; b1 = 0;
            } else if (h < 120) {
                r1 = x; g1 = c; b1 = 0;
            } else if (h < 180) {
                r1 = 0; g1 = c; b1 = x;
            } else if (h < 240) {
                r1 = 0; g1 = x; b1 = c;
            } else if (h < 300) {
                r1 = x; g1 = 0; b1 = c;
            } else {
                r1 = c; g1 = 0; b1 = x;
            }

            r = (byte)Math.Round((r1 + m) * 255);
            g = (byte)Math.Round((g1 + m) * 255);
            b = (byte)Math.Round((b1 + m) * 255);
            return true;
        }
    }
}
