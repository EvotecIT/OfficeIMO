using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;
using SixLabors.ImageSharp.PixelFormats;

namespace OfficeIMO.Word.Html.Helpers {
    internal enum WhiteSpaceMode {
        Normal,
        Pre,
        PreWrap,
        NoWrap,
    }

    internal static class CssStyleMapper {
        internal class CssProperties {
            public int? MarginLeft;
            public int? MarginRight;
            public int? MarginTop;
            public int? MarginBottom;
            public bool Underline;
            public bool Strike;
            public string? BackgroundColor;
            public int? LineHeight;
            public LineSpacingRuleValues? LineHeightRule;
            public WhiteSpaceMode? WhiteSpace;
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

        public static CssProperties ParseStyles(string? style) {
            CssProperties result = new();
            if (string.IsNullOrWhiteSpace(style)) {
                return result;
            }

            Dictionary<string, string> properties = Parse(style);

            if (properties.TryGetValue("margin", out string? margin)) {
                ApplyMarginShorthand(margin, result);
            }
            if (properties.TryGetValue("margin-left", out string? ml) && TryParseLength(ml, out int mL)) result.MarginLeft = mL;
            if (properties.TryGetValue("margin-right", out string? mr) && TryParseLength(mr, out int mR)) result.MarginRight = mR;
            if (properties.TryGetValue("margin-top", out string? mt) && TryParseLength(mt, out int mT)) result.MarginTop = mT;
            if (properties.TryGetValue("margin-bottom", out string? mb) && TryParseLength(mb, out int mB)) result.MarginBottom = mB;

            if (properties.TryGetValue("text-decoration", out string? deco)) {
                foreach (var part in deco.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                    switch (part.Trim().ToLowerInvariant()) {
                        case "underline":
                            result.Underline = true;
                            break;
                        case "line-through":
                            result.Strike = true;
                            break;
                    }
                }
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

        private static Dictionary<string, string> Parse(string? style) {
            Dictionary<string, string> dict = new(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(style)) {
                return dict;
            }
            var styleText = style ?? string.Empty;
            foreach (string part in styleText.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)) {
                string[] pieces = part.Split(new[] { ':' }, 2);
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
            if (value.StartsWith("rgb", StringComparison.OrdinalIgnoreCase)) {
                int start = value.IndexOf('(');
                int end = value.IndexOf(')');
                if (start >= 0 && end > start) {
                    var parts = value.Substring(start + 1, end - start - 1).Split(',');
                    if (parts.Length >= 3 &&
                        byte.TryParse(parts[0], NumberStyles.Integer, CultureInfo.InvariantCulture, out byte r) &&
                        byte.TryParse(parts[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out byte g) &&
                        byte.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out byte b)) {
                        var color = new Color(new Rgb24(r, g, b));
                        return color.ToHexColor();
                    }
                }
                return null;
            }
            try {
                var parsed = Color.Parse(value);
                return parsed.ToHexColor();
            } catch {
                if (!value.StartsWith("#", StringComparison.Ordinal)) {
                    try {
                        var parsed = Color.Parse("#" + value);
                        return parsed.ToHexColor();
                    } catch {
                        return null;
                    }
                }
                return null;
            }
        }
    }
}
