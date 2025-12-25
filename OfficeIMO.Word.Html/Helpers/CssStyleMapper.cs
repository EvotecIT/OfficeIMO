using DocumentFormat.OpenXml.Wordprocessing;
using SixLabors.ImageSharp.PixelFormats;
using System.Globalization;
using Color = SixLabors.ImageSharp.Color;

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
            internal bool Underline { get; set; }
            internal bool Strike { get; set; }
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

            if (properties.TryGetValue("padding", out string? padding)) {
                ApplyPaddingShorthand(padding, result);
            }
            if (properties.TryGetValue("padding-left", out string? pl) && TryParseLength(pl, out int pL)) result.PaddingLeft = pL;
            if (properties.TryGetValue("padding-right", out string? pr) && TryParseLength(pr, out int pR)) result.PaddingRight = pR;
            if (properties.TryGetValue("padding-top", out string? pt) && TryParseLength(pt, out int pT)) result.PaddingTop = pT;
            if (properties.TryGetValue("padding-bottom", out string? pb) && TryParseLength(pb, out int pB)) result.PaddingBottom = pB;

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
                    var color = new Color(new Rgb24(hr, hg, hb));
                    return color.ToHexColor();
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
