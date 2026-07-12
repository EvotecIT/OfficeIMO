using System.Globalization;
using Color = OfficeIMO.Drawing.OfficeColor;

namespace OfficeIMO.Word.Html {
    internal partial class HtmlToWordConverter {
        private static string? NormalizeColor(string? value) {
            if (string.IsNullOrWhiteSpace(value)) {
                return null;
            }
            string v = value!.Trim();
            if (v.StartsWith("hsl", StringComparison.OrdinalIgnoreCase)) {
                if (TryParseHsl(v, out byte hr, out byte hg, out byte hb)) {
                    var color = Color.FromRgb(hr, hg, hb);
                    return color.ToRgbHex();
                }
                return null;
            }
            if (v.StartsWith("rgb", StringComparison.OrdinalIgnoreCase)) {
                if (TryParseRgb(v, out byte rr, out byte rg, out byte rb)) {
                    var color = Color.FromRgb(rr, rg, rb);
                    return color.ToRgbHex();
                }
                return null;
            }
            try {
                var parsed = Color.Parse(v);
                return parsed.ToRgbHex();
            } catch {
                if (!v.StartsWith("#", StringComparison.Ordinal)) {
                    try {
                        var parsed = Color.Parse("#" + v);
                        return parsed.ToRgbHex();
                    } catch {
                        return null;
                    }
                }
                return null;
            }
        }

        private static bool TryParseRgb(string text, out byte r, out byte g, out byte b) {
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

            string[] parts = content.IndexOf(',') >= 0
                ? content.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                : content.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length < 3) {
                return false;
            }

            return TryParseRgbChannel(parts[0], out r) &&
                   TryParseRgbChannel(parts[1], out g) &&
                   TryParseRgbChannel(parts[2], out b);
        }

        private static bool TryParseRgbChannel(string text, out byte value) {
            value = 0;
            var t = text.Trim();
            double parsed;
            if (t.EndsWith("%", StringComparison.Ordinal)) {
                t = t.Substring(0, t.Length - 1);
                if (!double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed)) {
                    return false;
                }

                parsed = parsed * 255d / 100d;
            } else if (!double.TryParse(t, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed)) {
                return false;
            }

            parsed = parsed < 0 ? 0 : parsed > 255 ? 255 : parsed;
            value = (byte)Math.Round(parsed);
            return true;
        }
    }
}
