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
                if (CssStyleMapper.TryParseRgbColor(v, out byte rr, out byte rg, out byte rb)) {
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
    }
}
