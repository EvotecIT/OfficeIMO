using OfficeIMO.Drawing;

namespace OfficeIMO.Xps;

internal sealed partial class XpsSvgConverter {
    private XElement? Glyphs(XElement e, Dictionary<string, Resource> scope, string part, int depth) {
        Charge(depth);
        CheckAttributes(e, "FontUri FontRenderingEmSize OriginX OriginY UnicodeString Indices Fill BidiLevel IsSideways StyleSimulations RenderTransform Clip Opacity FixedPage.NavigateUri CaretStops DeviceFontName");
        if ((string?)e.Attribute("IsSideways") == "true") { Loss("Sideways glyphs"); return null; }
        if (((string?)e.Attribute("StyleSimulations") ?? "None") != "None") { Loss("Simulated font style"); return null; }
        foreach (var child in e.Elements()) if (!new[] { "Glyphs.Fill", "Glyphs.Clip", "Glyphs.RenderTransform" }.Contains(child.Name.LocalName)) Loss(child.Name.LocalName);
        string uri = (string?)e.Attribute("FontUri") ?? throw new InvalidDataException("Missing glyph font URI.");
        string[] uriParts = uri.Split('#');
        if (uriParts.Length > 2) throw new InvalidDataException("Invalid font face URI.");
        string name = XpsPackage.Resolve(part, uriParts[0]);
        int? faceIndex = uriParts.Length == 2 ? ParseInt(uriParts[1]) : null;
        string fontKey = name + "#" + faceIndex;
        if (!_fonts.TryGetValue(fontKey, out var font)) {
            if (_fonts.Count >= 256) throw new InvalidDataException("XPS page font count exceeded.");
            string type = _page.Document.ContentType(name);
            byte[] bytes = _page.Document.GetPartBytes(name);
            if (type == "application/vnd.ms-package.obfuscated-opentype") XpsFontEncoding.Toggle(bytes, name);
            else if (type != "application/vnd.ms-opentype") throw new InvalidDataException("Invalid XPS font content type.");
            font = OfficeTrueTypeFont.TryLoad(bytes, faceIndex);
            if (font == null) { Loss("Unsupported embedded font program: " + name); return null; }
            _fonts.Add(fontKey, font);
        }
        double size = XpsPackage.Number((string?)e.Attribute("FontRenderingEmSize"));
        if (size == 0) return null;
        XpsPage.ValidateDimension(size);
        double x = XpsPackage.Number((string?)e.Attribute("OriginX")); double y = XpsPackage.Number((string?)e.Attribute("OriginY"));
        int bidi = ParseInt((string?)e.Attribute("BidiLevel") ?? "0");
        bool rtl = (bidi & 1) == 1;
        string text = XpsPage.Unescape((string?)e.Attribute("UnicodeString") ?? "");
        string indices = (string?)e.Attribute("Indices") ?? "";
        string[] entries = indices.Length == 0 ? Array.Empty<string>() : indices.Split(';');
        var data = new StringBuilder();
        int textIndex = 0, entryIndex = 0;
        while (entryIndex < entries.Length || textIndex < text.Length) {
            Charge(depth);
            int codeUnits = 1, glyphCount = 1;
            string entry = entryIndex < entries.Length ? entries[entryIndex] : "";
            if (entry.StartsWith("(", StringComparison.Ordinal)) {
                int end = entry.IndexOf(')');
                if (end < 0) throw new InvalidDataException("Invalid XPS glyph cluster.");
                string[] counts = entry.Substring(1, end - 1).Split(':');
                if (counts.Length > 2) throw new InvalidDataException("Invalid glyph cluster counts.");
                codeUnits = ParseInt(counts[0]); glyphCount = counts.Length == 2 ? ParseInt(counts[1]) : 1;
                if (codeUnits == 0 || glyphCount == 0 || glyphCount > entries.Length - entryIndex) throw new InvalidDataException("Invalid XPS glyph cluster length.");
                entry = entry.Substring(end + 1);
            } else if (textIndex < text.Length && char.IsHighSurrogate(text[textIndex]) && textIndex + 1 < text.Length && char.IsLowSurrogate(text[textIndex + 1])) codeUnits = 2;
            if (text.Length > 0 && codeUnits > text.Length - textIndex) throw new InvalidDataException("Glyph cluster exceeds UnicodeString.");
            for (int g = 0; g < glyphCount; g++) {
                string[] fields = (g == 0 ? entry : entries[entryIndex + g]).Split(',');
                if (fields.Length > 4 || fields[0].Contains("(")) throw new InvalidDataException("Malformed XPS glyph mapping.");
                int glyph;
                if (fields[0].Length > 0) glyph = ParseInt(fields[0]);
                else {
                    if (glyphCount != 1 || textIndex >= text.Length) throw new InvalidDataException("Glyph ID or Unicode mapping required.");
                    int scalar = char.ConvertToUtf32(text, textIndex);
                    _ = font.TryGetGlyphMetrics(scalar, out glyph, out _);
                }
                double nativeAdvance = font.FixedGlyphAdvance(glyph, size);
                double advance = fields.Length > 1 && fields[1].Length > 0 ? XpsPackage.Number(fields[1]) * size / 100 : nativeAdvance;
                if (advance < 0) throw new InvalidDataException("Negative glyph advance.");
                double u = fields.Length > 2 && fields[2].Length > 0 ? XpsPackage.Number(fields[2]) * size / 100 : 0;
                double v = fields.Length > 3 && fields[3].Length > 0 ? XpsPackage.Number(fields[3]) * size / 100 : 0;
                double gx = rtl ? x - nativeAdvance - u : x + u;
                var contours = font.FixedGlyphContours(glyph, size, gx, y - v, Math.Max(1, 1000000 - _points), _token);
                foreach (var contour in contours) {
                    _points = checked(_points + contour.Count);
                    if (_points > 1000000) throw new InvalidDataException("XPS glyph outline point limit exceeded.");
                    for (int i = 0; i < contour.Count; i++) data.Append(i == 0 ? "M" : "L").Append(N(contour[i].X)).Append(' ').Append(N(contour[i].Y)).Append(' ');
                    data.Append("Z ");
                }
                x += rtl ? -advance : advance;
            }
            textIndex += codeUnits;
            entryIndex += glyphCount;
        }
        _outputCharacters = checked(_outputCharacters + data.Length);
        if (_outputCharacters > 32 * 1024 * 1024) throw new InvalidDataException("XPS SVG output budget exceeded.");
        if (data.Length == 0) {
            var empty = new XElement(Svg + "g");
            if (text.Length > 0) empty.SetAttributeValue("aria-label", text);
            return empty;
        }
        var path = new XElement(Svg + "path", new XAttribute("d", data.ToString()), new XAttribute("fill-rule", "nonzero"));
        if (text.Length > 0) path.SetAttributeValue("aria-label", text);
        Paint(e, "Fill", path, "fill", scope, part, depth);
        return ApplyImageFill(path);
    }
    private static int ParseInt(string text) {
        if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int value) || value < 0 || value > 65535) throw new InvalidDataException("Invalid XPS integer.");
        return value;
    }
}
