using OfficeIMO.Drawing;

namespace OfficeIMO.Xps;

internal sealed partial class XpsSvgConverter {
    private long _fontBytes;
    private XElement? Glyphs(XElement e, Dictionary<string, Resource> scope, string part, int depth, BrushRegion region) {
        int glyphOrdinal = _visualDepth == 0 ? _nativeGlyphOrdinal++ : -1;
        Charge(depth);
        CheckAttributes(e, "FontUri FontRenderingEmSize OriginX OriginY UnicodeString Indices Fill BidiLevel IsSideways StyleSimulations RenderTransform Clip Opacity OpacityMask FixedPage.NavigateUri CaretStops DeviceFontName");
        string sidewaysValue = (string?)e.Attribute("IsSideways") ?? "false";
        if (sidewaysValue != "true" && sidewaysValue != "false" && sidewaysValue != "1" && sidewaysValue != "0")
            throw new InvalidDataException("Invalid IsSideways value.");
        bool sideways = sidewaysValue == "true" || sidewaysValue == "1";
        string simulation = (string?)e.Attribute("StyleSimulations") ?? "None";
        if (!new[] { "None", "BoldSimulation", "ItalicSimulation", "BoldItalicSimulation" }.Contains(simulation)) throw new InvalidDataException("Invalid glyph style simulation.");
        bool bold = simulation == "BoldSimulation" || simulation == "BoldItalicSimulation";
        bool italic = simulation == "ItalicSimulation" || simulation == "BoldItalicSimulation";
        foreach (var child in e.Elements()) if (!new[] { "Glyphs.Fill", "Glyphs.Clip", "Glyphs.RenderTransform", "Glyphs.OpacityMask" }.Contains(child.Name.LocalName)) Loss(child.Name.LocalName);
        string uri = (string?)e.Attribute("FontUri") ?? throw new InvalidDataException("Missing glyph font URI.");
        string[] uriParts = uri.Split('#');
        if (uriParts.Length > 2) throw new InvalidDataException("Invalid font face URI.");
        string name = XpsPackage.Resolve(part, uriParts[0]);
        int? faceIndex = uriParts.Length == 2 ? ParseInt(uriParts[1]) : null;
        string fontKey = name + "#" + faceIndex;
        if (!_fonts.TryGetValue(fontKey, out var font)) {
            if (_fonts.Count >= 256) throw new InvalidDataException("XPS page font count exceeded.");
            string type = _page.Document.ContentType(name);
            int fontLength = _page.Document.Part(name).Length;
            if (fontLength > 64 * 1024 * 1024 - _fontBytes) throw new InvalidDataException("XPS font materialization budget exceeded.");
            _fontBytes += fontLength;
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
        if (sideways && rtl) throw new InvalidDataException("Sideways glyphs require an even BidiLevel.");
        string text = XpsPage.Unescape((string?)e.Attribute("UnicodeString") ?? "");
        string indices = (string?)e.Attribute("Indices") ?? "";
        string[] entries = indices.Length == 0 ? Array.Empty<string>() : indices.Split(';');
        double shear = italic ? Math.Tan(20 * Math.PI / 180) : 0;
        double minX = double.PositiveInfinity, minY = double.PositiveInfinity, maxX = double.NegativeInfinity, maxY = double.NegativeInfinity;
        var data = new StringBuilder();
        int textIndex = 0, entryIndex = 0;
        int firstSpan = _textSpans.Count;
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
            double clusterLeft = double.PositiveInfinity, clusterTop = double.PositiveInfinity;
            double clusterRight = double.NegativeInfinity, clusterBottom = double.NegativeInfinity;
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
                double topX = 0, topY = 0;
                double nativeAdvance = sideways
                    ? font.FixedGlyphVerticalMetrics(glyph, size, out topX, out topY)
                    : font.FixedGlyphAdvance(glyph, size);
                if (bold) nativeAdvance += size * 0.02;
                double advance = fields.Length > 1 && fields[1].Length > 0 ? XpsPackage.Number(fields[1]) * size / 100 : nativeAdvance;
                if (advance < 0) throw new InvalidDataException("Negative glyph advance.");
                double u = fields.Length > 2 && fields[2].Length > 0 ? XpsPackage.Number(fields[2]) * size / 100 : 0;
                double v = fields.Length > 3 && fields[3].Length > 0 ? XpsPackage.Number(fields[3]) * size / 100 : 0;
                double gx = rtl ? x - nativeAdvance - u : x + u;
                double cellLeft = rtl ? x - u - advance : gx;
                double cellRight = rtl ? x - u : gx + advance;
                clusterLeft = Math.Min(clusterLeft, cellLeft); clusterRight = Math.Max(clusterRight, cellRight);
                if (!sideways) {
                    double cellTop = y - v - font.BaselineOffset(size);
                    clusterTop = Math.Min(clusterTop, cellTop); clusterBottom = Math.Max(clusterBottom, cellTop + font.LineHeight(size));
                }
                var contours = font.FixedGlyphContours(glyph, size, sideways ? 0 : gx, sideways ? 0 : y - v, Math.Max(1, 1000000 - _points), _token);
                foreach (var contour in contours) {
                    _points = checked(_points + contour.Count);
                    if (_points > 1000000) throw new InvalidDataException("XPS glyph outline point limit exceeded.");
                    for (int i = 0; i < contour.Count; i++) {
                        // Rotate the outline about its top-center origin; advance and offsets remain in run coordinates.
                        double px = sideways ? gx + contour[i].Y + topY : contour[i].X;
                        double py = sideways ? y - v - contour[i].X + topX : contour[i].Y;
                        if (italic) {
                            if (sideways) py += shear * (px - gx - topY);
                            else px -= shear * (py - (y - v));
                        }
                        if (bold) { px += size * 0.01; py -= size * 0.01; }
                        clusterLeft = Math.Min(clusterLeft, px); clusterTop = Math.Min(clusterTop, py);
                        clusterRight = Math.Max(clusterRight, px); clusterBottom = Math.Max(clusterBottom, py);
                        minX = Math.Min(minX, px); minY = Math.Min(minY, py); maxX = Math.Max(maxX, px); maxY = Math.Max(maxY, py);
                        string point = (i == 0 ? "M" : "L") + N(px) + " " + N(py) + " ";
                        EnsureOutputCapacity((long)data.Length + point.Length + 2);
                        data.Append(point);
                    }
                    data.Append("Z ");
                }
                x += rtl ? -advance : advance;
            }
            if (_visualDepth == 0 && textIndex < text.Length) {
                // Keep native clusters intact: a ligature or a surrogate pair must not
                // be split into unrelated PDF replacement-text sequences.
                double left = clusterLeft, right = clusterRight;
                double top = double.IsPositiveInfinity(clusterTop) ? y - size / 2 : clusterTop;
                double bottom = double.IsNegativeInfinity(clusterBottom) ? y + size / 2 : clusterBottom;
                // Zero-advance combining clusters still need a finite selection region.
                right = Math.Max(right, left + size * 0.001);
                bottom = Math.Max(bottom, top + size * 0.001);
                _textSpans.Add(new XpsTextSpan(text.Substring(textIndex, codeUnits),
                    new OfficePoint(rtl ? right : left, top), new OfficePoint(rtl ? left : right, top),
                    new OfficePoint(rtl ? left : right, bottom), new OfficePoint(rtl ? right : left, bottom), glyphOrdinal));
            }
            textIndex += codeUnits;
            entryIndex += glyphCount;
        }

        JoinWhitespace(firstSpan, rtl);
        if (_visualDepth == 0) _nativeBounds[e] = double.IsInfinity(minX) ? new BrushRegion(x, y, 0, 0)
            : new BrushRegion(minX, minY, maxX - minX, maxY - minY);
        if (data.Length == 0) {
            var empty = Element("g");
            if (text.Length > 0) Set(empty, "aria-label", text);
            return empty;
        }
        var path = Element("path", new XAttribute("d", data.ToString()), new XAttribute("fill-rule", "nonzero"));
        XElement result;
        if (bold) {
            // An opaque coverage mask unions the original silhouette and its
            // widening stroke. Paint the native brush once so translucent text
            // does not darken along the original outline or between glyphs.
            Set(path, "fill", "#ffffff"); Set(path, "stroke", "#ffffff");
            Set(path, "stroke-width", N(size * 0.02)); Set(path, "stroke-linejoin", "round");
            double growth = size * 0.01;
            region = IntersectRegion(region, new BrushRegion(minX - growth, minY - growth, maxX - minX + growth * 2, maxY - minY + growth * 2));
            var paint = Element("path", new XAttribute("d", "M" + N(region.X) + "," + N(region.Y) + " h" + N(region.Width) + " v" + N(region.Height) + " h" + N(-region.Width) + " Z"));
            Paint(e, "Fill", paint, "fill", scope, part, depth);
            result = Element("g", ApplyCoverageMask(path, ApplyBrushFill(paint), region));
        } else {
            Paint(e, "Fill", path, "fill", scope, part, depth);
            result = ApplyBrushFill(path);
        }
        if (text.Length > 0) Set(result, "aria-label", text);
        return result;
    }
    private static int ParseInt(string text) {
        if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int value) || value < 0 || value > 65535) throw new InvalidDataException("Invalid XPS integer.");
        return value;
    }
}
