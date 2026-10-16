namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    private const int MinimumUnitsPerEm = 16;
    private const int MaximumUnitsPerEm = 16384;

    internal static bool IsCffProgramSubtype(string? subtype) =>
        subtype is "Type1" or "Type1C" or "CIDFontType0C" or "OpenType";

    /// <summary>
    /// Creates a drawing program for an embedded CFF (FontFile3 Type1C, CIDFontType0C, OpenType CFF) or
    /// Type 1 (FontFile) program, or null when the program cannot be read.
    /// </summary>
    internal static PdfDrawingFontProgram? TryCreateCffDrawingProgram(PdfFontResource resource, PdfDictionary font,
        Dictionary<int, PdfIndirectObject> objects, byte[] program, string? programSubtype) {
        string?[]? type1Encoding = null;
        byte[]? cffData;
        if (string.Equals(programSubtype, "Type1", StringComparison.Ordinal)) {
            PdfType1FontConverter.Result? converted = PdfType1FontConverter.TryConvert(program);
            if (converted == null) return null;
            cffData = converted.Cff;
            type1Encoding = converted.BuiltinEncoding;
        } else if (string.Equals(programSubtype, "OpenType", StringComparison.Ordinal)) {
            cffData = PdfCffFontProgram.TryExtractOpenTypeCff(program);
        } else {
            cffData = program;
        }
        if (cffData == null) return null;
        PdfCffFontProgram? cff = PdfCffFontProgram.TryParse(cffData);
        if (cff == null) return null;

        double[] matrix = cff.FontMatrix;
        double scale = Math.Abs(matrix[0]) > 0D ? Math.Abs(matrix[0]) : Math.Abs(matrix[3]);
        if (scale <= 0D || double.IsNaN(scale)) return null;
        int unitsPerEm = (int)Math.Round(1D / scale);
        if (unitsPerEm < MinimumUnitsPerEm || unitsPerEm > MaximumUnitsPerEm) return null;

        var advances = new double[cff.GlyphCount];
        for (int glyph = 0; glyph < advances.Length; glyph++) advances[glyph] = cff.GetAdvance(glyph) ?? 0D;

        if (string.Equals(resource.FontSubtype, "Type0", StringComparison.Ordinal)) {
            byte[] cidToGlyph = BuildCffCidToGlyphMap(cff);
            if (TryBuildCidWidthMap(font, objects, out CidWidthMap? widths) && widths != null) {
                for (int cid = 0; 2 * cid + 1 < cidToGlyph.Length && cid < CidCodeCount; cid++) {
                    int glyph = (cidToGlyph[2 * cid] << 8) | cidToGlyph[2 * cid + 1];
                    if (glyph > 0 && glyph < advances.Length) advances[glyph] = widths.GetWidth(cid) / 1000D / scale;
                }
            }
            return PdfTrueTypeUnicodeCmap.TryCreateFromCff(resource, cff, null, cidToGlyph, advances, unitsPerEm, program.LongLength);
        }

        string?[] names = ResolveSimpleGlyphNames(font, objects, cff, type1Encoding);
        var codeToGlyph = new int[256];
        for (int code = 0; code < 256; code++) codeToGlyph[code] = ResolveSimpleCffGlyph(cff, names[code], code, type1Encoding);
        ApplySimpleWidths(font, objects, codeToGlyph, advances, scale);
        return PdfTrueTypeUnicodeCmap.TryCreateFromCff(resource, cff, code => codeToGlyph[code], null, advances, unitsPerEm, program.LongLength);
    }

    private static int ResolveSimpleCffGlyph(PdfCffFontProgram cff, string? name, int code, string?[]? type1Encoding) {
        if (name != null) {
            if (cff.TryGetGlyph(name, out int named)) return named;
            // Producers sometimes name a glyph by a Unicode alias the subset spells differently.
            if (TryDecodeGlyphName(name, out string? text) && text is { Length: 1 } &&
                cff.TryGetGlyph("uni" + ((int)text[0]).ToString("X4", System.Globalization.CultureInfo.InvariantCulture), out int alias) && alias > 0) return alias;
            return 0;
        }
        if (type1Encoding != null) {
            return type1Encoding[code] is string builtin && cff.TryGetGlyph(builtin, out int builtinGlyph) ? builtinGlyph : 0;
        }
        return cff.BuiltinEncoding[code];
    }

    // ISO 32000-1, 9.6.6.2: Differences apply over the named base encoding, or over the font's built-in
    // encoding when none is named. Viewers use the built-in encoding for embedded programs without /Encoding.
    private static string?[] ResolveSimpleGlyphNames(PdfDictionary font, Dictionary<int, PdfIndirectObject> objects,
        PdfCffFontProgram cff, string?[]? type1Encoding) {
        var names = new string?[256];
        string? baseName = null;
        PdfArray? differences = null;
        PdfObject? encoding = font.Items.TryGetValue("Encoding", out PdfObject? encodingValue) ? ResolveObject(encodingValue, objects) : null;
        if (encoding is PdfName encodingName) {
            baseName = encodingName.Name;
        } else if (encoding is PdfDictionary encodingDictionary) {
            baseName = encodingDictionary.Get<PdfName>("BaseEncoding")?.Name;
            differences = ResolveArray(encodingDictionary.Items.TryGetValue("Differences", out PdfObject? differencesValue) ? differencesValue : null, objects);
        }
        string?[]? baseNames = PdfFontEncodingNames.Get(baseName);
        if (baseNames != null) {
            Array.Copy(baseNames, names, 256);
        } else if (type1Encoding != null) {
            Array.Copy(type1Encoding, names, 256);
        } else {
            for (int code = 0; code < 256; code++) {
                int glyph = cff.BuiltinEncoding[code];
                if (glyph > 0 && cff.GlyphNames != null && glyph < cff.GlyphNames.Length) names[code] = cff.GlyphNames[glyph];
            }
        }
        if (differences != null) {
            int code = -1;
            foreach (PdfObject item in differences.Items) {
                PdfObject? resolved = ResolveObject(item, objects);
                if (resolved is PdfNumber number) {
                    code = double.IsNaN(number.Value) || number.Value < 0D || number.Value > 255D ? -1 : (int)number.Value;
                } else if (resolved is PdfName glyphName && code >= 0 && code <= 255) {
                    names[code++] = glyphName.Name;
                }
            }
        }
        return names;
    }

    private static void ApplySimpleWidths(PdfDictionary font, Dictionary<int, PdfIndirectObject> objects, int[] codeToGlyph,
        double[] advances, double scale) {
        PdfArray? widths = ResolveArray(font.Items.TryGetValue("Widths", out PdfObject? widthsValue) ? widthsValue : null, objects);
        if (widths == null || ResolveObject(font.Items.TryGetValue("FirstChar", out PdfObject? firstValue) ? firstValue : null, objects) is not PdfNumber first) return;
        if (double.IsNaN(first.Value) || first.Value < 0D || first.Value > 255D) return;
        int firstChar = (int)first.Value;
        var assigned = new HashSet<int>();
        for (int index = 0; index < widths.Items.Count && firstChar + index < 256; index++) {
            int glyph = codeToGlyph[firstChar + index];
            if (glyph <= 0 || glyph >= advances.Length || !assigned.Add(glyph) ||
                ResolveObject(widths.Items[index], objects) is not PdfNumber width ||
                double.IsNaN(width.Value) || double.IsInfinity(width.Value) || width.Value < 0D) continue;
            advances[glyph] = width.Value / 1000D / scale;
        }
    }

    private static byte[] BuildCffCidToGlyphMap(PdfCffFontProgram cff) {
        int maximumCid = cff.GlyphCount - 1;
        if (cff.Cids != null) {
            maximumCid = 0;
            foreach (int cid in cff.Cids) maximumCid = Math.Max(maximumCid, Math.Min(cid, CidCodeCount - 1));
        }
        var map = new byte[(maximumCid + 1) * 2];
        if (cff.Cids == null) {
            for (int glyph = 1; glyph <= maximumCid; glyph++) {
                map[2 * glyph] = (byte)(glyph >> 8);
                map[2 * glyph + 1] = (byte)glyph;
            }
            return map;
        }
        for (int glyph = cff.Cids.Length - 1; glyph > 0; glyph--) {
            int cid = cff.Cids[glyph];
            if (cid <= 0 || cid > maximumCid) continue;
            map[2 * cid] = (byte)(glyph >> 8);
            map[2 * cid + 1] = (byte)glyph;
        }
        return map;
    }
}
