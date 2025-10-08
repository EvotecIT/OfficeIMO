namespace OfficeIMO.Pdf;

internal static class ResourceResolver {
    public static Dictionary<string, PdfFontResource> GetFontsForPage(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects) {
        var fonts = new Dictionary<string, PdfFontResource>(System.StringComparer.Ordinal);
        var dict = GetInheritedDictionary(page, "Resources", objects);
        if (dict is null) return fonts;
        if (!dict.Items.TryGetValue("Font", out var fontDictObj)) return fonts;
        var fontDict = ResolveDict(fontDictObj, objects);
        if (fontDict is null) return fonts;
        foreach (var kv in fontDict.Items) {
            var fontVal = ResolveDict(kv.Value, objects);
            if (fontVal is null) continue;
            string baseFont = (fontVal.Get<PdfName>("BaseFont")?.Name) ?? "";
            string encoding = (fontVal.Get<PdfName>("Encoding")?.Name) ?? "WinAnsiEncoding"; // default heuristic
            bool hasToUnicode = fontVal.Items.ContainsKey("ToUnicode");
            ToUnicodeCMap? cmap = null;
            if (hasToUnicode) {
                if (fontVal.Items.TryGetValue("ToUnicode", out var tu) && tu is PdfReference r && objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfStream s) {
                    var data = PdfSyntax.HasFlateDecode(s.Dictionary) ? Filters.FlateDecoder.Decode(s.Data) : s.Data;
                    if (!ToUnicodeCMap.TryParse(data, out cmap)) cmap = null;
                }
            }
            fonts[kv.Key] = new PdfFontResource(kv.Key, baseFont, encoding, hasToUnicode, cmap);
        }
        return fonts;
    }

    public static Dictionary<string, System.Func<byte[], string>> GetFontDecoders(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects) {
        var map = new Dictionary<string, System.Func<byte[], string>>(System.StringComparer.Ordinal);
        var fonts = GetFontsForPage(page, objects);
        foreach (var kv in fonts) {
            var dec = BuildDecoderForFont(kv.Value);
            map[kv.Key] = dec;
        }
        return map;
    }

    public static Dictionary<string, System.Func<byte[], double>> GetFontWidthProviders(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects) {
        var map = new Dictionary<string, System.Func<byte[], double>>(System.StringComparer.Ordinal);
        var dict = GetInheritedDictionary(page, "Resources", objects);
        if (dict is null) return map;
        if (!dict.Items.TryGetValue("Font", out var fontDictObj)) return map;
        var fontDict = ResolveDict(fontDictObj, objects);
        if (fontDict is null) return map;
        foreach (var kv in fontDict.Items) {
            var fontVal = ResolveDict(kv.Value, objects);
            if (fontVal is null) continue;
            var subtype = fontVal.Get<PdfName>("Subtype")?.Name ?? string.Empty;
            if (string.Equals(subtype, "Type0", System.StringComparison.Ordinal)) {
                if (TryBuildCidWidthMap(fontVal, objects, out var cidMap)) {
                    var localMap = cidMap!;
                    map[kv.Key] = bytes => SumWidthsCid(bytes, localMap);
                } else {
                    map[kv.Key] = bytes => bytes != null ? (bytes.Length / 2) * 1000.0 : 0.0; // conservative default
                }
            } else {
                int firstChar = (int)(fontVal.Get<PdfNumber>("FirstChar")?.Value ?? 0);
                int lastChar = (int)(fontVal.Get<PdfNumber>("LastChar")?.Value ?? 255);
                var widths = ResolveArray(fontVal.Items.TryGetValue("Widths", out var wobj) ? wobj : null, objects);
                double[] tbl;
                if (widths is null) {
                    tbl = new double[lastChar - firstChar + 1];
                    for (int i = 0; i < tbl.Length; i++) tbl[i] = 500.0;
                } else {
                    tbl = new double[widths.Items.Count];
                    for (int i = 0; i < widths.Items.Count; i++) tbl[i] = widths.Items[i] is PdfNumber num ? num.Value : 500.0;
                }
                map[kv.Key] = bytes => SumWidthsSimple(bytes, firstChar, tbl);
            }
        }
        return map;
    }

    private static double SumWidthsSimple(byte[] bytes, int firstChar, double[] widths) {
        if (bytes == null) return 0.0;
        double sum = 0.0;
        for (int i = 0; i < bytes.Length; i++) {
            int code = bytes[i];
            int idx = code - firstChar;
            if (idx >= 0 && idx < widths.Length) sum += widths[idx]; else sum += 500.0;
        }
        return sum;
    }

    private sealed class CidWidthMap {
        public double DefaultWidth1000 { get; }
        public Dictionary<int, double> Widths { get; }
        public CidWidthMap(double dw, Dictionary<int, double> map) { DefaultWidth1000 = dw; Widths = map; }
    }

    private static bool TryBuildCidWidthMap(PdfDictionary type0Font, Dictionary<int, PdfIndirectObject> objects, out CidWidthMap? map) {
        map = null;
        if (!type0Font.Items.TryGetValue("DescendantFonts", out var dfObj)) return false;
        var dfArr = ResolveArray(dfObj, objects);
        if (dfArr is null || dfArr.Items.Count == 0) return false;
        var desc = ResolveDict(dfArr.Items[0], objects);
        if (desc is null) return false;
        double dw = (desc.Get<PdfNumber>("DW")?.Value) ?? 1000.0;
        var widthsObj = desc.Items.TryGetValue("W", out var w) ? w : null;
        var wArr = ResolveArray(widthsObj, objects);
        var dict = new Dictionary<int, double>();
        if (wArr is not null) {
            // Parse sequences: <startCid> [w1 w2 ...] | <startCid> <endCid> <w>
            for (int i = 0; i < wArr.Items.Count; i++) {
                var startObj = wArr.Items[i] as PdfNumber; if (startObj is null) break;
                int startCid = (int)startObj.Value; i++;
                if (i >= wArr.Items.Count) break;
                var next = wArr.Items[i];
                if (next is PdfArray list) {
                    for (int j = 0; j < list.Items.Count; j++) {
                        if (list.Items[j] is PdfNumber wn) dict[startCid + j] = wn.Value; else dict[startCid + j] = dw;
                    }
                } else if (next is PdfNumber endCidNum) {
                    int endCid = (int)endCidNum.Value; i++;
                    if (i >= wArr.Items.Count) break;
                    var wNum = wArr.Items[i] as PdfNumber; double wv = wNum?.Value ?? dw;
                    for (int cid = startCid; cid <= endCid; cid++) dict[cid] = wv;
                }
            }
        }
        map = new CidWidthMap(dw, dict); return true;
    }

    private static double SumWidthsCid(byte[] bytes, CidWidthMap map) {
        if (bytes == null || bytes.Length == 0) return 0.0;
        double sum = 0.0;
        // Assume Identity-H two-byte big-endian CIDs (common). If odd length, ignore trailing byte.
        for (int i = 0; i + 1 < bytes.Length; i += 2) {
            int cid = (bytes[i] << 8) | bytes[i + 1];
            if (!map.Widths.TryGetValue(cid, out var w)) w = map.DefaultWidth1000;
            sum += w;
        }
        return sum;
    }

    public static Dictionary<string, System.Func<byte[], string>> GetFontDecodersForForm(PdfDictionary formDict, Dictionary<int, PdfIndirectObject> objects) {
        var map = new Dictionary<string, System.Func<byte[], string>>(System.StringComparer.Ordinal);
        if (!formDict.Items.TryGetValue("Resources", out var resObj)) return map;
        var res = ResolveDict(resObj, objects);
        if (res is null) return map;
        if (!res.Items.TryGetValue("Font", out var fontObj)) return map;
        var fontDict = ResolveDict(fontObj, objects);
        if (fontDict is null) return map;
        foreach (var kv in fontDict.Items) {
            var fontVal = ResolveDict(kv.Value, objects);
            if (fontVal is null) continue;
            string baseFont = (fontVal.Get<PdfName>("BaseFont")?.Name) ?? "";
            string encoding = (fontVal.Get<PdfName>("Encoding")?.Name) ?? "WinAnsiEncoding";
            bool hasToUnicode = fontVal.Items.ContainsKey("ToUnicode");
            ToUnicodeCMap? cmap = null;
            if (hasToUnicode) {
                if (fontVal.Items.TryGetValue("ToUnicode", out var tu) && tu is PdfReference r && objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfStream s) {
                    var data = PdfSyntax.HasFlateDecode(s.Dictionary) ? Filters.FlateDecoder.Decode(s.Data) : s.Data;
                    if (!ToUnicodeCMap.TryParse(data, out cmap)) cmap = null;
                }
            }
            var resName = kv.Key;
            var dec = BuildDecoderForFont(new PdfFontResource(resName, baseFont, encoding, hasToUnicode, cmap));
            map[resName] = dec;
        }
        return map;
    }

    public static Dictionary<string, byte[]> GetFormXObjectStreams(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects) {
        var result = new Dictionary<string, byte[]>(System.StringComparer.Ordinal);
        var res = GetInheritedDictionary(page, "Resources", objects);
        if (res is null) return result;
        if (!res.Items.TryGetValue("XObject", out var xoObj)) return result;
        var xo = ResolveDict(xoObj, objects);
        if (xo is null) return result;
        foreach (var kv in xo.Items) {
            if (kv.Value is PdfReference r && objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfStream s) {
                var subtype = s.Dictionary.Get<PdfName>("Subtype")?.Name;
                if (string.Equals(subtype, "Form", System.StringComparison.Ordinal)) {
                    var data = PdfSyntax.HasFlateDecode(s.Dictionary) ? Filters.FlateDecoder.Decode(s.Data) : s.Data;
                    result[kv.Key] = data;
                }
            }
        }
        return result;
    }

    private static System.Func<byte[], string> BuildDecoderForFont(PdfFontResource font) {
        // Prefer font-specific ToUnicode map when present
        if (font.HasToUnicode && font.CMap is not null) return font.CMap.MapBytes;
        // Fall back to WinAnsi
        return PdfWinAnsiEncoding.Decode;
    }

    private static PdfDictionary? GetInheritedDictionary(PdfDictionary page, string key, Dictionary<int, PdfIndirectObject> objects) {
        PdfDictionary? current = page;
        while (current is not null) {
            if (current.Items.TryGetValue(key, out var v)) {
                var dict = ResolveDict(v, objects);
                if (dict is not null) return dict;
            }
            if (!current.Items.TryGetValue("Parent", out var p) || p is not PdfReference pr || !objects.TryGetValue(pr.ObjectNumber, out var indr) || indr.Value is not PdfDictionary parent) break;
            current = parent;
        }
        return null;
    }

    private static PdfDictionary? ResolveDict(PdfObject obj, Dictionary<int, PdfIndirectObject> objects) {
        if (obj is PdfDictionary d) return d;
        if (obj is PdfReference r && objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfDictionary dd) return dd;
        return null;
    }

    private static PdfArray? ResolveArray(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        if (obj is null) return null;
        if (obj is PdfArray a) return a;
        if (obj is PdfReference r && objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfArray aa) return aa;
        return null;
    }
}
