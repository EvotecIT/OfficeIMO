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
}
