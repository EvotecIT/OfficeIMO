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
            fonts[kv.Key] = new PdfFontResource(kv.Key, baseFont, encoding, hasToUnicode);
        }
        return fonts;
    }

    public static Dictionary<string, System.Func<byte[], string>> GetFontDecoders(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects) {
        var map = new Dictionary<string, System.Func<byte[], string>>(System.StringComparer.Ordinal);
        var fonts = GetFontsForPage(page, objects);
        foreach (var kv in fonts) {
            var decoder = BuildDecoderForFont(kv.Value, objects);
            map[kv.Key] = decoder;
        }
        return map;
    }

    private static System.Func<byte[], string> BuildDecoderForFont(PdfFontResource font, Dictionary<int, PdfIndirectObject> objects) {
        // ToUnicode takes precedence
        if (font.HasToUnicode && TryGetToUnicodeCMap(font.ResourceName, objects, out var cmap)) {
            return bytes => cmap!.MapBytes(bytes);
        }
        // Fall back to WinAnsi
        return PdfWinAnsiEncoding.Decode;
    }

    private static bool TryGetToUnicodeCMap(string resourceName, Dictionary<int, PdfIndirectObject> objects, out ToUnicodeCMap? cmap) {
        // Search any font object with this resource name (best-effort for now)
        foreach (var obj in objects.Values) {
            if (obj.Value is PdfDictionary d) {
                if (d.Get<PdfName>("Type")?.Name == "Font") {
                    if (d.Items.TryGetValue("ToUnicode", out var tu)) {
                        if (tu is PdfReference r && objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfStream s) {
                            if (ToUnicodeCMap.TryParse(s.Data, out cmap)) return true;
                        }
                    }
                }
            }
        }
        cmap = null; return false;
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
