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
            fonts[kv.Key] = CreateFontResource(kv.Key, fontVal, objects);
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
            var fontResource = CreateFontResource(kv.Key, fontVal, objects);
            var subtype = fontVal.Get<PdfName>("Subtype")?.Name ?? string.Empty;
            if (string.Equals(subtype, "Type0", System.StringComparison.Ordinal)) {
                if (TryBuildCidWidthMap(fontVal, objects, out var cidMap)) {
                    var localMap = cidMap!;
                    map[kv.Key] = bytes => SumWidthsCid(bytes, localMap);
                } else {
                    map[kv.Key] = bytes => bytes != null ? (bytes.Length / 2d) * 1000.0 : 0.0; // conservative default
                }
            } else {
                int firstChar = (int)(fontVal.Get<PdfNumber>("FirstChar")?.Value ?? 0);
                var widths = ResolveArray(fontVal.Items.TryGetValue("Widths", out var wobj) ? wobj : null, objects);
                if (widths is null) {
                    if (PdfWriter.TryGetStandardFontByBaseFontName(fontResource.BaseFont, out var standardFont)) {
                        var decoder = BuildDecoderForFont(fontResource);
                        map[kv.Key] = bytes => PdfWriter.EstimateSimpleTextWidth1000(decoder(bytes), standardFont);
                    } else {
                        map[kv.Key] = bytes => (bytes?.Length ?? 0) * 500.0;
                    }

                    continue;
                } else {
                    var tbl = new double[widths.Items.Count];
                    for (int i = 0; i < widths.Items.Count; i++) tbl[i] = widths.Items[i] is PdfNumber num ? num.Value : 500.0;
                    map[kv.Key] = bytes => SumWidthsSimple(bytes, firstChar, tbl);
                }
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
            var resName = kv.Key;
            var dec = BuildDecoderForFont(CreateFontResource(resName, fontVal, objects));
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
            if (kv.Value is PdfReference r && PdfObjectLookup.TryGet(objects, r, out var ind) && ind.Value is PdfStream s) {
                var subtype = s.Dictionary.Get<PdfName>("Subtype")?.Name;
                if (string.Equals(subtype, "Form", System.StringComparison.Ordinal)) {
                    var data = Filters.StreamDecoder.Decode(s.Dictionary, s.Data, objects);
                    result[kv.Key] = data;
                }
            }
        }
        return result;
    }

    public static IReadOnlyList<PdfExtractedImage> GetImageXObjectsForPage(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects, int pageNumber) {
        var result = new List<PdfExtractedImage>();
        var res = GetInheritedDictionary(page, "Resources", objects);
        if (res is null) return result;
        if (!res.Items.TryGetValue("XObject", out var xoObj)) return result;
        var xo = ResolveDict(xoObj, objects);
        if (xo is null) return result;

        foreach (var kv in xo.Items) {
            int objectNumber = 0;
            PdfStream? stream = null;
            if (kv.Value is PdfReference reference &&
                PdfObjectLookup.TryGet(objects, reference, out var indirect) &&
                indirect.Value is PdfStream referencedStream) {
                objectNumber = reference.ObjectNumber;
                stream = referencedStream;
            } else if (kv.Value is PdfStream directStream) {
                stream = directStream;
            }

            if (stream is null) {
                continue;
            }

            var subtype = stream.Dictionary.Get<PdfName>("Subtype")?.Name;
            if (!string.Equals(subtype, "Image", System.StringComparison.Ordinal)) {
                continue;
            }

            result.Add(BuildExtractedImage(pageNumber, kv.Key, objectNumber, stream, objects));
        }

        return result;
    }

    private static System.Func<byte[], string> BuildDecoderForFont(PdfFontResource font) {
        // Prefer font-specific ToUnicode map when present
        if (font.HasToUnicode && font.CMap is not null) return font.CMap.MapBytes;
        var baseDecoder = BuildBaseEncodingDecoder(font.Encoding);
        if (font.Differences is not null && font.Differences.Count > 0) {
            var differences = font.Differences;
            return bytes => DecodeWithDifferences(bytes, differences, baseDecoder);
        }

        return baseDecoder;
    }

    private static System.Func<byte[], string> BuildBaseEncodingDecoder(string encoding) {
        if (string.Equals(encoding, "StandardEncoding", System.StringComparison.Ordinal)) {
            return PdfStandardEncoding.Decode;
        }

        if (string.Equals(encoding, "MacRomanEncoding", System.StringComparison.Ordinal)) {
            return PdfMacRomanEncoding.Decode;
        }

        return PdfWinAnsiEncoding.Decode;
    }

    private static PdfFontResource CreateFontResource(string resourceName, PdfDictionary fontVal, Dictionary<int, PdfIndirectObject> objects) {
        string baseFont = (fontVal.Get<PdfName>("BaseFont")?.Name) ?? "";
        string encoding = GetDefaultEncodingForBaseFont(baseFont);
        IReadOnlyDictionary<int, string>? differences = null;
        if (fontVal.Items.TryGetValue("Encoding", out var encodingObj)) {
            if (ResolveObject(encodingObj, objects) is PdfName encodingName) {
                encoding = encodingName.Name;
            } else if (ResolveDict(encodingObj, objects) is PdfDictionary encodingDict) {
                encoding = encodingDict.Get<PdfName>("BaseEncoding")?.Name ?? encoding;
                differences = BuildDifferencesMap(encodingDict.Items.TryGetValue("Differences", out var diffObj) ? diffObj : null, objects);
            }
        }

        bool hasToUnicode = fontVal.Items.ContainsKey("ToUnicode");
        ToUnicodeCMap? cmap = null;
        if (hasToUnicode) {
            if (fontVal.Items.TryGetValue("ToUnicode", out var tu) &&
                tu is PdfReference r &&
                PdfObjectLookup.TryGet(objects, r, out var ind) &&
                ind.Value is PdfStream s) {
                var data = Filters.StreamDecoder.Decode(s.Dictionary, s.Data, objects);
                if (!ToUnicodeCMap.TryParse(data, out cmap)) cmap = null;
            }
        }

        return new PdfFontResource(resourceName, baseFont, encoding, hasToUnicode, cmap, differences);
    }

    private static string GetDefaultEncodingForBaseFont(string baseFont) {
        string fontName = StripSubsetPrefix(baseFont);
        switch (fontName) {
            case "Courier":
            case "Courier-Bold":
            case "Courier-Oblique":
            case "Courier-BoldOblique":
            case "Helvetica":
            case "Helvetica-Bold":
            case "Helvetica-Oblique":
            case "Helvetica-BoldOblique":
            case "Times-Roman":
            case "Times-Bold":
            case "Times-Italic":
            case "Times-BoldItalic":
                return "StandardEncoding";
            default:
                return "WinAnsiEncoding";
        }
    }

    private static string StripSubsetPrefix(string baseFont) {
        if (baseFont.Length > 7 && baseFont[6] == '+') {
            for (int i = 0; i < 6; i++) {
                char ch = baseFont[i];
                if (ch < 'A' || ch > 'Z') {
                    return baseFont;
                }
            }

            return baseFont.Substring(7);
        }

        return baseFont;
    }

    private static Dictionary<int, string>? BuildDifferencesMap(PdfObject? differencesObj, Dictionary<int, PdfIndirectObject> objects) {
        var differences = ResolveArray(differencesObj, objects);
        if (differences is null) return null;

        var map = new Dictionary<int, string>();
        int code = -1;
        foreach (var item in differences.Items) {
            if (ResolveObject(item, objects) is PdfNumber number) {
                code = (int)number.Value;
                continue;
            }

            if (code < 0 || code > 255) {
                continue;
            }

            if (ResolveObject(item, objects) is PdfName glyphName &&
                TryDecodeGlyphName(glyphName.Name, out string? value)) {
                map[code] = value!;
            }

            code++;
        }

        return map.Count == 0 ? null : map;
    }

    private static string DecodeWithDifferences(byte[] bytes, IReadOnlyDictionary<int, string> differences, System.Func<byte[], string> baseDecoder) {
        if (bytes is null || bytes.Length == 0) return string.Empty;
        var builder = new System.Text.StringBuilder(bytes.Length);
        for (int i = 0; i < bytes.Length; i++) {
            int code = bytes[i];
            if (differences.TryGetValue(code, out string? value)) {
                builder.Append(value);
            } else {
                builder.Append(baseDecoder(new[] { bytes[i] }));
            }
        }

        return builder.ToString();
    }

    private static bool TryDecodeGlyphName(string glyphName, out string? value) {
        value = null;
        if (string.IsNullOrEmpty(glyphName) || string.Equals(glyphName, ".notdef", System.StringComparison.Ordinal)) {
            return false;
        }

        int variantIndex = glyphName.IndexOf('.');
        if (variantIndex > 0) {
            glyphName = glyphName.Substring(0, variantIndex);
        }

        if (glyphName.Length == 1) {
            value = glyphName;
            return true;
        }

        if (TryDecodeUnicodeGlyphName(glyphName, out value)) {
            return true;
        }

        if (TryDecodeCompositeGlyphName(glyphName, out value)) {
            return true;
        }

        switch (glyphName) {
            case "space": value = " "; return true;
            case "exclam": value = "!"; return true;
            case "quotedbl": value = "\""; return true;
            case "numbersign": value = "#"; return true;
            case "dollar": value = "$"; return true;
            case "percent": value = "%"; return true;
            case "ampersand": value = "&"; return true;
            case "quotesingle": value = "'"; return true;
            case "parenleft": value = "("; return true;
            case "parenright": value = ")"; return true;
            case "asterisk": value = "*"; return true;
            case "plus": value = "+"; return true;
            case "comma": value = ","; return true;
            case "hyphen": value = "-"; return true;
            case "period": value = "."; return true;
            case "slash": value = "/"; return true;
            case "colon": value = ":"; return true;
            case "semicolon": value = ";"; return true;
            case "less": value = "<"; return true;
            case "equal": value = "="; return true;
            case "greater": value = ">"; return true;
            case "question": value = "?"; return true;
            case "at": value = "@"; return true;
            case "bracketleft": value = "["; return true;
            case "backslash": value = "\\"; return true;
            case "bracketright": value = "]"; return true;
            case "asciicircum": value = "^"; return true;
            case "underscore": value = "_"; return true;
            case "grave": value = "`"; return true;
            case "circumflex": value = "\u02C6"; return true;
            case "tilde": value = "\u02DC"; return true;
            case "braceleft": value = "{"; return true;
            case "bar": value = "|"; return true;
            case "braceright": value = "}"; return true;
            case "asciitilde": value = "~"; return true;
            case "exclamdown": value = "\u00A1"; return true;
            case "cent": value = "\u00A2"; return true;
            case "sterling": value = "\u00A3"; return true;
            case "currency": value = "\u00A4"; return true;
            case "yen": value = "\u00A5"; return true;
            case "brokenbar": value = "\u00A6"; return true;
            case "section": value = "\u00A7"; return true;
            case "dieresis": value = "\u00A8"; return true;
            case "copyright": value = "\u00A9"; return true;
            case "ordfeminine": value = "\u00AA"; return true;
            case "guillemotleft": value = "\u00AB"; return true;
            case "logicalnot": value = "\u00AC"; return true;
            case "registered": value = "\u00AE"; return true;
            case "macron": value = "\u00AF"; return true;
            case "degree": value = "\u00B0"; return true;
            case "plusminus": value = "\u00B1"; return true;
            case "twosuperior": value = "\u00B2"; return true;
            case "threesuperior": value = "\u00B3"; return true;
            case "acute": value = "\u00B4"; return true;
            case "mu": value = "\u00B5"; return true;
            case "paragraph": value = "\u00B6"; return true;
            case "periodcentered": value = "\u00B7"; return true;
            case "cedilla": value = "\u00B8"; return true;
            case "onesuperior": value = "\u00B9"; return true;
            case "ordmasculine": value = "\u00BA"; return true;
            case "guillemotright": value = "\u00BB"; return true;
            case "onequarter": value = "\u00BC"; return true;
            case "onehalf": value = "\u00BD"; return true;
            case "threequarters": value = "\u00BE"; return true;
            case "questiondown": value = "\u00BF"; return true;
            case "Agrave": value = "\u00C0"; return true;
            case "Aacute": value = "\u00C1"; return true;
            case "Acircumflex": value = "\u00C2"; return true;
            case "Atilde": value = "\u00C3"; return true;
            case "Adieresis": value = "\u00C4"; return true;
            case "Aring": value = "\u00C5"; return true;
            case "AE": value = "\u00C6"; return true;
            case "Ccedilla": value = "\u00C7"; return true;
            case "Egrave": value = "\u00C8"; return true;
            case "Eacute": value = "\u00C9"; return true;
            case "Ecircumflex": value = "\u00CA"; return true;
            case "Edieresis": value = "\u00CB"; return true;
            case "Igrave": value = "\u00CC"; return true;
            case "Iacute": value = "\u00CD"; return true;
            case "Icircumflex": value = "\u00CE"; return true;
            case "Idieresis": value = "\u00CF"; return true;
            case "Eth": value = "\u00D0"; return true;
            case "Ntilde": value = "\u00D1"; return true;
            case "Ograve": value = "\u00D2"; return true;
            case "Oacute": value = "\u00D3"; return true;
            case "Ocircumflex": value = "\u00D4"; return true;
            case "Otilde": value = "\u00D5"; return true;
            case "Odieresis": value = "\u00D6"; return true;
            case "multiply": value = "\u00D7"; return true;
            case "Oslash": value = "\u00D8"; return true;
            case "Ugrave": value = "\u00D9"; return true;
            case "Uacute": value = "\u00DA"; return true;
            case "Ucircumflex": value = "\u00DB"; return true;
            case "Udieresis": value = "\u00DC"; return true;
            case "Yacute": value = "\u00DD"; return true;
            case "Thorn": value = "\u00DE"; return true;
            case "germandbls": value = "\u00DF"; return true;
            case "agrave": value = "\u00E0"; return true;
            case "aacute": value = "\u00E1"; return true;
            case "acircumflex": value = "\u00E2"; return true;
            case "atilde": value = "\u00E3"; return true;
            case "adieresis": value = "\u00E4"; return true;
            case "aring": value = "\u00E5"; return true;
            case "ae": value = "\u00E6"; return true;
            case "ccedilla": value = "\u00E7"; return true;
            case "egrave": value = "\u00E8"; return true;
            case "eacute": value = "\u00E9"; return true;
            case "ecircumflex": value = "\u00EA"; return true;
            case "edieresis": value = "\u00EB"; return true;
            case "igrave": value = "\u00EC"; return true;
            case "iacute": value = "\u00ED"; return true;
            case "icircumflex": value = "\u00EE"; return true;
            case "idieresis": value = "\u00EF"; return true;
            case "eth": value = "\u00F0"; return true;
            case "ntilde": value = "\u00F1"; return true;
            case "ograve": value = "\u00F2"; return true;
            case "oacute": value = "\u00F3"; return true;
            case "ocircumflex": value = "\u00F4"; return true;
            case "otilde": value = "\u00F5"; return true;
            case "odieresis": value = "\u00F6"; return true;
            case "divide": value = "\u00F7"; return true;
            case "oslash": value = "\u00F8"; return true;
            case "ugrave": value = "\u00F9"; return true;
            case "uacute": value = "\u00FA"; return true;
            case "ucircumflex": value = "\u00FB"; return true;
            case "udieresis": value = "\u00FC"; return true;
            case "yacute": value = "\u00FD"; return true;
            case "thorn": value = "\u00FE"; return true;
            case "ydieresis": value = "\u00FF"; return true;
            case "Scaron": value = "\u0160"; return true;
            case "scaron": value = "\u0161"; return true;
            case "Zcaron": value = "\u017D"; return true;
            case "zcaron": value = "\u017E"; return true;
            case "Ydieresis": value = "\u0178"; return true;
            case "florin": value = "\u0192"; return true;
            case "OE": value = "\u0152"; return true;
            case "oe": value = "\u0153"; return true;
            case "Lslash": value = "\u0141"; return true;
            case "lslash": value = "\u0142"; return true;
            case "Euro": value = "\u20AC"; return true;
            case "bullet": value = "\u2022"; return true;
            case "dagger": value = "\u2020"; return true;
            case "daggerdbl": value = "\u2021"; return true;
            case "endash": value = "\u2013"; return true;
            case "emdash": value = "\u2014"; return true;
            case "quoteleft": value = "\u2018"; return true;
            case "quoteright": value = "\u2019"; return true;
            case "quotesinglbase": value = "\u201A"; return true;
            case "quotedblleft": value = "\u201C"; return true;
            case "quotedblright": value = "\u201D"; return true;
            case "quotedblbase": value = "\u201E"; return true;
            case "ellipsis": value = "\u2026"; return true;
            case "perthousand": value = "\u2030"; return true;
            case "guilsinglleft": value = "\u2039"; return true;
            case "guilsinglright": value = "\u203A"; return true;
            case "fi": value = "fi"; return true;
            case "fl": value = "fl"; return true;
            default: return false;
        }
    }

    private static bool TryDecodeCompositeGlyphName(string glyphName, out string? value) {
        value = null;
        if (glyphName.IndexOf('_') < 0) {
            return false;
        }

        var builder = new System.Text.StringBuilder(glyphName.Length);
        string[] parts = glyphName.Split('_');
        foreach (string part in parts) {
            if (string.IsNullOrEmpty(part) || !TryDecodeGlyphName(part, out string? partValue)) {
                return false;
            }

            builder.Append(partValue);
        }

        value = builder.ToString();
        return true;
    }

    private static bool TryDecodeUnicodeGlyphName(string glyphName, out string? value) {
        value = null;
        if (glyphName.Length > 3 &&
            glyphName.StartsWith("uni", System.StringComparison.Ordinal) &&
            (glyphName.Length - 3) % 4 == 0) {
            var builder = new System.Text.StringBuilder((glyphName.Length - 3) / 4);
            for (int i = 3; i < glyphName.Length; i += 4) {
                if (!TryParseHexCodePoint(glyphName.Substring(i, 4), out int codePoint)) {
                    return false;
                }

                builder.Append((char)codePoint);
            }

            value = builder.ToString();
            return true;
        }

        if (glyphName.Length >= 5 &&
            glyphName.Length <= 7 &&
            glyphName[0] == 'u' &&
            TryParseHexCodePoint(glyphName.Substring(1), out int scalar) &&
            scalar <= 0x10FFFF) {
#if NET5_0_OR_GREATER
            value = char.ConvertFromUtf32(scalar);
#else
            value = scalar <= 0xFFFF ? ((char)scalar).ToString() : char.ConvertFromUtf32(scalar);
#endif
            return true;
        }

        return false;
    }

    private static bool TryParseHexCodePoint(string text, out int value) {
        value = 0;
        for (int i = 0; i < text.Length; i++) {
            int digit;
            char ch = text[i];
            if (ch >= '0' && ch <= '9') digit = ch - '0';
            else if (ch >= 'A' && ch <= 'F') digit = ch - 'A' + 10;
            else if (ch >= 'a' && ch <= 'f') digit = ch - 'a' + 10;
            else return false;
            value = (value << 4) | digit;
        }

        return true;
    }

    private static PdfDictionary? GetInheritedDictionary(PdfDictionary page, string key, Dictionary<int, PdfIndirectObject> objects) {
        PdfDictionary? current = page;
        while (current is not null) {
            if (current.Items.TryGetValue(key, out var v)) {
                var dict = ResolveDict(v, objects);
                if (dict is not null) return dict;
            }
            if (!current.Items.TryGetValue("Parent", out var p) || p is not PdfReference pr || !PdfObjectLookup.TryGet(objects, pr, out var indr) || indr.Value is not PdfDictionary parent) break;
            current = parent;
        }
        return null;
    }

    private static PdfDictionary? ResolveDict(PdfObject obj, Dictionary<int, PdfIndirectObject> objects) {
        if (obj is PdfDictionary d) return d;
        if (obj is PdfReference r && PdfObjectLookup.TryGet(objects, r, out var ind) && ind.Value is PdfDictionary dd) return dd;
        return null;
    }

    private static PdfArray? ResolveArray(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        if (obj is null) return null;
        if (obj is PdfArray a) return a;
        if (obj is PdfReference r && PdfObjectLookup.TryGet(objects, r, out var ind) && ind.Value is PdfArray aa) return aa;
        return null;
    }

    private static PdfExtractedImage BuildExtractedImage(
        int pageNumber,
        string resourceName,
        int objectNumber,
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects) {
        int width = (int)(stream.Dictionary.Get<PdfNumber>("Width")?.Value ?? 0);
        int height = (int)(stream.Dictionary.Get<PdfNumber>("Height")?.Value ?? 0);
        int bitsPerComponent = (int)(stream.Dictionary.Get<PdfNumber>("BitsPerComponent")?.Value ?? 0);
        string colorSpace = GetNameOrEmpty(stream.Dictionary.Items.TryGetValue("ColorSpace", out var colorSpaceObj) ? colorSpaceObj : null, objects);
        string filter = GetFilterName(stream.Dictionary.Items.TryGetValue("Filter", out var filterObj) ? filterObj : null, objects);

        byte[] bytes = stream.Data;
        string? extension = null;
        string? mimeType = null;
        bool isImageFile = false;

        if (string.Equals(filter, "DCTDecode", System.StringComparison.Ordinal)) {
            extension = "jpg";
            mimeType = "image/jpeg";
            isImageFile = true;
        } else if (string.Equals(filter, "FlateDecode", System.StringComparison.Ordinal) &&
                   TryBuildPngFile(stream, width, height, bitsPerComponent, colorSpace, objects, out var pngBytes)) {
            bytes = pngBytes;
            extension = "png";
            mimeType = "image/png";
            isImageFile = true;
        }

        return new PdfExtractedImage(
            pageNumber,
            resourceName,
            objectNumber,
            width,
            height,
            bitsPerComponent,
            colorSpace,
            filter,
            bytes,
            extension,
            mimeType,
            isImageFile);
    }

    private static string GetFilterName(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        var resolved = ResolveObject(obj, objects);
        if (resolved is PdfName name) {
            return name.Name;
        }

        if (resolved is PdfArray array) {
            var names = new List<string>();
            foreach (var item in array.Items) {
                var itemResolved = ResolveObject(item, objects);
                if (itemResolved is PdfName itemName) {
                    names.Add(itemName.Name);
                }
            }

            return string.Join(",", names);
        }

        return string.Empty;
    }

    private static string GetNameOrEmpty(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        var resolved = ResolveObject(obj, objects);
        if (resolved is PdfName name) {
            return name.Name;
        }

        if (resolved is PdfArray array && array.Items.Count > 0) {
            var first = ResolveObject(array.Items[0], objects);
            if (first is PdfName firstName) {
                return firstName.Name;
            }
        }

        return string.Empty;
    }

    private static PdfObject? ResolveObject(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        return PdfObjectLookup.Resolve(objects, obj);
    }

    private static bool TryBuildPngFile(
        PdfStream stream,
        int width,
        int height,
        int bitsPerComponent,
        string colorSpace,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] pngBytes) {
        pngBytes = Array.Empty<byte>();
        if (width <= 0 || height <= 0 || bitsPerComponent != 8) {
            return false;
        }

        int colorType;
        if (string.Equals(colorSpace, "DeviceGray", System.StringComparison.Ordinal)) {
            colorType = 0;
        } else if (string.Equals(colorSpace, "DeviceRGB", System.StringComparison.Ordinal)) {
            colorType = 2;
        } else {
            return false;
        }

        if (stream.Dictionary.Items.ContainsKey("SMask")) {
            return TryBuildPngFileWithSoftMask(stream, width, height, bitsPerComponent, colorType, objects, out pngBytes);
        }

        PdfDictionary? decodeParms = null;
        if (stream.Dictionary.Items.TryGetValue("DecodeParms", out var decodeParmsObj)) {
            decodeParms = ResolveDict(decodeParmsObj, objects);
        }

        int predictor = (int)(decodeParms?.Get<PdfNumber>("Predictor")?.Value ?? 1);
        if (predictor < 10 || predictor > 15) {
            return false;
        }

        using var ms = new MemoryStream();
        WritePngSignature(ms);
        WritePngChunk(ms, "IHDR", BuildIhdr(width, height, bitsPerComponent, colorType));
        WritePngChunk(ms, "IDAT", stream.Data);
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        pngBytes = ms.ToArray();
        return true;
    }

    private static bool TryBuildPngFileWithSoftMask(
        PdfStream stream,
        int width,
        int height,
        int bitsPerComponent,
        int colorType,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] pngBytes) {
        pngBytes = Array.Empty<byte>();
        if (!stream.Dictionary.Items.TryGetValue("SMask", out var softMaskObj)) {
            return false;
        }

        PdfStream? softMask = ResolveStream(softMaskObj, objects);
        if (softMask is null) {
            return false;
        }

        int softMaskWidth = (int)(softMask.Dictionary.Get<PdfNumber>("Width")?.Value ?? 0);
        int softMaskHeight = (int)(softMask.Dictionary.Get<PdfNumber>("Height")?.Value ?? 0);
        int softMaskBitsPerComponent = (int)(softMask.Dictionary.Get<PdfNumber>("BitsPerComponent")?.Value ?? 0);
        string softMaskColorSpace = GetNameOrEmpty(softMask.Dictionary.Items.TryGetValue("ColorSpace", out var softMaskColorSpaceObj) ? softMaskColorSpaceObj : null, objects);
        string softMaskFilter = GetFilterName(softMask.Dictionary.Items.TryGetValue("Filter", out var softMaskFilterObj) ? softMaskFilterObj : null, objects);
        if (softMaskWidth != width ||
            softMaskHeight != height ||
            softMaskBitsPerComponent != bitsPerComponent ||
            !string.Equals(softMaskColorSpace, "DeviceGray", System.StringComparison.Ordinal) ||
            !string.Equals(softMaskFilter, "FlateDecode", System.StringComparison.Ordinal)) {
            return false;
        }

        int baseColors;
        int alphaColorType;
        if (colorType == 0) {
            baseColors = 1;
            alphaColorType = 4;
        } else if (colorType == 2) {
            baseColors = 3;
            alphaColorType = 6;
        } else {
            return false;
        }

        byte[] basePixels = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
        byte[] alphaPixels = Filters.StreamDecoder.Decode(softMask.Dictionary, softMask.Data, objects);
        int baseRowLength = width * baseColors;
        int alphaRowLength = width;
        int expectedBaseLength = baseRowLength * height;
        int expectedAlphaLength = alphaRowLength * height;
        if (basePixels.Length < expectedBaseLength || alphaPixels.Length < expectedAlphaLength) {
            return false;
        }

        int outputChannels = baseColors + 1;
        byte[] scanlines = new byte[(1 + width * outputChannels) * height];
        for (int row = 0; row < height; row++) {
            int outputRow = row * (1 + width * outputChannels);
            int baseRow = row * baseRowLength;
            int alphaRow = row * alphaRowLength;
            scanlines[outputRow] = 0;

            for (int pixel = 0; pixel < width; pixel++) {
                int outputPixel = outputRow + 1 + pixel * outputChannels;
                int basePixel = baseRow + pixel * baseColors;
                for (int channel = 0; channel < baseColors; channel++) {
                    scanlines[outputPixel + channel] = basePixels[basePixel + channel];
                }

                scanlines[outputPixel + baseColors] = alphaPixels[alphaRow + pixel];
            }
        }

        using var ms = new MemoryStream();
        WritePngSignature(ms);
        WritePngChunk(ms, "IHDR", BuildIhdr(width, height, bitsPerComponent, alphaColorType));
        WritePngChunk(ms, "IDAT", DeflateZlibStored(scanlines));
        WritePngChunk(ms, "IEND", Array.Empty<byte>());
        pngBytes = ms.ToArray();
        return true;
    }

    private static PdfStream? ResolveStream(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        var resolved = ResolveObject(obj, objects);
        return resolved as PdfStream;
    }

    private static byte[] BuildIhdr(int width, int height, int bitDepth, int colorType) {
        var ihdr = new byte[13];
        WriteInt32BigEndian(ihdr, 0, width);
        WriteInt32BigEndian(ihdr, 4, height);
        ihdr[8] = (byte)bitDepth;
        ihdr[9] = (byte)colorType;
        ihdr[10] = 0;
        ihdr[11] = 0;
        ihdr[12] = 0;
        return ihdr;
    }

    private static void WritePngSignature(Stream stream) {
        byte[] signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
        stream.Write(signature, 0, signature.Length);
    }

    private static void WritePngChunk(Stream stream, string type, byte[] data) {
        byte[] typeBytes = Encoding.ASCII.GetBytes(type);
        var length = new byte[4];
        WriteInt32BigEndian(length, 0, data.Length);
        stream.Write(length, 0, length.Length);
        stream.Write(typeBytes, 0, typeBytes.Length);
        stream.Write(data, 0, data.Length);

        uint crc = ComputeCrc32(typeBytes, data);
        var crcBytes = new byte[4];
        WriteUInt32BigEndian(crcBytes, 0, crc);
        stream.Write(crcBytes, 0, crcBytes.Length);
    }

    private static void WriteInt32BigEndian(byte[] buffer, int offset, int value) {
        buffer[offset] = (byte)((value >> 24) & 0xFF);
        buffer[offset + 1] = (byte)((value >> 16) & 0xFF);
        buffer[offset + 2] = (byte)((value >> 8) & 0xFF);
        buffer[offset + 3] = (byte)(value & 0xFF);
    }

    private static void WriteUInt32BigEndian(byte[] buffer, int offset, uint value) {
        buffer[offset] = (byte)((value >> 24) & 0xFF);
        buffer[offset + 1] = (byte)((value >> 16) & 0xFF);
        buffer[offset + 2] = (byte)((value >> 8) & 0xFF);
        buffer[offset + 3] = (byte)(value & 0xFF);
    }

    private static byte[] DeflateZlibStored(byte[] data) {
        using var ms = new MemoryStream();
        ms.WriteByte(0x78);
        ms.WriteByte(0x01);

        int offset = 0;
        do {
            int blockLength = Math.Min(65535, data.Length - offset);
            bool final = offset + blockLength >= data.Length;
            ms.WriteByte(final ? (byte)1 : (byte)0);
            ms.WriteByte((byte)(blockLength & 0xFF));
            ms.WriteByte((byte)((blockLength >> 8) & 0xFF));
            ushort nlen = (ushort)~blockLength;
            ms.WriteByte((byte)(nlen & 0xFF));
            ms.WriteByte((byte)((nlen >> 8) & 0xFF));
            ms.Write(data, offset, blockLength);
            offset += blockLength;
        } while (offset < data.Length);

        uint adler = Adler32(data);
        ms.WriteByte((byte)((adler >> 24) & 0xFF));
        ms.WriteByte((byte)((adler >> 16) & 0xFF));
        ms.WriteByte((byte)((adler >> 8) & 0xFF));
        ms.WriteByte((byte)(adler & 0xFF));
        return ms.ToArray();
    }

    private static uint Adler32(byte[] data) {
        const uint mod = 65521;
        uint a = 1;
        uint b = 0;
        for (int i = 0; i < data.Length; i++) {
            a = (a + data[i]) % mod;
            b = (b + a) % mod;
        }

        return (b << 16) | a;
    }

    private static uint ComputeCrc32(byte[] typeBytes, byte[] data) {
        uint crc = 0xFFFFFFFF;
        for (int i = 0; i < typeBytes.Length; i++) {
            crc = UpdateCrc32(crc, typeBytes[i]);
        }

        for (int i = 0; i < data.Length; i++) {
            crc = UpdateCrc32(crc, data[i]);
        }

        return crc ^ 0xFFFFFFFF;
    }

    private static uint UpdateCrc32(uint crc, byte value) {
        crc ^= value;
        for (int i = 0; i < 8; i++) {
            if ((crc & 1) != 0) {
                crc = (crc >> 1) ^ 0xEDB88320;
            } else {
                crc >>= 1;
            }
        }

        return crc;
    }
}
