using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    private const int MaxCidWidthEntries = 65536;
    private const int MaxCidWidthRangeEntries = 4096;

    public static Dictionary<string, PdfFontResource> GetFontsForPage(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects) {
        var dict = GetInheritedDictionary(page, "Resources", objects);
        return GetFontsForResources(dict, objects);
    }

    public static Dictionary<string, PdfFontResource> GetFontsForResources(PdfDictionary? resources, Dictionary<int, PdfIndirectObject> objects) {
        var fonts = new Dictionary<string, PdfFontResource>(System.StringComparer.Ordinal);
        if (resources is null) return fonts;
        if (!resources.Items.TryGetValue("Font", out var fontDictObj)) return fonts;
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
                    int count = System.Math.Min(list.Items.Count, MaxCidWidthEntries - dict.Count);
                    for (int j = 0; j < count; j++) {
                        if (list.Items[j] is PdfNumber wn) dict[startCid + j] = wn.Value; else dict[startCid + j] = dw;
                    }
                } else if (next is PdfNumber endCidNum) {
                    int endCid = (int)endCidNum.Value; i++;
                    if (i >= wArr.Items.Count) break;
                    var wNum = wArr.Items[i] as PdfNumber; double wv = wNum?.Value ?? dw;
                    int rangeLength = endCid >= startCid ? endCid - startCid + 1 : 0;
                    if (rangeLength <= 0) continue;

                    int count = System.Math.Min(rangeLength, MaxCidWidthRangeEntries);
                    count = System.Math.Min(count, MaxCidWidthEntries - dict.Count);
                    for (int offset = 0; offset < count; offset++) dict[startCid + offset] = wv;
                }

                if (dict.Count >= MaxCidWidthEntries) {
                    break;
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
        foreach (var kv in GetFontsForResources(res, objects)) {
            map[kv.Key] = BuildDecoderForFont(kv.Value);
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

    public static IReadOnlyList<PdfExtractedImage> GetImageXObjectsForPage(PdfDictionary page, Dictionary<int, PdfIndirectObject> objects, int pageNumber, IReadOnlyList<PdfImagePlacement>? imagePlacements = null) {
        var result = new List<PdfExtractedImage>();
        var res = GetInheritedDictionary(page, "Resources", objects);
        if (res is null) return result;
        result.AddRange(GetImageXObjectsForResources(res, objects, pageNumber, imagePlacements));
        return result;
    }

    internal static IReadOnlyList<PdfExtractedImage> GetImageXObjectsForResources(PdfDictionary resources, Dictionary<int, PdfIndirectObject> objects, int pageNumber, IReadOnlyList<PdfImagePlacement>? imagePlacements = null, bool colorizeImageMasks = false, PdfReadLimits? limits = null) {
        var result = new List<PdfExtractedImage>();
        Dictionary<string, List<PdfImagePlacement>>? placedImagesByKey = null;
        Dictionary<string, List<PdfImagePlacement>>? placedImagesByResourceNameWithoutIdentity = null;
        if (imagePlacements is not null) {
            placedImagesByKey = new Dictionary<string, List<PdfImagePlacement>>(System.StringComparer.Ordinal);
            placedImagesByResourceNameWithoutIdentity = new Dictionary<string, List<PdfImagePlacement>>(System.StringComparer.Ordinal);
            for (int i = 0; i < imagePlacements.Count; i++) {
                PdfImagePlacement placement = imagePlacements[i];
                if (!string.IsNullOrEmpty(placement.ResourceName)) {
                    if (placement.ObjectNumber > 0 || placement.DirectStreamIdentity != 0) {
                        AddPlacedImage(placedImagesByKey, BuildImagePlacementKey(placement.ResourceName, placement.ObjectNumber, placement.DirectStreamIdentity), placement);
                    } else {
                        AddPlacedImage(placedImagesByResourceNameWithoutIdentity, placement.ResourceName, placement);
                    }
                }
            }
        }

        PdfReadLimits effectiveLimits = limits ?? PdfReadLimits.Default;
        int traversedObjects = 0;
        CollectImageXObjectsFromResources(resources, objects, pageNumber, result, new HashSet<(PdfStream Stream, PdfDictionary Resources)>(), new HashSet<string>(System.StringComparer.Ordinal), placedImagesByKey, placedImagesByResourceNameWithoutIdentity, colorizeImageMasks, effectiveLimits, depth: 0, ref traversedObjects);
        return result;
    }

    private static void CollectImageXObjectsFromResources(PdfDictionary resources, Dictionary<int, PdfIndirectObject> objects, int pageNumber, List<PdfExtractedImage> result, HashSet<(PdfStream Stream, PdfDictionary Resources)> visitedFormContexts, HashSet<string> addedImageKeys, Dictionary<string, List<PdfImagePlacement>>? placedImagesByKey, Dictionary<string, List<PdfImagePlacement>>? placedImagesByResourceNameWithoutIdentity, bool colorizeImageMasks, PdfReadLimits limits, int depth, ref int traversedObjects) {
        if (depth > limits.MaxContentNestingDepth) {
            throw PdfReadLimitException.Create(PdfReadLimitKind.ContentNestingDepth, limits.MaxContentNestingDepth, depth);
        }

        if (!resources.Items.TryGetValue("XObject", out var xoObj)) return;
        var xo = ResolveDict(xoObj, objects);
        if (xo is null) return;

        foreach (var kv in xo.Items) {
            traversedObjects++;
            if (traversedObjects > limits.MaxContentOperands) {
                throw PdfReadLimitException.Create(PdfReadLimitKind.ContentOperands, limits.MaxContentOperands, traversedObjects);
            }

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
            if (string.Equals(subtype, "Image", System.StringComparison.Ordinal)) {
                int directStreamIdentity = objectNumber == 0
                    ? System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(stream)
                    : 0;
                IReadOnlyList<PdfImagePlacement>? matchingPlacements = GetPlacedImageMatches(kv.Key, objectNumber, directStreamIdentity, placedImagesByKey, placedImagesByResourceNameWithoutIdentity);
                if (matchingPlacements is { Count: 0 }) {
                    continue;
                }

                bool isImageMask = PdfImageMaskNormalizer.IsImageMask(stream, objects);
                if (isImageMask && matchingPlacements is not null) {
                    for (int placementIndex = 0; placementIndex < matchingPlacements.Count; placementIndex++) {
                        OfficeColor imageMaskColor = matchingPlacements[placementIndex].ImageMaskColor;
                        if (!addedImageKeys.Add(BuildImageResourceKey(pageNumber, kv.Key, objectNumber, directStreamIdentity, imageMaskColor))) {
                            continue;
                        }

                        result.Add(BuildExtractedImage(pageNumber, kv.Key, objectNumber, directStreamIdentity, stream, objects, imageMaskColor, resources, colorizeImageMasks));
                    }
                } else {
                    if (!addedImageKeys.Add(BuildImageResourceKey(pageNumber, kv.Key, objectNumber, directStreamIdentity))) {
                        continue;
                    }

                    result.Add(BuildExtractedImage(pageNumber, kv.Key, objectNumber, directStreamIdentity, stream, objects, resources: resources));
                }

                continue;
            }

            if (!string.Equals(subtype, "Form", System.StringComparison.Ordinal)) {
                continue;
            }

            PdfDictionary? formResources = null;
            if (stream.Dictionary.Items.TryGetValue("Resources", out var formResourcesObj)) {
                formResources = ResolveDict(formResourcesObj, objects);
            }

            formResources ??= resources;
            if (!visitedFormContexts.Add((stream, formResources))) {
                continue;
            }

            CollectImageXObjectsFromResources(formResources, objects, pageNumber, result, visitedFormContexts, addedImageKeys, placedImagesByKey, placedImagesByResourceNameWithoutIdentity, colorizeImageMasks, limits, depth + 1, ref traversedObjects);
        }
    }

    private static void AddPlacedImage(Dictionary<string, List<PdfImagePlacement>> placedImages, string key, PdfImagePlacement placement) {
        if (!placedImages.TryGetValue(key, out List<PdfImagePlacement>? placements)) {
            placements = new List<PdfImagePlacement>();
            placedImages[key] = placements;
        }

        placements.Add(placement);
    }

    private static IReadOnlyList<PdfImagePlacement>? GetPlacedImageMatches(string resourceName, int objectNumber, int directStreamIdentity, Dictionary<string, List<PdfImagePlacement>>? placedImagesByKey, Dictionary<string, List<PdfImagePlacement>>? placedImagesByResourceNameWithoutIdentity) {
        if (placedImagesByKey is null && placedImagesByResourceNameWithoutIdentity is null) {
            return null;
        }

        if (objectNumber > 0 || directStreamIdentity != 0) {
            return placedImagesByKey != null &&
                placedImagesByKey.TryGetValue(BuildImagePlacementKey(resourceName, objectNumber, directStreamIdentity), out List<PdfImagePlacement>? placements)
                    ? placements
                    : Array.Empty<PdfImagePlacement>();
        }

        return placedImagesByResourceNameWithoutIdentity != null &&
            placedImagesByResourceNameWithoutIdentity.TryGetValue(resourceName, out List<PdfImagePlacement>? namePlacements)
                ? namePlacements
                : Array.Empty<PdfImagePlacement>();
    }

    private static string BuildImagePlacementKey(string resourceName, int objectNumber, int directStreamIdentity) {
        return resourceName +
            "|" +
            objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            "|" +
            directStreamIdentity.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private static string BuildImageResourceKey(int pageNumber, string resourceName, int objectNumber, int directStreamIdentity) {
        return pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            "|" +
            resourceName +
            "|" +
            objectNumber.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            "|" +
            directStreamIdentity.ToString(System.Globalization.CultureInfo.InvariantCulture);
    }

    private static string BuildImageResourceKey(int pageNumber, string resourceName, int objectNumber, int directStreamIdentity, OfficeColor imageMaskColor) =>
        BuildImageResourceKey(pageNumber, resourceName, objectNumber, directStreamIdentity) +
        "|" +
        imageMaskColor.R.ToString(System.Globalization.CultureInfo.InvariantCulture) +
        "," +
        imageMaskColor.G.ToString(System.Globalization.CultureInfo.InvariantCulture) +
        "," +
        imageMaskColor.B.ToString(System.Globalization.CultureInfo.InvariantCulture) +
        "," +
        imageMaskColor.A.ToString(System.Globalization.CultureInfo.InvariantCulture);

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

        byte[]? embeddedTrueTypeFont = TryReadEmbeddedTrueTypeFont(fontVal, objects);
        return new PdfFontResource(resourceName, baseFont, encoding, hasToUnicode, cmap, differences, embeddedTrueTypeFont);
    }

    private static byte[]? TryReadEmbeddedTrueTypeFont(PdfDictionary font, Dictionary<int, PdfIndirectObject> objects) {
        PdfDictionary fontWithDescriptor = font;
        if (string.Equals(font.Get<PdfName>("Subtype")?.Name, "Type0", System.StringComparison.Ordinal) &&
            font.Items.TryGetValue("DescendantFonts", out PdfObject? descendantsObject)) {
            PdfArray? descendants = ResolveArray(descendantsObject, objects);
            PdfDictionary? descendant = descendants is { Items.Count: > 0 }
                ? ResolveDict(descendants.Items[0], objects)
                : null;
            if (descendant != null) fontWithDescriptor = descendant;
        }

        PdfDictionary? descriptor = fontWithDescriptor.Items.TryGetValue("FontDescriptor", out PdfObject? descriptorObject)
            ? ResolveDict(descriptorObject, objects)
            : null;
        if (descriptor == null) return null;

        PdfStream? program = ResolveObject(
            descriptor.Items.TryGetValue("FontFile2", out PdfObject? fontFile2) ? fontFile2 : null,
            objects) as PdfStream;
        if (program == null && descriptor.Items.TryGetValue("FontFile3", out PdfObject? fontFile3)) {
            PdfStream? candidate = ResolveObject(fontFile3, objects) as PdfStream;
            string? subtype = candidate?.Dictionary.Get<PdfName>("Subtype")?.Name;
            if (candidate != null && (string.Equals(subtype, "OpenType", System.StringComparison.Ordinal) || string.Equals(subtype, "TrueType", System.StringComparison.Ordinal))) {
                program = candidate;
            }
        }
        if (program == null || Filters.StreamDecoder.GetUnsupportedFilters(program.Dictionary, objects).Count != 0) return null;

        byte[] bytes;
        try {
            bytes = Filters.StreamDecoder.Decode(program.Dictionary, program.Data, objects);
        } catch (InvalidDataException) {
            return null;
        } catch (NotSupportedException) {
            return null;
        }
        return OfficeTrueTypeFont.TryLoad(bytes) == null ? null : bytes;
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

    internal static PdfExtractedImage BuildExtractedImage(
        int pageNumber,
        string resourceName,
        int objectNumber,
        int directStreamIdentity,
        PdfStream stream,
        Dictionary<int, PdfIndirectObject> objects,
        OfficeColor? imageMaskColor = null,
        PdfDictionary? resources = null,
        bool colorizeImageMask = false) {
        int width = (int)(stream.Dictionary.Get<PdfNumber>("Width")?.Value ?? 0);
        int height = (int)(stream.Dictionary.Get<PdfNumber>("Height")?.Value ?? 0);
        int bitsPerComponent = (int)(stream.Dictionary.Get<PdfNumber>("BitsPerComponent")?.Value ?? 0);
        bool isImageMask = PdfImageMaskNormalizer.IsImageMask(stream, objects);
        if (isImageMask && bitsPerComponent == 0) {
            bitsPerComponent = 1;
        }

        PdfObject? colorSpaceObject = stream.Dictionary.Items.TryGetValue("ColorSpace", out var colorSpaceObj) ? colorSpaceObj : null;
        PdfObject? resolvedColorSpaceObject = ResolveColorSpaceResource(colorSpaceObject, resources, objects);
        PdfObject? effectiveColorSpaceObject = resolvedColorSpaceObject ?? colorSpaceObject;
        string colorSpace = isImageMask ? "ImageMask" : GetNameOrEmpty(effectiveColorSpaceObject, objects);
        string filter = GetFilterName(stream.Dictionary.Items.TryGetValue("Filter", out var filterObj) ? filterObj : null, objects);

        byte[] bytes = stream.Data;
        string? extension = null;
        string? mimeType = null;
        bool isImageFile = false;
        string? transparencyMaskKind = GetTransparencyMaskKind(stream.Dictionary, objects);
        bool transparencyMaskResolved = false;

        if (string.Equals(filter, "DCTDecode", System.StringComparison.Ordinal)) {
            extension = "jpg";
            mimeType = OfficeImageInfo.GetMimeType(OfficeImageFormat.Jpeg);
            isImageFile = true;
        } else if (isImageMask && TryBuildExtractedImageMaskPng(stream, width, height, bitsPerComponent, objects, colorizeImageMask ? imageMaskColor : null, out var imageMaskPngBytes)) {
            bytes = imageMaskPngBytes;
            extension = "png";
            mimeType = OfficeImageInfo.GetMimeType(OfficeImageFormat.Png);
            isImageFile = true;
        } else if (TryBuildPngFile(stream, width, height, bitsPerComponent, effectiveColorSpaceObject, colorSpace, filter, objects, out var pngBytes)) {
            bytes = pngBytes;
            extension = "png";
            mimeType = OfficeImageInfo.GetMimeType(OfficeImageFormat.Png);
            isImageFile = true;
            transparencyMaskResolved = IsTransparencyMaskResolvedByPngNormalization(transparencyMaskKind);
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
            isImageFile,
            transparencyMaskKind,
            transparencyMaskResolved,
            directStreamIdentity,
            isImageMask,
            imageMaskColor ?? OfficeColor.Black);
    }

    private static string? GetTransparencyMaskKind(PdfDictionary dictionary, Dictionary<int, PdfIndirectObject> objects) {
        if (dictionary.Items.TryGetValue("SMask", out var softMaskObj)) {
            var resolvedSoftMask = ResolveObject(softMaskObj, objects);
            if (resolvedSoftMask is PdfName softMaskName &&
                string.Equals(softMaskName.Name, "None", System.StringComparison.Ordinal)) {
                return null;
            }

            return "soft-mask";
        }

        if (!dictionary.Items.TryGetValue("Mask", out var maskObj)) {
            return null;
        }

        var resolvedMask = ResolveObject(maskObj, objects);
        if (resolvedMask is PdfArray) {
            return "color-key-mask";
        }

        if (resolvedMask is PdfStream) {
            return "explicit-mask-image";
        }

        if (resolvedMask is PdfName maskName &&
            string.Equals(maskName.Name, "None", System.StringComparison.Ordinal)) {
            return null;
        }

        return "mask";
    }

    private static bool IsTransparencyMaskResolvedByPngNormalization(string? transparencyMaskKind) =>
        string.Equals(transparencyMaskKind, "soft-mask", System.StringComparison.Ordinal) ||
        string.Equals(transparencyMaskKind, "color-key-mask", System.StringComparison.Ordinal);

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

    private static PdfObject? ResolveColorSpaceResource(PdfObject? colorSpaceObject, PdfDictionary? resources, Dictionary<int, PdfIndirectObject> objects) {
        PdfObject? resolved = ResolveObject(colorSpaceObject, objects);
        if (resolved is not PdfName name || resources == null) {
            return resolved;
        }

        if (!resources.Items.TryGetValue("ColorSpace", out PdfObject? colorSpaceResourcesObject)) {
            return resolved;
        }

        PdfDictionary? colorSpaceResources = ResolveDict(colorSpaceResourcesObject, objects);
        if (colorSpaceResources == null ||
            !colorSpaceResources.Items.TryGetValue(name.Name, out PdfObject? resourceColorSpaceObject)) {
            return resolved;
        }

        return ResolveObject(resourceColorSpaceObject, objects) ?? resourceColorSpaceObject;
    }

    private static PdfObject? ResolveObject(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        return PdfObjectLookup.Resolve(objects, obj);
    }

    private static bool TryBuildPngFile(
        PdfStream stream,
        int width,
        int height,
        int bitsPerComponent,
        PdfObject? colorSpaceObj,
        string colorSpace,
        string filter,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] pngBytes) {
        pngBytes = Array.Empty<byte>();
        if (width <= 0 || height <= 0) {
            return false;
        }

        if (PdfIndexedImageNormalizer.TryBuildPngFile(colorSpaceObj, width, height, bitsPerComponent, stream, objects, out pngBytes)) {
            return true;
        }

        if (bitsPerComponent != 8) {
            return false;
        }

        if (!PdfImageColorSpaceNormalization.TryResolve(colorSpaceObj, colorSpace, objects, out var colorNormalization)) {
            return false;
        }

        if (HasSoftMask(stream.Dictionary, objects)) {
            if (colorNormalization.SourceColorCount != 1 &&
                colorNormalization.SourceColorCount != 3 &&
                colorNormalization.SourceColorCount != 4) {
                return false;
            }

            var decodeTransform = PdfImageDecodeTransform.CreateColor(stream.Dictionary, colorNormalization.SourceColorCount, objects);
            return TryBuildPngFileWithSoftMask(
                stream,
                width,
                height,
                bitsPerComponent,
                colorNormalization.SourceColorCount,
                colorNormalization.PngColorType,
                decodeTransform,
                objects,
                out pngBytes);
        }

        var colorDecodeTransform = PdfImageDecodeTransform.CreateColor(stream.Dictionary, colorNormalization.SourceColorCount, objects);
        var colorKeyMask = PdfImageColorKeyMask.Create(stream.Dictionary, colorNormalization.SourceColorCount, objects);
        if (colorKeyMask is not null) {
            if (Filters.StreamDecoder.GetUnsupportedFilters(stream.Dictionary, objects).Count != 0) {
                return false;
            }

            byte[] pixels = string.IsNullOrEmpty(filter)
                ? stream.Data
                : Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
            return TryBuildPngFileFromDecodedPixelsWithColorKeyMask(
                width,
                height,
                bitsPerComponent,
                colorNormalization.SourceColorCount,
                colorNormalization.PngColorType,
                colorDecodeTransform,
                colorKeyMask,
                pixels,
                out pngBytes);
        }

        if (string.IsNullOrEmpty(filter)) {
            return TryBuildPngFileFromDecodedPixels(width, height, bitsPerComponent, colorNormalization.SourceColorCount, colorNormalization.PngColorType, colorDecodeTransform, stream.Data, out pngBytes);
        }

        if (!string.Equals(filter, "FlateDecode", System.StringComparison.Ordinal)) {
            return TryBuildPngFileFromSupportedDecodedStream(stream, width, height, bitsPerComponent, colorNormalization.SourceColorCount, colorNormalization.PngColorType, colorDecodeTransform, objects, out pngBytes);
        }

        PdfDictionary? decodeParms = null;
        if (stream.Dictionary.Items.TryGetValue("DecodeParms", out var decodeParmsObj)) {
            decodeParms = ResolveDict(decodeParmsObj, objects);
        }

        int predictor = (int)(decodeParms?.Get<PdfNumber>("Predictor")?.Value ?? 1);
        if (predictor <= 1 || predictor == 2) {
            byte[] pixels = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
            return TryBuildPngFileFromDecodedPixels(width, height, bitsPerComponent, colorNormalization.SourceColorCount, colorNormalization.PngColorType, colorDecodeTransform, pixels, out pngBytes);
        }

        if (predictor < 10 || predictor > 15) {
            return false;
        }

        if ((colorNormalization.SourceColorCount != 1 && colorNormalization.SourceColorCount != 3) ||
            colorDecodeTransform is not null ||
            !CanWrapPngPredictorScanlines(decodeParms, width, bitsPerComponent, colorNormalization.SourceColorCount)) {
            byte[] pixels = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
            return TryBuildPngFileFromDecodedPixels(width, height, bitsPerComponent, colorNormalization.SourceColorCount, colorNormalization.PngColorType, colorDecodeTransform, pixels, out pngBytes);
        }

        pngBytes = OfficePngWriter.CreateFromCompressedScanlines(width, height, bitsPerComponent, colorNormalization.PngColorType, stream.Data);
        return true;
    }

    private static bool CanWrapPngPredictorScanlines(PdfDictionary? decodeParms, int width, int bitsPerComponent, int sourceColorCount) {
        if (decodeParms is null) {
            return false;
        }

        int columns = (int)(decodeParms.Get<PdfNumber>("Columns")?.Value ?? 1);
        int colors = (int)(decodeParms.Get<PdfNumber>("Colors")?.Value ?? 1);
        int decodeBitsPerComponent = (int)(decodeParms.Get<PdfNumber>("BitsPerComponent")?.Value ?? 8);
        return columns == width &&
               colors == sourceColorCount &&
               decodeBitsPerComponent == bitsPerComponent;
    }

    private static bool TryBuildPngFileFromSupportedDecodedStream(
        PdfStream stream,
        int width,
        int height,
        int bitsPerComponent,
        int sourceColorCount,
        int pngColorType,
        PdfImageDecodeTransform? decodeTransform,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] pngBytes) {
        pngBytes = Array.Empty<byte>();
        if (Filters.StreamDecoder.GetUnsupportedFilters(stream.Dictionary, objects).Count != 0) {
            return false;
        }

        byte[] pixels = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
        return TryBuildPngFileFromDecodedPixels(width, height, bitsPerComponent, sourceColorCount, pngColorType, decodeTransform, pixels, out pngBytes);
    }

    private static bool TryBuildPngFileFromDecodedPixels(
        int width,
        int height,
        int bitsPerComponent,
        int sourceColorCount,
        int pngColorType,
        PdfImageDecodeTransform? decodeTransform,
        byte[] pixels,
        out byte[] pngBytes) {
        pngBytes = Array.Empty<byte>();
        if (pixels.Length == 0) {
            return false;
        }

        if ((sourceColorCount != 1 && sourceColorCount != 3 && sourceColorCount != 4) ||
            (pngColorType != 0 && pngColorType != 2)) {
            return false;
        }

        long sourceRowLengthLong = (long)width * sourceColorCount;
        long expectedLengthLong = sourceRowLengthLong * height;
        long outputRowLengthLong = (long)width * (pngColorType == 0 ? 1 : 3);
        if (sourceRowLengthLong > int.MaxValue ||
            expectedLengthLong > int.MaxValue ||
            outputRowLengthLong > int.MaxValue) {
            return false;
        }

        int sourceRowLength = (int)sourceRowLengthLong;
        int expectedLength = (int)expectedLengthLong;
        int outputRowLength = (int)outputRowLengthLong;
        if (pixels.Length < expectedLength) {
            return false;
        }

        byte[] scanlines = new byte[(1 + outputRowLength) * height];
        for (int row = 0; row < height; row++) {
            int outputRow = row * (1 + outputRowLength);
            int sourceRow = row * sourceRowLength;
            scanlines[outputRow] = 0;
            if (sourceColorCount == 4) {
                CopyDeviceCmykRowAsRgb(pixels, sourceRow, scanlines, outputRow + 1, width, decodeTransform);
            } else if (decodeTransform is not null) {
                CopyDecodedColorRow(pixels, sourceRow, scanlines, outputRow + 1, width, sourceColorCount, decodeTransform);
            } else {
                Buffer.BlockCopy(pixels, sourceRow, scanlines, outputRow + 1, outputRowLength);
            }
        }

        pngBytes = OfficePngWriter.EncodeScanlines(
            width,
            height,
            bitsPerComponent,
            pngColorType,
            scanlines,
            OfficePngCompression.Stored);
        return true;
    }

    private static bool TryBuildPngFileFromDecodedPixelsWithColorKeyMask(
        int width,
        int height,
        int bitsPerComponent,
        int sourceColorCount,
        int pngColorType,
        PdfImageDecodeTransform? decodeTransform,
        PdfImageColorKeyMask colorKeyMask,
        byte[] pixels,
        out byte[] pngBytes) {
        pngBytes = Array.Empty<byte>();
        if (pixels.Length == 0) {
            return false;
        }

        if ((sourceColorCount != 1 && sourceColorCount != 3 && sourceColorCount != 4) ||
            (pngColorType != 0 && pngColorType != 2)) {
            return false;
        }

        int outputBaseColors = pngColorType == 0 ? 1 : 3;
        int alphaColorType = pngColorType == 0 ? 4 : 6;
        long sourceRowLengthLong = (long)width * sourceColorCount;
        long expectedLengthLong = sourceRowLengthLong * height;
        long outputRowLengthLong = (long)width * (outputBaseColors + 1);
        if (sourceRowLengthLong > int.MaxValue ||
            expectedLengthLong > int.MaxValue ||
            outputRowLengthLong > int.MaxValue) {
            return false;
        }

        int sourceRowLength = (int)sourceRowLengthLong;
        int expectedLength = (int)expectedLengthLong;
        int outputRowLength = (int)outputRowLengthLong;
        if (pixels.Length < expectedLength) {
            return false;
        }

        int outputChannels = outputBaseColors + 1;
        byte[] scanlines = new byte[(1 + outputRowLength) * height];
        for (int row = 0; row < height; row++) {
            int outputRow = row * (1 + outputRowLength);
            int sourceRow = row * sourceRowLength;
            scanlines[outputRow] = 0;

            for (int pixel = 0; pixel < width; pixel++) {
                int sourcePixel = sourceRow + pixel * sourceColorCount;
                int outputPixel = outputRow + 1 + pixel * outputChannels;
                if (sourceColorCount == 4) {
                    CopyDeviceCmykRowAsRgb(pixels, sourcePixel, scanlines, outputPixel, 1, decodeTransform);
                } else {
                    for (int channel = 0; channel < outputBaseColors; channel++) {
                        scanlines[outputPixel + channel] = TransformColorComponent(pixels[sourcePixel + channel], channel, decodeTransform);
                    }
                }

                scanlines[outputPixel + outputBaseColors] = colorKeyMask.IsTransparent(pixels, sourcePixel)
                    ? (byte)0
                    : (byte)255;
            }
        }

        pngBytes = OfficePngWriter.EncodeScanlines(
            width,
            height,
            bitsPerComponent,
            alphaColorType,
            scanlines,
            OfficePngCompression.Stored);
        return true;
    }

    private static void CopyDecodedColorRow(
        byte[] source,
        int sourceOffset,
        byte[] target,
        int targetOffset,
        int width,
        int sourceColorCount,
        PdfImageDecodeTransform decodeTransform) {
        int rowLength = width * sourceColorCount;
        for (int channel = 0; channel < rowLength; channel++) {
            target[targetOffset + channel] = decodeTransform.TransformColorComponent(source[sourceOffset + channel], channel % sourceColorCount);
        }
    }

    private static void CopyDeviceCmykRowAsRgb(byte[] source, int sourceOffset, byte[] target, int targetOffset, int width, PdfImageDecodeTransform? decodeTransform) {
        for (int pixel = 0; pixel < width; pixel++) {
            int sourcePixel = sourceOffset + pixel * 4;
            int targetPixel = targetOffset + pixel * 3;
            byte c = TransformColorComponent(source[sourcePixel], 0, decodeTransform);
            byte m = TransformColorComponent(source[sourcePixel + 1], 1, decodeTransform);
            byte y = TransformColorComponent(source[sourcePixel + 2], 2, decodeTransform);
            byte k = TransformColorComponent(source[sourcePixel + 3], 3, decodeTransform);

            target[targetPixel] = ConvertDeviceCmykComponentToRgb(c, k);
            target[targetPixel + 1] = ConvertDeviceCmykComponentToRgb(m, k);
            target[targetPixel + 2] = ConvertDeviceCmykComponentToRgb(y, k);
        }
    }

    private static byte TransformColorComponent(byte sample, int componentIndex, PdfImageDecodeTransform? decodeTransform) {
        return decodeTransform is null ? sample : decodeTransform.TransformColorComponent(sample, componentIndex);
    }

    private static byte ConvertDeviceCmykComponentToRgb(byte colorant, byte black) {
        int ink = colorant + black;
        return (byte)(255 - (ink > 255 ? 255 : ink));
    }

    private static bool HasSoftMask(PdfDictionary dictionary, Dictionary<int, PdfIndirectObject> objects) {
        if (!dictionary.Items.TryGetValue("SMask", out var softMaskObj)) {
            return false;
        }

        return ResolveObject(softMaskObj, objects) is not PdfName softMaskName ||
               !string.Equals(softMaskName.Name, "None", System.StringComparison.Ordinal);
    }

    private static bool TryBuildPngFileWithSoftMask(
        PdfStream stream,
        int width,
        int height,
        int bitsPerComponent,
        int sourceColorCount,
        int pngColorType,
        PdfImageDecodeTransform? decodeTransform,
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
        if (softMaskWidth != width ||
            softMaskHeight != height ||
            softMaskBitsPerComponent != bitsPerComponent ||
            !string.Equals(softMaskColorSpace, "DeviceGray", System.StringComparison.Ordinal) ||
            Filters.StreamDecoder.GetUnsupportedFilters(softMask.Dictionary, objects).Count != 0) {
            return false;
        }

        int outputBaseColors;
        int alphaColorType;
        if (sourceColorCount == 1 && pngColorType == 0) {
            outputBaseColors = 1;
            alphaColorType = 4;
        } else if (sourceColorCount == 3 && pngColorType == 2) {
            outputBaseColors = 3;
            alphaColorType = 6;
        } else if (sourceColorCount == 4 && pngColorType == 2) {
            outputBaseColors = 3;
            alphaColorType = 6;
        } else {
            return false;
        }

        byte[] basePixels = Filters.StreamDecoder.Decode(stream.Dictionary, stream.Data, objects);
        byte[] alphaPixels = Filters.StreamDecoder.Decode(softMask.Dictionary, softMask.Data, objects);
        long baseRowLengthLong = (long)width * sourceColorCount;
        long alphaRowLengthLong = width;
        long expectedBaseLengthLong = baseRowLengthLong * height;
        long expectedAlphaLengthLong = alphaRowLengthLong * height;
        long outputRowLengthLong = (long)width * (outputBaseColors + 1);
        if (baseRowLengthLong > int.MaxValue ||
            expectedBaseLengthLong > int.MaxValue ||
            expectedAlphaLengthLong > int.MaxValue ||
            outputRowLengthLong > int.MaxValue) {
            return false;
        }

        int baseRowLength = (int)baseRowLengthLong;
        int alphaRowLength = (int)alphaRowLengthLong;
        int expectedBaseLength = (int)expectedBaseLengthLong;
        int expectedAlphaLength = (int)expectedAlphaLengthLong;
        int outputRowLength = (int)outputRowLengthLong;
        if (basePixels.Length < expectedBaseLength || alphaPixels.Length < expectedAlphaLength) {
            return false;
        }

        int outputChannels = outputBaseColors + 1;
        var alphaDecodeTransform = PdfImageDecodeTransform.CreateColor(softMask.Dictionary, 1, objects);
        byte[] scanlines = new byte[(1 + outputRowLength) * height];
        for (int row = 0; row < height; row++) {
            int outputRow = row * (1 + outputRowLength);
            int baseRow = row * baseRowLength;
            int alphaRow = row * alphaRowLength;
            scanlines[outputRow] = 0;

            for (int pixel = 0; pixel < width; pixel++) {
                int outputPixel = outputRow + 1 + pixel * outputChannels;
                int basePixel = baseRow + pixel * sourceColorCount;
                if (sourceColorCount == 4) {
                    CopyDeviceCmykRowAsRgb(basePixels, basePixel, scanlines, outputPixel, 1, decodeTransform);
                } else {
                    for (int channel = 0; channel < outputBaseColors; channel++) {
                        scanlines[outputPixel + channel] = TransformColorComponent(basePixels[basePixel + channel], channel, decodeTransform);
                    }
                }

                scanlines[outputPixel + outputBaseColors] = TransformColorComponent(alphaPixels[alphaRow + pixel], 0, alphaDecodeTransform);
            }
        }

        pngBytes = OfficePngWriter.EncodeScanlines(
            width,
            height,
            bitsPerComponent,
            alphaColorType,
            scanlines,
            OfficePngCompression.Stored);
        return true;
    }

    private static bool TryDecodeSoftMask(
        PdfStream stream,
        int width,
        int height,
        int bitsPerComponent,
        Dictionary<int, PdfIndirectObject> objects,
        out byte[] alphaPixels) {
        alphaPixels = Array.Empty<byte>();
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
        if (softMaskWidth != width ||
            softMaskHeight != height ||
            softMaskBitsPerComponent != bitsPerComponent ||
            !string.Equals(softMaskColorSpace, "DeviceGray", System.StringComparison.Ordinal) ||
            Filters.StreamDecoder.GetUnsupportedFilters(softMask.Dictionary, objects).Count != 0) {
            return false;
        }

        alphaPixels = Filters.StreamDecoder.Decode(softMask.Dictionary, softMask.Data, objects);
        ApplySoftMaskDecode(softMask.Dictionary, alphaPixels, objects);
        return true;
    }

    private static void ApplySoftMaskDecode(PdfDictionary softMaskDictionary, byte[] alphaPixels, Dictionary<int, PdfIndirectObject> objects) {
        if (alphaPixels.Length == 0 ||
            !softMaskDictionary.Items.ContainsKey("Decode")) {
            return;
        }

        for (int i = 0; i < alphaPixels.Length; i++) {
            alphaPixels[i] = DecodeImageComponent(softMaskDictionary, 0, alphaPixels[i], objects);
        }
    }

    private static byte ConvertCmykComponent(byte component, byte black) {
        double value = (1D - component / 255D) * (1D - black / 255D);
        return (byte)System.Math.Round(value * 255D);
    }

    private static PdfStream? ResolveStream(PdfObject? obj, Dictionary<int, PdfIndirectObject> objects) {
        var resolved = ResolveObject(obj, objects);
        return resolved as PdfStream;
    }

}
