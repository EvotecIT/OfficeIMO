using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    private const string DefaultAppearanceFontName = "Helv";
    private const int MaxInheritedCidWidthEntries = 65536;
    private const int MaxInheritedCidWidthRangeEntries = 4096;

    private static PdfDictionary? TryReadDefaultResources(Dictionary<int, PdfIndirectObject> objects, PdfDictionary dictionary) {
        if (!dictionary.Items.TryGetValue("DR", out PdfObject? defaultResourcesObject)) {
            return null;
        }

        return ResolveDictionary(objects, defaultResourcesObject);
    }

    private static bool TryCreateInheritedTextAppearanceFontPlan(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary? defaultResources,
        PdfDictionary? widgetAppearanceResources,
        PdfDictionary? widgetPageResources,
        string displayValue,
        out TextAppearanceFontPlan? fontPlan) {
        fontPlan = null;
        if (string.IsNullOrEmpty(displayValue)) {
            return false;
        }

        foreach (PdfDictionary resources in EnumerateCandidateTextAppearanceResources(defaultResources, widgetAppearanceResources, widgetPageResources)) {
            if (TryCreateInheritedTextAppearanceFontPlan(objects, resources, displayValue, out fontPlan)) {
                return true;
            }
        }

        return false;
    }

    private static bool TryCreateInheritedTextAppearanceFontPlan(
        Dictionary<int, PdfIndirectObject> objects,
        PdfDictionary? defaultResources,
        string displayValue,
        out TextAppearanceFontPlan? fontPlan) {
        fontPlan = null;
        if (string.IsNullOrEmpty(displayValue) || defaultResources is null) {
            return false;
        }

        if (!defaultResources.Items.TryGetValue("Font", out PdfObject? fontsObject) ||
            ResolveDictionary(objects, fontsObject) is not PdfDictionary fonts) {
            return false;
        }

        if (fonts.Items.TryGetValue(DefaultAppearanceFontName, out PdfObject? defaultFontResourceObject) &&
            TryCreateInheritedTextAppearanceFontPlan(objects, DefaultAppearanceFontName, defaultFontResourceObject, displayValue, out fontPlan)) {
            return true;
        }

        foreach (string fontResourceName in fonts.Items.Keys
            .Where(name => !string.Equals(name, DefaultAppearanceFontName, StringComparison.Ordinal))
            .OrderBy(name => name, StringComparer.Ordinal)) {
            if (TryCreateInheritedTextAppearanceFontPlan(objects, fontResourceName, fonts.Items[fontResourceName], displayValue, out fontPlan)) {
                return true;
            }
        }

        return false;
    }

    private static IEnumerable<PdfDictionary> EnumerateCandidateTextAppearanceResources(params PdfDictionary?[] resources) {
        var seen = new List<PdfDictionary>();
        for (int i = 0; i < resources.Length; i++) {
            PdfDictionary? candidate = resources[i];
            if (candidate is null || seen.Any(item => ReferenceEquals(item, candidate))) {
                continue;
            }

            seen.Add(candidate);
            yield return candidate;
        }
    }

    private static PdfDictionary? TryReadNormalAppearanceResources(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget) {
        if (!TryGetNormalAppearanceObject(objects, widget, out PdfObject? normalAppearance)) {
            return null;
        }

        PdfStream? stream = ResolveObject(objects, normalAppearance) as PdfStream;
        if (stream is null ||
            !stream.Dictionary.Items.TryGetValue("Resources", out PdfObject? resourcesObject)) {
            return null;
        }

        return ResolveDictionary(objects, resourcesObject);
    }

    private static PdfDictionary? TryReadWidgetPageResources(Dictionary<int, PdfIndirectObject> objects, PdfDictionary widget) {
        if (!widget.Items.TryGetValue("P", out PdfObject? pageObject) ||
            ResolveDictionary(objects, pageObject) is not PdfDictionary page) {
            return null;
        }

        return TryReadInheritedPageResources(objects, page);
    }

    private static PdfDictionary? TryReadInheritedPageResources(Dictionary<int, PdfIndirectObject> objects, PdfDictionary page) {
        PdfDictionary? current = page;
        int guard = 0;
        while (current is not null && guard++ < 100) {
            if (current.Items.TryGetValue("Resources", out PdfObject? resourcesObject) &&
                ResolveDictionary(objects, resourcesObject) is PdfDictionary resources) {
                return resources;
            }

            current = current.Items.TryGetValue("Parent", out PdfObject? parentObject) &&
                parentObject is PdfReference parentReference &&
                PdfObjectLookup.TryGet(objects, parentReference, out PdfIndirectObject? parentIndirect) &&
                parentIndirect.Value is PdfDictionary parent
                    ? parent
                    : null;
        }

        return null;
    }

    private static bool TryCreateInheritedTextAppearanceFontPlan(
        Dictionary<int, PdfIndirectObject> objects,
        string fontResourceName,
        PdfObject fontResourceObject,
        string displayValue,
        out TextAppearanceFontPlan? fontPlan) {
        fontPlan = null;
        if (ResolveDictionary(objects, fontResourceObject) is not PdfDictionary font ||
            font.Get<PdfName>("Subtype")?.Name != "Type0" ||
            !font.Items.TryGetValue("ToUnicode", out PdfObject? toUnicodeObject) ||
            ResolveObject(objects, toUnicodeObject) is not PdfStream toUnicodeStream) {
            return false;
        }

        byte[] toUnicodeData = StreamDecoder.Decode(toUnicodeStream.Dictionary, toUnicodeStream.Data, objects);
        if (!ToUnicodeCMap.TryParse(toUnicodeData, out ToUnicodeCMap? cmap) ||
            cmap is null ||
            !TryCreateTextSegmentEncoder(cmap, displayValue, out string? mappedHex, out Func<string, string?>? segmentEncoder)) {
            return false;
        }

        var appearanceFonts = new PdfDictionary();
        appearanceFonts.Items[fontResourceName] = fontResourceObject;

        var resources = new PdfDictionary();
        resources.Items["Font"] = appearanceFonts;
        TryCreateInheritedType0TextMeasure(objects, font, cmap, out Func<string, double, double>? measureTextSegmentWidth);
        fontPlan = new TextAppearanceFontPlan(fontResourceName, resources, mappedHex, segmentEncoder, measureTextSegmentWidth, encodeTextSegments: null, materialize: null);
        return true;
    }

    private static bool TryCreateInheritedType0TextMeasure(Dictionary<int, PdfIndirectObject> objects, PdfDictionary font, ToUnicodeCMap cmap, out Func<string, double, double>? measureTextSegmentWidth) {
        measureTextSegmentWidth = null;
        if (!font.Items.TryGetValue("DescendantFonts", out PdfObject? descendantFontsObject) ||
            ResolveObject(objects, descendantFontsObject) is not PdfArray descendantFonts ||
            descendantFonts.Items.Count == 0 ||
            ResolveDictionary(objects, descendantFonts.Items[0]) is not PdfDictionary descendantFont) {
            return false;
        }

        double defaultWidth = descendantFont.Get<PdfNumber>("DW")?.Value ?? 1000D;
        var widths = new Dictionary<int, double>();
        if (ResolveObject(objects, descendantFont.Items.TryGetValue("W", out PdfObject? widthsObject) ? widthsObject : null) is PdfArray widthArray &&
            !TryReadCidWidths(widthArray, widths)) {
            return false;
        }

        measureTextSegmentWidth = (text, fontSize) => {
            if (!cmap.TryEncodeTextCodes(text, out IReadOnlyList<string> codeHexValues)) {
                throw new InvalidOperationException("The inherited appearance font cannot encode the text segment for measurement.");
            }

            double width = 0D;
            foreach (string codeHex in codeHexValues) {
                int code = Convert.ToInt32(codeHex, 16);
                width += (widths.TryGetValue(code, out double glyphWidth) ? glyphWidth : defaultWidth) * fontSize / 1000D;
            }

            return width;
        };
        return true;
    }

    private static bool TryReadCidWidths(PdfArray widthArray, Dictionary<int, double> widths) {
        for (int index = 0; index < widthArray.Items.Count;) {
            if (widthArray.Items[index++] is not PdfNumber firstNumber) {
                return false;
            }

            int firstCode = (int)firstNumber.Value;
            if (index >= widthArray.Items.Count) {
                return false;
            }

            PdfObject widthSpec = widthArray.Items[index++];
            if (widthSpec is PdfArray explicitWidths) {
                int count = Math.Min(explicitWidths.Items.Count, MaxInheritedCidWidthEntries - widths.Count);
                for (int offset = 0; offset < count; offset++) {
                    if (explicitWidths.Items[offset] is not PdfNumber widthNumber) {
                        return false;
                    }

                    widths[firstCode + offset] = widthNumber.Value;
                }

                if (widths.Count >= MaxInheritedCidWidthEntries) {
                    break;
                }

                continue;
            }

            if (widthSpec is not PdfNumber lastNumber ||
                index >= widthArray.Items.Count ||
                widthArray.Items[index++] is not PdfNumber rangeWidthNumber) {
                return false;
            }

            int lastCode = (int)lastNumber.Value;
            int rangeLength = lastCode >= firstCode ? lastCode - firstCode + 1 : 0;
            if (rangeLength <= 0) {
                continue;
            }

            int rangeCount = Math.Min(rangeLength, MaxInheritedCidWidthRangeEntries);
            rangeCount = Math.Min(rangeCount, MaxInheritedCidWidthEntries - widths.Count);
            for (int offset = 0; offset < rangeCount; offset++) {
                widths[firstCode + offset] = rangeWidthNumber.Value;
            }

            if (widths.Count >= MaxInheritedCidWidthEntries) {
                break;
            }
        }

        return true;
    }

    private static bool TryCreateTextSegmentEncoder(ToUnicodeCMap cmap, string displayValue, out string? encodedTextHex, out Func<string, string?>? encodeTextSegmentHex) {
        encodedTextHex = null;
        encodeTextSegmentHex = segment => cmap.TryEncodeText(segment, out string segmentHex) ? segmentHex : null;
        if (cmap.TryEncodeText(displayValue, out string mappedHex)) {
            encodedTextHex = mappedHex;
            return true;
        }

        string normalized = displayValue.Replace("\r\n", "\n").Replace('\r', '\n');
        string[] lines = normalized.Split('\n');
        if (lines.Length == 1) {
            encodeTextSegmentHex = null;
            return false;
        }

        for (int i = 0; i < lines.Length; i++) {
            if (lines[i].Length > 0 && !cmap.TryEncodeText(lines[i], out _)) {
                encodeTextSegmentHex = null;
                return false;
            }
        }

        return true;
    }

    private sealed class TextAppearanceFontPlan {
        private readonly Action<Dictionary<int, PdfIndirectObject>, TextAppearanceFontPlan>? _materialize;

        public TextAppearanceFontPlan(
            string fontResourceName,
            PdfDictionary resources,
            string? encodedTextHex,
            Func<string, string?>? encodeTextSegmentHex,
            Func<string, double, double>? measureTextSegmentWidth,
            Func<string, IReadOnlyList<PdfTextAppearanceSegment>>? encodeTextSegments,
            Action<Dictionary<int, PdfIndirectObject>, TextAppearanceFontPlan>? materialize) {
            FontResourceName = fontResourceName;
            Resources = resources;
            EncodedTextHex = encodedTextHex;
            EncodeTextSegmentHex = encodeTextSegmentHex;
            MeasureTextSegmentWidth = measureTextSegmentWidth;
            EncodeTextSegments = encodeTextSegments;
            _materialize = materialize;
        }

        public string FontResourceName { get; }

        public PdfDictionary Resources { get; }

        public string? EncodedTextHex { get; }

        public Func<string, string?>? EncodeTextSegmentHex { get; }

        public Func<string, double, double>? MeasureTextSegmentWidth { get; }

        public Func<string, IReadOnlyList<PdfTextAppearanceSegment>>? EncodeTextSegments { get; }

        public void Materialize(Dictionary<int, PdfIndirectObject> objects) {
            _materialize?.Invoke(objects, this);
        }
    }
}
