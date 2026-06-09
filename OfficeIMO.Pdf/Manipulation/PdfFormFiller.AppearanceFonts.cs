using OfficeIMO.Pdf.Filters;

namespace OfficeIMO.Pdf;

public static partial class PdfFormFiller {
    private const string DefaultAppearanceFontName = "Helv";

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
        fontPlan = new TextAppearanceFontPlan(fontResourceName, resources, mappedHex, segmentEncoder, measureTextSegmentWidth: null, encodeTextSegments: null, materialize: null);
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
