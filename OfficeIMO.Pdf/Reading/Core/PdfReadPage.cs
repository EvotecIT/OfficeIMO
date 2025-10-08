namespace OfficeIMO.Pdf;

/// <summary>
/// Represents a single page parsed from the PDF.
/// Provides access to plain text and basic text spans based on content stream operators.
/// </summary>
public sealed class PdfReadPage {
    private readonly PdfDictionary _pageDict;
    private readonly Dictionary<int, PdfIndirectObject> _objects;

    internal PdfReadPage(int objectNumber, PdfDictionary pageDict, Dictionary<int, PdfIndirectObject> objects) {
        ObjectNumber = objectNumber; _pageDict = pageDict; _objects = objects;
    }

    /// <summary>Underlying object number for the page.</summary>
    public int ObjectNumber { get; }

    /// <summary>Extracts plain text from this page without column reordering.</summary>
    public string ExtractText() {
        var spans = GetTextSpans();
        var opts = new TextLayoutEngine.Options { ForceSingleColumn = true };
        var lines = TextLayoutEngine.BuildLines(spans, opts);
        return TextLayoutEngine.EmitText(lines, TextLayoutEngine.DetectColumns(lines, GetPageSize().Width, opts), null);
    }

    /// <summary>
    /// Attempts to read page size from MediaBox (or CropBox) and returns width/height in points.
    /// Falls back to 612x792 (US Letter) when not present or malformed.
    /// </summary>
    public (double Width, double Height) GetPageSize() {
        static (double, double) ParseBox(PdfObject? box) {
            if (box is PdfArray arr && arr.Items.Count >= 4 &&
                arr.Items[0] is PdfNumber llx && arr.Items[1] is PdfNumber lly &&
                arr.Items[2] is PdfNumber urx && arr.Items[3] is PdfNumber ury) {
                double w = urx.Value - llx.Value;
                double h = ury.Value - lly.Value;
                if (w > 0 && h > 0) return (w, h);
            }
            return (612, 792); // default Letter
        }
        if (_pageDict.Items.TryGetValue("MediaBox", out var media)) return ParseBox(media);
        if (_pageDict.Items.TryGetValue("CropBox", out var crop)) return ParseBox(crop);
        return (612, 792);
    }

    /// <summary>Gets text spans (text with position and font info) from this page.</summary>
    public IReadOnlyList<PdfTextSpan> GetTextSpans() {
        var spans = new List<PdfTextSpan>();
        var streams = GetContentStreams();
        var decoders = ResourceResolver.GetFontDecoders(_pageDict, _objects);
        var widthProviders = ResourceResolver.GetFontWidthProviders(_pageDict, _objects);
        string DecodeWithFont(string fontRes, byte[] bytes) =>
            decoders.TryGetValue(fontRes, out var dec) ? dec(bytes) : PdfWinAnsiEncoding.Decode(bytes);
        double SumWidth1000(string fontRes, byte[] bytes) => widthProviders.TryGetValue(fontRes, out var wp) ? wp(bytes) : (bytes?.Length ?? 0) * 500.0;
        foreach (var s in streams) {
            var content = PdfEncoding.Latin1GetString(s);
            spans.AddRange(TextContentParser.Parse(content, DecodeWithFont, SumWidth1000));
        }
        // Additionally parse Form XObjects referenced via /Resources/XObject
        var formStreams = ResourceResolver.GetFormXObjectStreams(_pageDict, _objects);
        foreach (var kv in formStreams) {
            var formName = kv.Key;
            var bytes = kv.Value;
            var content = PdfEncoding.Latin1GetString(bytes);
            // Build decoders using the form's own resources if present
            var formDict = GetFormDict(formName);
            var formDecoders = ResourceResolver.GetFontDecodersForForm(formDict, _objects);
            var formWidths = ResourceResolver.GetFontWidthProviders(formDict, _objects);
            string DecodeWithFormFont(string fontRes, byte[] input) => formDecoders.TryGetValue(fontRes, out var dec) ? dec(input) : DecodeWithFont(fontRes, input);
            double SumWidth1000Form(string fontRes, byte[] input) => formWidths.TryGetValue(fontRes, out var wp) ? wp(input) : SumWidth1000(fontRes, input);
            // If the form has a /Matrix, inject it as a cm operator to apply CTM correctly
            if (formDict is not null && formDict.Items.TryGetValue("Matrix", out var mObj) && mObj is PdfArray arr && arr.Items.Count >= 6) {
                double A() => (arr.Items[0] as PdfNumber)?.Value ?? 1;
                double B() => (arr.Items[1] as PdfNumber)?.Value ?? 0;
                double C() => (arr.Items[2] as PdfNumber)?.Value ?? 0;
                double D() => (arr.Items[3] as PdfNumber)?.Value ?? 1;
                double E() => (arr.Items[4] as PdfNumber)?.Value ?? 0;
                double F() => (arr.Items[5] as PdfNumber)?.Value ?? 0;
                string prefix = $"q {A().ToString(System.Globalization.CultureInfo.InvariantCulture)} {B().ToString(System.Globalization.CultureInfo.InvariantCulture)} {C().ToString(System.Globalization.CultureInfo.InvariantCulture)} {D().ToString(System.Globalization.CultureInfo.InvariantCulture)} {E().ToString(System.Globalization.CultureInfo.InvariantCulture)} {F().ToString(System.Globalization.CultureInfo.InvariantCulture)} cm ";
                content = prefix + content + " Q";
            }
            spans.AddRange(TextContentParser.Parse(content, DecodeWithFormFont, SumWidth1000Form));
        }
        return spans;
    }

    private PdfDictionary GetFormDict(string name) {
        if (_pageDict.Items.TryGetValue("Resources", out var resObj) && resObj is PdfReference rr && _objects.TryGetValue(rr.ObjectNumber, out var indr) && indr.Value is PdfDictionary res) {
            if (res.Items.TryGetValue("XObject", out var xoObj)) {
                PdfDictionary? xoDict = null;
                if (xoObj is PdfReference xref && _objects.TryGetValue(xref.ObjectNumber, out var indxo) && indxo.Value is PdfDictionary xod) xoDict = xod;
                if (xoObj is PdfDictionary d) xoDict = d;
                if (xoDict is not null && xoDict.Items.TryGetValue(name, out var formObj)) {
                    if (formObj is PdfReference fr && _objects.TryGetValue(fr.ObjectNumber, out var indForm) && indForm.Value is PdfStream s && s.Dictionary is not null) {
                        return s.Dictionary;
                    }
                }
            }
        }
        return _pageDict;
    }

    private static double GlyphWidthEmForBase(string baseFont) {
        if (string.IsNullOrEmpty(baseFont)) return 0.55;
        if (ContainsIgnoreCase(baseFont, "courier")) return 0.6;
        if (ContainsIgnoreCase(baseFont, "times")) return 0.5;
        if (ContainsIgnoreCase(baseFont, "helvetica")) return 0.55;
        return 0.55;
    }

    private static bool ContainsIgnoreCase(string source, string value) {
#if NET8_0_OR_GREATER
        return source.Contains(value, System.StringComparison.OrdinalIgnoreCase);
#else
        return source.IndexOf(value, System.StringComparison.OrdinalIgnoreCase) >= 0;
#endif
    }

    /// <summary>
    /// Returns a shallow list of content stream bytes for the page (handles single or array of streams).
    /// </summary>
    private List<byte[]> GetContentStreams() {
        var result = new List<byte[]>();
        var contents = _pageDict.Items.TryGetValue("Contents", out var obj) ? obj : null;
        if (contents is PdfReference r) {
            if (_objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfStream s) result.Add(DecodeIfNeeded(s));
        } else if (contents is PdfArray arr) {
            foreach (var item in arr.Items) {
                if (item is PdfReference rr) {
                    if (_objects.TryGetValue(rr.ObjectNumber, out var ind2) && ind2.Value is PdfStream s2) result.Add(DecodeIfNeeded(s2));
                }
            }
        }
        return result;
    }

    private static byte[] DecodeIfNeeded(PdfStream s) {
        if (PdfSyntax.HasFlateDecode(s.Dictionary)) {
            try { return Filters.FlateDecoder.Decode(s.Data); } catch { return s.Data; }
        }
        return s.Data;
    }
}
