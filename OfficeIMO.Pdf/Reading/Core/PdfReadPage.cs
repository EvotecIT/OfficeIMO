using System.Text;

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

    /// <summary>Extracts plain text from this page by parsing content operators.</summary>
    public string ExtractText() {
        var spans = GetTextSpans();
        var sb = new StringBuilder();
        for (int i = 0; i < spans.Count; i++) {
            if (i > 0) sb.Append('\n');
            sb.Append(spans[i].Text);
        }
        return sb.ToString();
    }

    /// <summary>Gets text spans (text with position and font info) from this page.</summary>
    public IReadOnlyList<PdfTextSpan> GetTextSpans() {
        var spans = new List<PdfTextSpan>();
        var streams = GetContentStreams();
        var decoders = ResourceResolver.GetFontDecoders(_pageDict, _objects);
        var fonts = ResourceResolver.GetFontsForPage(_pageDict, _objects);
        string DecodeWithFont(string fontRes, byte[] bytes) =>
            decoders.TryGetValue(fontRes, out var dec) ? dec(bytes) : PdfWinAnsiEncoding.Decode(bytes);
        double WidthEmForFont(string fontRes) {
            if (fonts.TryGetValue(fontRes, out var f)) return GlyphWidthEmForBase(f.BaseFont);
            return 0.55; // default
        }
        foreach (var s in streams) {
            var content = PdfEncoding.Latin1GetString(s);
            spans.AddRange(TextContentParser.Parse(content, DecodeWithFont, WidthEmForFont));
        }
        return spans;
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
            if (_objects.TryGetValue(r.ObjectNumber, out var ind) && ind.Value is PdfStream s) result.Add(s.Data);
        } else if (contents is PdfArray arr) {
            foreach (var item in arr.Items) {
                if (item is PdfReference rr) {
                    if (_objects.TryGetValue(rr.ObjectNumber, out var ind2) && ind2.Value is PdfStream s2) result.Add(s2.Data);
                }
            }
        }
        return result;
    }
}
