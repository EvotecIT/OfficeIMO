namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    private static PdfDictionary? ResolveFontDescriptor(
        PdfDictionary font,
        Dictionary<int, PdfIndirectObject> objects,
        out string? programFontSubtype) {
        PdfDictionary fontWithDescriptor = font;
        programFontSubtype = font.Get<PdfName>("Subtype")?.Name;
        if (string.Equals(programFontSubtype, "Type0", StringComparison.Ordinal) &&
            font.Items.TryGetValue("DescendantFonts", out PdfObject? descendantsObject)) {
            PdfArray? descendants = ResolveArray(descendantsObject, objects);
            PdfDictionary? descendant = descendants is { Items.Count: > 0 }
                ? ResolveDict(descendants.Items[0], objects)
                : null;
            if (descendant != null) {
                fontWithDescriptor = descendant;
                programFontSubtype = descendant.Get<PdfName>("Subtype")?.Name;
            }
        }

        return fontWithDescriptor.Items.TryGetValue("FontDescriptor", out PdfObject? descriptorObject)
            ? ResolveDict(descriptorObject, objects)
            : null;
    }

    private const int MaxCidToGlyphMapBytes = 65536 * 2;

    /// <summary>
    /// Reads a CIDFontType2 CIDToGIDMap. An empty array means the identity mapping; null means the
    /// font is not a CIDFontType2 font or the stream is unusable.
    /// </summary>
    internal static byte[]? TryReadCidToGlyphMap(PdfDictionary font, Dictionary<int, PdfIndirectObject> objects) {
        if (!string.Equals(font.Get<PdfName>("Subtype")?.Name, "Type0", StringComparison.Ordinal) ||
            !font.Items.TryGetValue("DescendantFonts", out PdfObject? descendantsObject) ||
            ResolveArray(descendantsObject, objects) is not { Items.Count: > 0 } descendants ||
            ResolveDict(descendants.Items[0], objects) is not PdfDictionary descendant ||
            !string.Equals(descendant.Get<PdfName>("Subtype")?.Name, "CIDFontType2", StringComparison.Ordinal)) return null;
        if (!descendant.Items.TryGetValue("CIDToGIDMap", out PdfObject? mapObject)) return Array.Empty<byte>();
        PdfObject? map = ResolveObject(mapObject, objects);
        if (map is PdfName { Name: "Identity" }) return Array.Empty<byte>();
        if (map is not PdfStream stream || Filters.StreamDecoder.GetUnsupportedFilters(stream.Dictionary, objects).Count != 0) return null;
        return Filters.StreamDecoder.TryDecode(stream, MaxCidToGlyphMapBytes, out byte[] bytes, objects) &&
            bytes.Length >= 2 ? bytes : null;
    }

    private static int? TryReadFontDescriptorInteger(
        PdfDictionary? descriptor,
        Dictionary<int, PdfIndirectObject> objects,
        string key,
        int minimum,
        int maximum) {
        if (descriptor == null ||
            !descriptor.Items.TryGetValue(key, out PdfObject? value) ||
            ResolveObject(value, objects) is not PdfNumber number ||
            double.IsNaN(number.Value) ||
            double.IsInfinity(number.Value) ||
            number.Value != Math.Truncate(number.Value) ||
            number.Value < minimum ||
            number.Value > maximum) return null;
        return (int)number.Value;
    }
}
