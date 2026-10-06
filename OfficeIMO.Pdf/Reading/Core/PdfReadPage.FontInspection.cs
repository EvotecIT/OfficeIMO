namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal PdfDictionary? GetFontInspectionResources() =>
        ResolveDictionary(GetInheritedValue("Resources"));

    internal PdfSourceTextFont? GetSourceTextFont(PdfTextSpan span) => GetSourceTextFont(span.FontResource, span.BaseFont);

    internal PdfSourceTextFont? GetSourceTextFont(string resourceName, string? baseFont) {
        PdfFontResourceSet resources = _fontResourceCache.GetOrCreate(GetFontInspectionResources(), _objects);
        if (!resources.Fonts.TryGetValue(resourceName, out PdfFontResource? font) ||
            !string.Equals(font.BaseFont, baseFont, StringComparison.Ordinal) ||
            !(font.FontSubtype == "TrueType" || (font.FontSubtype == "Type0" && font.Encoding == "Identity-H")) || font.IsVerticalWriting ||
            font.CMap is null || font.EmbeddedTrueTypeFont is null ||
            !resources.WidthProviders.TryGetValue(resourceName, out Func<byte[], double>? width)) return null;
        return new PdfSourceTextFont(font, width);
    }
}
