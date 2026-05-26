namespace OfficeIMO.Pdf;

internal static class PdfStandardFontDictionaryBuilder {
    private const string FontType = "Font";
    private const string Type1Subtype = "Type1";
    private const string WinAnsiEncoding = "WinAnsiEncoding";

    internal static string BuildStandardType1FontObject(PdfStandardFont font) {
        string baseFont = font.ToBaseFontName();
        return "<< /Type /" + PdfSyntaxEscaper.Name(FontType) +
            " /Subtype /" + PdfSyntaxEscaper.Name(Type1Subtype) +
            " /BaseFont /" + PdfSyntaxEscaper.Name(baseFont) +
            " /Encoding /" + PdfSyntaxEscaper.Name(WinAnsiEncoding) +
            " >>\n";
    }

    internal static PdfDictionary BuildStandardType1FontDictionary(PdfStandardFont font) {
        string baseFont = font.ToBaseFontName();
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName(FontType);
        dictionary.Items["Subtype"] = new PdfName(Type1Subtype);
        dictionary.Items["BaseFont"] = new PdfName(baseFont);
        dictionary.Items["Encoding"] = new PdfName(WinAnsiEncoding);
        return dictionary;
    }
}
