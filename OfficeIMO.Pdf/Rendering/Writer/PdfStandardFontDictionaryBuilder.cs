namespace OfficeIMO.Pdf;

internal static class PdfStandardFontDictionaryBuilder {
    private const string FontType = "Font";
    private const string Type1Subtype = "Type1";
    private const string TrueTypeSubtype = "TrueType";
    private const string Type0Subtype = "Type0";
    private const string CidFontType2Subtype = "CIDFontType2";
    private const string WinAnsiEncoding = "WinAnsiEncoding";

    internal static string BuildStandardType1FontObject(PdfStandardFont font, int toUnicodeObjectId = 0) {
        if (toUnicodeObjectId < 0) {
            throw new ArgumentOutOfRangeException(nameof(toUnicodeObjectId), "PDF ToUnicode object number cannot be negative.");
        }

        string baseFont = font.ToBaseFontName();
        string body = "<< /Type /" + PdfSyntaxEscaper.Name(FontType) +
            " /Subtype /" + PdfSyntaxEscaper.Name(Type1Subtype) +
            " /BaseFont /" + PdfSyntaxEscaper.Name(baseFont) +
            " /Encoding /" + PdfSyntaxEscaper.Name(WinAnsiEncoding);
        if (toUnicodeObjectId > 0) {
            body += " /ToUnicode " + PdfSyntaxEscaper.IndirectReference(toUnicodeObjectId);
        }

        return body + " >>\n";
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

    internal static string BuildEmbeddedTrueTypeFontObject(PdfTrueTypeFontProgram font, int descriptorObjectId, int toUnicodeObjectId = 0) {
        Guard.NotNull(font, nameof(font));
        if (descriptorObjectId <= 0) {
            throw new ArgumentOutOfRangeException(nameof(descriptorObjectId), "PDF font descriptor object number must be positive.");
        }

        if (toUnicodeObjectId < 0) {
            throw new ArgumentOutOfRangeException(nameof(toUnicodeObjectId), "PDF ToUnicode object number cannot be negative.");
        }

        var sb = new StringBuilder();
        sb.Append("<< /Type /").Append(PdfSyntaxEscaper.Name(FontType))
            .Append(" /Subtype /").Append(PdfSyntaxEscaper.Name(TrueTypeSubtype))
            .Append(" /BaseFont /").Append(PdfSyntaxEscaper.Name(font.FontName))
            .Append(" /Encoding /").Append(PdfSyntaxEscaper.Name(WinAnsiEncoding))
            .Append(" /FirstChar 32 /LastChar 255 /Widths [");
        int[] widths = font.BuildWinAnsiWidths();
        for (int i = 0; i < widths.Length; i++) {
            if (i > 0) {
                sb.Append(' ');
            }

            sb.Append(widths[i].ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        sb.Append("] /FontDescriptor ")
            .Append(PdfSyntaxEscaper.IndirectReference(descriptorObjectId));
        if (toUnicodeObjectId > 0) {
            sb.Append(" /ToUnicode ")
                .Append(PdfSyntaxEscaper.IndirectReference(toUnicodeObjectId));
        }

        sb.Append(" >>\n");
        return sb.ToString();
    }

    internal static string BuildEmbeddedType0FontObject(PdfTrueTypeFontProgram font, int descendantFontObjectId, int toUnicodeObjectId) {
        Guard.NotNull(font, nameof(font));
        if (descendantFontObjectId <= 0) {
            throw new ArgumentOutOfRangeException(nameof(descendantFontObjectId), "PDF descendant font object number must be positive.");
        }

        if (toUnicodeObjectId <= 0) {
            throw new ArgumentOutOfRangeException(nameof(toUnicodeObjectId), "PDF Type0 embedded fonts require a ToUnicode object number.");
        }

        return "<< /Type /" + PdfSyntaxEscaper.Name(FontType) +
            " /Subtype /" + PdfSyntaxEscaper.Name(Type0Subtype) +
            " /BaseFont /" + PdfSyntaxEscaper.Name(font.FontName) +
            " /Encoding /Identity-H" +
            " /DescendantFonts [ " + PdfSyntaxEscaper.IndirectReference(descendantFontObjectId) + " ]" +
            " /ToUnicode " + PdfSyntaxEscaper.IndirectReference(toUnicodeObjectId) +
            " >>\n";
    }

    internal static string BuildCidFontType2DescendantObject(PdfTrueTypeFontProgram font, int descriptorObjectId) {
        Guard.NotNull(font, nameof(font));
        if (descriptorObjectId <= 0) {
            throw new ArgumentOutOfRangeException(nameof(descriptorObjectId), "PDF font descriptor object number must be positive.");
        }

        var sb = new StringBuilder();
        sb.Append("<< /Type /").Append(PdfSyntaxEscaper.Name(FontType))
            .Append(" /Subtype /").Append(PdfSyntaxEscaper.Name(CidFontType2Subtype))
            .Append(" /BaseFont /").Append(PdfSyntaxEscaper.Name(font.FontName))
            .Append(" /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >>")
            .Append(" /FontDescriptor ").Append(PdfSyntaxEscaper.IndirectReference(descriptorObjectId))
            .Append(" /DW 500 /W [0 [");

        for (int glyphId = 0; glyphId < font.GlyphCount; glyphId++) {
            if (glyphId > 0) {
                sb.Append(' ');
            }

            sb.Append(font.GetGlyphWidth1000(glyphId).ToString(System.Globalization.CultureInfo.InvariantCulture));
        }

        sb.Append("]] /CIDToGIDMap /Identity >>\n");
        return sb.ToString();
    }

    internal static string BuildTrueTypeFontDescriptorObject(PdfTrueTypeFontProgram font, int fontFileObjectId) {
        Guard.NotNull(font, nameof(font));
        if (fontFileObjectId <= 0) {
            throw new ArgumentOutOfRangeException(nameof(fontFileObjectId), "PDF embedded font file object number must be positive.");
        }

        return "<< /Type /FontDescriptor /FontName /" + PdfSyntaxEscaper.Name(font.FontName) +
            " /Flags " + font.Flags.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " /FontBBox [" + string.Join(" ", font.FontBBox.Select(value => value.ToString(System.Globalization.CultureInfo.InvariantCulture))) + "]" +
            " /ItalicAngle " + font.ItalicAngle.ToString("0.###", System.Globalization.CultureInfo.InvariantCulture) +
            " /Ascent " + font.Ascent.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " /Descent " + font.Descent.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " /CapHeight " + font.CapHeight.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " /StemV " + font.StemV.ToString(System.Globalization.CultureInfo.InvariantCulture) +
            " /FontFile2 " + PdfSyntaxEscaper.IndirectReference(fontFileObjectId) +
            " >>\n";
    }
}
