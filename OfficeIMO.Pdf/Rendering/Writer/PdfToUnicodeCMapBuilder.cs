using System.Globalization;

namespace OfficeIMO.Pdf;

internal static class PdfToUnicodeCMapBuilder {
    internal static byte[] BuildWinAnsiToUnicodeCMap() {
        var sb = new StringBuilder();
        sb.Append("/CIDInit /ProcSet findresource begin\n");
        sb.Append("12 dict begin\n");
        sb.Append("begincmap\n");
        sb.Append("/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n");
        sb.Append("/CMapName /OfficeIMO-WinAnsi-UCS def\n");
        sb.Append("/CMapType 2 def\n");
        sb.Append("1 begincodespacerange\n");
        sb.Append("<00> <FF>\n");
        sb.Append("endcodespacerange\n");

        var mappings = BuildMappings();
        for (int index = 0; index < mappings.Count; index += 100) {
            int count = Math.Min(100, mappings.Count - index);
            sb.Append(count.ToString(CultureInfo.InvariantCulture)).Append(" beginbfchar\n");
            for (int offset = 0; offset < count; offset++) {
                var mapping = mappings[index + offset];
                sb.Append('<')
                    .Append(mapping.Code.ToString("X2", CultureInfo.InvariantCulture))
                    .Append("> <")
                    .Append(((int)mapping.Unicode).ToString("X4", CultureInfo.InvariantCulture))
                    .Append(">\n");
            }

            sb.Append("endbfchar\n");
        }

        sb.Append("endcmap\n");
        sb.Append("CMapName currentdict /CMap defineresource pop\n");
        sb.Append("end\n");
        sb.Append("end\n");
        return Encoding.ASCII.GetBytes(sb.ToString());
    }

    internal static byte[] BuildIdentityGlyphToUnicodeCMap(PdfTrueTypeFontProgram font) {
        Guard.NotNull(font, nameof(font));

        return BuildIdentityGlyphToUnicodeCMap(font.GetGlyphToUnicodeMappings());
    }

    internal static byte[] BuildIdentityGlyphToUnicodeCMap(PdfOpenTypeCffFontProgram font) {
        Guard.NotNull(font, nameof(font));

        return BuildIdentityGlyphToUnicodeCMap(font.GetGlyphToUnicodeMappings());
    }

    private static byte[] BuildIdentityGlyphToUnicodeCMap(IReadOnlyList<(int GlyphId, string UnicodeText)> mappings) {
        var sb = new StringBuilder();
        sb.Append("/CIDInit /ProcSet findresource begin\n");
        sb.Append("12 dict begin\n");
        sb.Append("begincmap\n");
        sb.Append("/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n");
        sb.Append("/CMapName /OfficeIMO-Identity-Glyph-UCS def\n");
        sb.Append("/CMapType 2 def\n");
        sb.Append("1 begincodespacerange\n");
        sb.Append("<0000> <FFFF>\n");
        sb.Append("endcodespacerange\n");

        for (int index = 0; index < mappings.Count; index += 100) {
            int count = Math.Min(100, mappings.Count - index);
            sb.Append(count.ToString(CultureInfo.InvariantCulture)).Append(" beginbfchar\n");
            for (int offset = 0; offset < count; offset++) {
                var mapping = mappings[index + offset];
                sb.Append('<')
                    .Append(mapping.GlyphId.ToString("X4", CultureInfo.InvariantCulture))
                    .Append("> <");
                AppendUtf16Hex(sb, mapping.UnicodeText);
                sb.Append(">\n");
            }

            sb.Append("endbfchar\n");
        }

        sb.Append("endcmap\n");
        sb.Append("CMapName currentdict /CMap defineresource pop\n");
        sb.Append("end\n");
        sb.Append("end\n");
        return Encoding.ASCII.GetBytes(sb.ToString());
    }

    private static List<(byte Code, char Unicode)> BuildMappings() {
        var mappings = new List<(byte Code, char Unicode)>();
        for (int code = 0; code <= 255; code++) {
            byte value = (byte)code;
            char unicode = PdfWinAnsiEncoding.Decode(value);
            if (IsMappable(value, unicode)) {
                mappings.Add((value, unicode));
            }
        }

        return mappings;
    }

    private static bool IsMappable(byte code, char unicode) {
        if (code < 0x20 || code == 0x7F || code == 0x81 || code == 0x8D || code == 0x8F || code == 0x90 || code == 0x9D) {
            return false;
        }

        return unicode >= ' ';
    }

    private static void AppendUtf16Hex(StringBuilder sb, string text) {
        for (int index = 0; index < text.Length; index++) {
            sb.Append(((int)text[index]).ToString("X4", CultureInfo.InvariantCulture));
        }
    }
}
