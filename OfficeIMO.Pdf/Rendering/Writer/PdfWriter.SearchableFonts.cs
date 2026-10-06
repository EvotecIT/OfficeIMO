using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // These fonts carry Unicode and geometry only and always use Tr 3. Simple
    // fonts reuse space; contiguous CFF runs retain distinct native glyph IDs
    // when available so one font can carry the run without Unicode aliasing.
    // Custom widths are exclusive to these invisible banks: PDF Association
    // TechNote 0010, A006, exempts Tr 3 glyphs from font-program width consistency.
    // https://pdfa.org/wp-content/uploads/2017/07/TechNote0010.pdf
    private sealed class SearchableFontBank {
        internal string Name { get; }
        internal List<string> Scalars { get; } = new();
        internal List<double> Widths { get; } = new();
        internal List<string> Codes { get; } = new();
        internal string? CffSpaceCode { get; }
        internal SearchableFontBank(int index, PdfTextShowCommand space, bool cff) {
            Name = "FL" + index.ToString(CultureInfo.InvariantCulture);
            CffSpaceCode = cff ? space.GlyphHex : null;
        }
    }

    private sealed class SearchableFontSet {
        internal List<SearchableFontBank> Banks { get; } = new();
        private readonly Dictionary<(string Scalar, double Width, string? Glyph), (SearchableFontBank Bank, int Code)> _codes = new();

        internal IEnumerable<(SearchableFontBank Bank, string Hex)> Encode(string text, PdfTextShowCommand space, bool cff, double? width = null, PdfOpenTypeCffFontProgram? cffProgram = null) {
            SearchableFontBank? current = null;
            var hex = new StringBuilder();
            for (int index = 0; index < text.Length; index++) {
                int length = char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1]) ? 2 : 1;
                string scalar = text.Substring(index, length);
                index += length - 1;
                string? glyph = cff ? space.GlyphHex : null;
                if (cffProgram != null && cffProgram.TryGetGlyphId(char.ConvertToUtf32(scalar, 0), out int glyphId) && glyphId > 0)
                    glyph = cffProgram.EncodeTextAsGlyphHex(scalar);
                double exactWidth = width ?? space.AdvanceWidth1000.GetValueOrDefault();
                double roundedWidth = Math.Round(exactWidth, 9);
                var key = (scalar, roundedWidth > 0 ? roundedWidth : exactWidth, glyph);
                if (!_codes.TryGetValue(key, out var mapped)) {
                    if (Banks.Count == 0 || Banks[Banks.Count - 1].Scalars.Count == 255 || (cff && Banks[Banks.Count - 1].Codes.Contains(glyph!)))
                        Banks.Add(new SearchableFontBank(Banks.Count + 1, space, cff));
                    var bank = Banks[Banks.Count - 1];
                    bank.Scalars.Add(scalar);
                    bank.Widths.Add(key.Item2);
                    bank.Codes.Add(glyph ?? bank.Scalars.Count.ToString("X2", CultureInfo.InvariantCulture));
                    mapped = (bank, bank.Scalars.Count);
                    _codes.Add(key, mapped);
                }
                if (current != null && current != mapped.Bank) {
                    yield return (current, hex.ToString());
                    hex.Clear();
                }
                current = mapped.Bank;
                hex.Append(mapped.Bank.Codes[mapped.Code - 1]);
            }
            if (current != null) yield return (current, hex.ToString());
        }
    }

    private static void MaterializeSearchableFonts(IList<byte[]> objects,
        IReadOnlyList<(int SourceId, int Id, SearchableFontBank Bank)> pending, int sourceId, string fontName, int descriptorId = 0, bool trueType = false, int descendantId = 0) {
        foreach (var entry in pending) {
            if (entry.SourceId != sourceId) continue;
            var bank = entry.Bank;
            int unicodeId = AddStreamObject(objects, BuildSearchableCMap(bank));
            string encoding = "<< /Type /Encoding /Differences [1 " + string.Join(" ", bank.Scalars.Select(_ => "/space")) + "] >>";
            string body = "<< /Type /Font /BaseFont /" + PdfSyntaxEscaper.Name(fontName) +
                " /Subtype /" + (trueType ? "TrueType" : "Type1") +
                " /FirstChar 1 /LastChar " + bank.Scalars.Count.ToString(CultureInfo.InvariantCulture) +
                " /Widths [" + string.Join(" ", bank.Widths.Select(width => width.ToString(CultureInfo.InvariantCulture))) + "]";
            if (descendantId != 0) {
                // Keep CFF CID ownership intact. PDFKit resolves these mappings by CID;
                // distinct native CIDs share a bank; unavailable glyphs use separate
                // space mappings rather than aliasing different Unicode strings.
                int logicalDescendant = AddObject(objects, "<< /Type /Font /Subtype /CIDFontType0 /BaseFont /" + PdfSyntaxEscaper.Name(fontName) +
                    " /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor " +
                    PdfSyntaxEscaper.IndirectReference(descriptorId) + " /W [" + string.Join(" ", bank.Codes.Select((code, i) =>
                        int.Parse(code, NumberStyles.HexNumber, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture) + " [" + bank.Widths[i].ToString(CultureInfo.InvariantCulture) + "]")) + "] >>");
                body = "<< /Type /Font /Subtype /Type0 /BaseFont /" + PdfSyntaxEscaper.Name(fontName) +
                    " /DescendantFonts [" + PdfSyntaxEscaper.IndirectReference(logicalDescendant) + "]";
                encoding = "/Identity-H";
            } else if (descriptorId != 0) body += " /FontDescriptor " + PdfSyntaxEscaper.IndirectReference(descriptorId);
            ReplaceObject(objects, entry.Id, body + " /Encoding " + encoding + " /ToUnicode " + PdfSyntaxEscaper.IndirectReference(unicodeId) + " >>");
        }
    }

    private static byte[] BuildSearchableCMap(SearchableFontBank bank) {
        var map = new StringBuilder("/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> def\n/CMapName /OfficeIMO-Logical");
        map.Append("-UCS def\n/CMapType 2")
            .Append(" def\n1 begincodespacerange\n").Append(bank.CffSpaceCode == null ? "<01> <FF>" : "<0000> <FFFF>").Append("\nendcodespacerange\n");
        for (int start = 0; start < bank.Scalars.Count; start += 100) {
            int count = Math.Min(100, bank.Scalars.Count - start);
            map.Append(count.ToString(CultureInfo.InvariantCulture)).Append(" beginbfchar\n");
            for (int i = start; i < start + count; i++) {
                map.Append('<').Append(bank.Codes[i]).Append("> ");
                map.Append('<');
                foreach (char unit in bank.Scalars[i]) map.Append(((int)unit).ToString("X4", CultureInfo.InvariantCulture));
                map.Append('>');
                map.Append('\n');
            }
            map.Append("endbfchar\n");
        }
        map.Append("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");
        return Encoding.ASCII.GetBytes(map.ToString());
    }
}
