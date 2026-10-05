using System.Globalization;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    // These fonts carry Unicode and geometry only. They reuse the configured font's
    // space glyph, including its embedded font program, and are always painted with Tr 3.
    private sealed class SearchableFontBank {
        internal string Name { get; }
        internal List<string> Scalars { get; } = new();
        internal double SpaceWidth { get; }
        internal string? CffSpaceCode { get; }
        internal SearchableFontBank(int index, PdfTextShowCommand space, bool cff) {
            Name = "FL" + index.ToString(CultureInfo.InvariantCulture);
            SpaceWidth = space.AdvanceWidth1000.GetValueOrDefault();
            CffSpaceCode = cff ? space.GlyphHex : null;
        }
    }

    private sealed class SearchableFontSet {
        internal List<SearchableFontBank> Banks { get; } = new();
        private readonly Dictionary<string, (SearchableFontBank Bank, int Code)> _codes = new(StringComparer.Ordinal);

        internal IEnumerable<(SearchableFontBank Bank, string Hex)> Encode(string text, PdfTextShowCommand space, bool cff) {
            SearchableFontBank? current = null;
            var hex = new StringBuilder();
            for (int index = 0; index < text.Length; index++) {
                int length = char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1]) ? 2 : 1;
                string scalar = text.Substring(index, length);
                index += length - 1;
                if (!_codes.TryGetValue(scalar, out var mapped)) {
                    if (Banks.Count == 0 || Banks[Banks.Count - 1].Scalars.Count == (cff ? 1 : 255))
                        Banks.Add(new SearchableFontBank(Banks.Count + 1, space, cff));
                    var bank = Banks[Banks.Count - 1];
                    bank.Scalars.Add(scalar);
                    mapped = (bank, bank.Scalars.Count);
                    _codes.Add(scalar, mapped);
                }
                if (current != null && current != mapped.Bank) {
                    yield return (current, hex.ToString());
                    hex.Clear();
                }
                current = mapped.Bank;
                hex.Append(mapped.Bank.CffSpaceCode ?? mapped.Code.ToString("X2", CultureInfo.InvariantCulture));
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
                " /Widths [" + string.Join(" ", bank.Scalars.Select(_ => bank.SpaceWidth.ToString(CultureInfo.InvariantCulture))) + "]";
            if (descendantId != 0) {
                // Keep CFF CID ownership intact. PDFKit resolves these mappings by CID;
                // a separate mapping per scalar avoids aliasing every character to space.
                body = "<< /Type /Font /Subtype /Type0 /BaseFont /" + PdfSyntaxEscaper.Name(fontName) +
                    " /DescendantFonts [" + PdfSyntaxEscaper.IndirectReference(descendantId) + "]";
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
                map.Append('<').Append(bank.CffSpaceCode ?? (i + 1).ToString("X2", CultureInfo.InvariantCulture)).Append("> ");
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
