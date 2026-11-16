using System.Globalization;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

/// <summary>
/// Composite fonts whose ToUnicode map is absent or incomplete: the CID's text comes from the Adobe character
/// collection and then from the embedded font program, as PDF viewers read it.
/// </summary>
public sealed class PdfCompositeCidTextTests {
    // Adobe-Japan1 CIDs 1-95 are the printable ASCII characters, so CID = character - 0x1F.
    private const string HelloJapan1Cids = "00290046004D004D0050";

    [Theory]
    [InlineData("Identity-H")]
    [InlineData("Identity-V")]
    public void IdentityCodesWithoutToUnicodeReadThroughTheAdobeCollection(string encoding) {
        byte[] pdf = BuildCompositePdf(encoding, "Japan1", HelloJapan1Cids);

        Assert.Contains("Hello", PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(pdf).Pages[0].GetTextSpans());
        Assert.Equal(PdfTextDecodingSource.PredefinedCMap, span.DecodingEvidence.Source);
        Assert.False(span.DecodingEvidence.UsedFallback);
        PdfDocumentReadResult read = PdfDocument.Load(pdf).Read();
        Assert.Contains("Hello", read.Text, StringComparison.Ordinal);
        Assert.Empty(read.TextDecodingDiagnostics);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void IdentityCodesWithoutToUnicodeReadThroughTheEmbeddedTrueTypeCmap(bool cidToGlyphStream) {
        byte[] font = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Typography", "OfficeIMOBaselineSans-Regular.ttf"));
        PdfTrueTypeFontProgram program = PdfTrueTypeFontProgram.Parse(font);
        const string expected = "Zürich Text";
        int[] glyphs = expected.Select(character => {
            Assert.True(program.TryGetGlyphId(character, out int glyph));
            return glyph;
        }).ToArray();
        // With a CIDToGIDMap stream the codes are CIDs 1..n; with Identity they are the glyph indices themselves.
        string codes = string.Concat((cidToGlyphStream ? Enumerable.Range(1, glyphs.Length) : glyphs)
            .Select(static code => code.ToString("X4", CultureInfo.InvariantCulture)));
        byte[]? map = null;
        if (cidToGlyphStream) {
            map = new byte[(glyphs.Length + 1) * 2];
            for (int cid = 1; cid <= glyphs.Length; cid++) {
                map[2 * cid] = (byte)(glyphs[cid - 1] >> 8);
                map[2 * cid + 1] = (byte)glyphs[cid - 1];
            }
        }

        byte[] pdf = BuildCompositePdf("Identity-H", "Identity", codes, descendantSubtype: "CIDFontType2", fontProgram: font, cidToGlyphMap: map);

        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(pdf).Pages[0].GetTextSpans());
        Assert.Equal(expected, span.Text);
        Assert.Equal(PdfTextDecodingSource.FontProgram, span.DecodingEvidence.Source);
        Assert.False(span.DecodingEvidence.UsedFallback);
        Assert.Contains(expected, PdfDocument.Load(pdf).Read().Text, StringComparison.Ordinal);
    }

    // Identity codes are Japan1 CIDs (H, B, C, D); UniJIS-UCS2-H codes are UTF-16. ToUnicode maps only the first code.
    [Theory]
    [InlineData("Identity-H", "0029002300240025")]
    [InlineData("UniJIS-UCS2-H", "0041004200430044")]
    public void ToUnicodeKeepsPrecedenceAndTheCollectionSuppliesCodesItOmits(string encoding, string codes) {
        string firstCode = codes.Substring(0, 4);
        string toUnicode = "1 beginbfchar <" + firstCode + "> <0058> endbfchar";
        byte[] pdf = BuildCompositePdf(encoding, "Japan1", codes, toUnicode: toUnicode);

        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(pdf).Pages[0].GetTextSpans());
        Assert.Equal("XBCD", span.Text);
        Assert.Equal(PdfTextDecodingSource.ToUnicode, span.DecodingEvidence.Source);
        Assert.Equal(0, span.DecodingEvidence.UnmappedCodeBytes);
        Assert.Equal("XBCD", PdfDocument.Load(pdf).Read().Text.Trim());

        // A code neither ToUnicode nor the collection maps stays unmapped: a named CMap refuses the run, and an
        // Identity run reports the code's bytes as unmapped.
        byte[] unmapped = BuildCompositePdf(encoding, "Japan1", firstCode + "FFFF", toUnicode: toUnicode);
        if (encoding == "UniJIS-UCS2-H") {
            Assert.Throws<PdfUnsupportedTextMappingException>(() => PdfReadDocument.Open(unmapped).ExtractText());
        } else {
            PdfTextSpan partial = Assert.Single(PdfReadDocument.Open(unmapped).Pages[0].GetTextSpans());
            Assert.Equal(2, partial.DecodingEvidence.UnmappedCodeBytes);
            Assert.True(partial.DecodingEvidence.UsedFallback);
        }
    }

    [Fact]
    public void IdentityToUnicodeNameReadsEachCodeAsUtf16() {
        string codes = string.Concat("Grüße".Select(static character => ((int)character).ToString("X4", CultureInfo.InvariantCulture)));
        byte[] pdf = BuildCompositePdf("Identity-H", "Identity", codes, toUnicodeName: "Identity-H");

        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(pdf).Pages[0].GetTextSpans());

        Assert.Equal("Grüße", span.Text);
        Assert.Equal(PdfTextDecodingSource.ToUnicode, span.DecodingEvidence.Source);
        Assert.Equal("Grüße", PdfDocument.Load(pdf).Read().Text.Trim());
    }

    // UTF-8 style codespaces mix one- to four-byte codes. The embedded CMap splits the text even when its collection has
    // no bundled Unicode map, so ToUnicode can map each code.
    [Fact]
    public void EmbeddedMixedLengthEncodingCMapSplitsCodesForToUnicode() {
        const string encodingCMap = "/CIDInit /ProcSet findresource begin 12 dict begin begincmap\n" +
            "/CIDSystemInfo << /Registry (Adobe) /Ordering (Custom) /Supplement 0 >> def /CMapName /Custom-UTF8 def /CMapType 1 def\n" +
            "4 begincodespacerange\n<00> <7F>\n<C080> <DFBF>\n<E08080> <EFBFBF>\n<F0808080> <F7BFBFBF>\nendcodespacerange\n" +
            "4 begincidchar\n<61> 1\n<62> 2\n<e38081> 3\n<f0a8a780> 4\nendcidchar\nendcmap CMapName currentdict /CMap defineresource pop end end";
        const string toUnicode = "4 beginbfchar\n<61> <0061>\n<62> <0062>\n<e38081> <3001>\n<f0a8a780> <D862DDC0>\nendbfchar";

        byte[] pdf = BuildCompositePdf(encodingCMap, "Custom", "61E3808162F0A8A780", toUnicode: toUnicode, embeddedEncoding: true);

        Assert.Equal("a、b\U000289C0", PdfReadDocument.Open(pdf).Pages[0].ExtractText().Trim());
        Assert.Equal("a、b\U000289C0", PdfDocument.Load(pdf).Read().Text.Trim());
    }

    // A simple font's ToUnicode map can cover only some codes; Differences name the rest, including accented
    // letters spelled as a base letter and an accent.
    [Fact]
    public void PartialSimpleToUnicodeReadsOmittedCodesThroughDifferences() {
        string content = "BT /F1 12 Tf 20 100 Td <01020341> Tj ET";
        byte[] pdf = BuildPdf(
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 300 200] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica /Encoding << /Type /Encoding /BaseEncoding /WinAnsiEncoding " +
                "/Differences [1 /ccaron /uring /Scommaaccent] >> /ToUnicode 6 0 R >>",
            Stream(content),
            Stream("1 beginbfchar <41> <0058> endbfchar"));

        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(pdf).Pages[0].GetTextSpans());

        Assert.Equal("čůȘX", span.Text);
        Assert.Equal(PdfTextDecodingSource.ToUnicode, span.DecodingEvidence.Source);
        Assert.False(span.DecodingEvidence.UsedFallback);
        Assert.Equal("čůȘX", PdfDocument.Load(pdf, new PdfLoadOptions { UseWinAnsiFallback = false }).Read().Text.Trim());
    }

    private static byte[] BuildCompositePdf(string encoding, string ordering, string codes, string? toUnicode = null,
        string? toUnicodeName = null, string descendantSubtype = "CIDFontType0", byte[]? fontProgram = null,
        byte[]? cidToGlyphMap = null, bool embeddedEncoding = false) {
        var objects = new List<object> {
            "<< /Type /Catalog /Pages 2 0 R >>",
            "<< /Type /Pages /Count 1 /Kids [3 0 R] >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 4 0 R >> >> /Contents 5 0 R >>",
            string.Empty, // 4: Type0 font
            Stream("BT /F1 18 Tf 72 700 Td <" + codes + "> Tj ET"),
            string.Empty // 6: descendant
        };
        string fontEntries = string.Empty;
        if (embeddedEncoding) {
            objects.Add(Stream(encoding));
            fontEntries += " /Encoding " + objects.Count + " 0 R";
        } else {
            fontEntries += " /Encoding /" + encoding;
        }
        if (toUnicode != null) {
            objects.Add(Stream(toUnicode));
            fontEntries += " /ToUnicode " + objects.Count + " 0 R";
        } else if (toUnicodeName != null) {
            fontEntries += " /ToUnicode /" + toUnicodeName;
        }
        string descendantEntries = string.Empty;
        if (fontProgram != null) {
            objects.Add((Dictionary: "<< /Length1 " + fontProgram.Length.ToString(CultureInfo.InvariantCulture) + " >>", Data: fontProgram));
            objects.Add("<< /Type /FontDescriptor /FontName /OfficeIMOBaselineSans /Flags 32 /FontBBox [0 -250 1000 900] /ItalicAngle 0 " +
                "/Ascent 900 /Descent -250 /CapHeight 700 /StemV 80 /FontFile2 " + (objects.Count) + " 0 R >>");
            descendantEntries += " /FontDescriptor " + objects.Count + " 0 R";
        }
        if (cidToGlyphMap != null) {
            objects.Add((Dictionary: "<< >>", Data: cidToGlyphMap));
            descendantEntries += " /CIDToGIDMap " + objects.Count + " 0 R";
        } else if (descendantSubtype == "CIDFontType2") {
            descendantEntries += " /CIDToGIDMap /Identity";
        }
        objects[3] = "<< /Type /Font /Subtype /Type0 /BaseFont /TestCid" + fontEntries + " /DescendantFonts [6 0 R] >>";
        objects[5] = "<< /Type /Font /Subtype /" + descendantSubtype + " /BaseFont /TestCid /CIDSystemInfo << /Registry (Adobe) /Ordering (" +
            ordering + ") /Supplement 0 >> /DW 1000" + descendantEntries + " >>";
        return BuildPdf(objects.ToArray());
    }

    private static string Stream(string content) =>
        "<< /Length " + Encoding.ASCII.GetByteCount(content).ToString(CultureInfo.InvariantCulture) + " >>\nstream\n" + content + "\nendstream";

    // Objects are either complete object text or a stream's dictionary and binary data.
    private static byte[] BuildPdf(params object[] objects) {
        using var pdf = new MemoryStream();
        void Write(string text) {
            byte[] bytes = Encoding.ASCII.GetBytes(text);
            pdf.Write(bytes, 0, bytes.Length);
        }
        Write("%PDF-1.7\n");
        var offsets = new List<long>();
        for (int index = 0; index < objects.Length; index++) {
            offsets.Add(pdf.Position);
            Write((index + 1).ToString(CultureInfo.InvariantCulture) + " 0 obj\n");
            if (objects[index] is ValueTuple<string, byte[]> stream) {
                string dictionary = stream.Item1.Substring(0, stream.Item1.Length - 2) +
                    "/Length " + stream.Item2.Length.ToString(CultureInfo.InvariantCulture) + " >>";
                Write(dictionary + "\nstream\n");
                pdf.Write(stream.Item2, 0, stream.Item2.Length);
                Write("\nendstream");
            } else {
                Write((string)objects[index]);
            }
            Write("\nendobj\n");
        }
        long xref = pdf.Position;
        Write("xref\n0 " + (objects.Length + 1).ToString(CultureInfo.InvariantCulture) + "\n0000000000 65535 f \n");
        foreach (long offset in offsets) Write(offset.ToString("D10", CultureInfo.InvariantCulture) + " 00000 n \n");
        Write("trailer\n<< /Size " + (objects.Length + 1).ToString(CultureInfo.InvariantCulture) + " /Root 1 0 R >>\nstartxref\n" +
            xref.ToString(CultureInfo.InvariantCulture) + "\n%%EOF\n");
        return pdf.ToArray();
    }
}
