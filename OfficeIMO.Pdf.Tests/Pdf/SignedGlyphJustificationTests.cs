using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSignedGlyphJustificationTests {
    [Theory]
    [InlineData("simple", 50D, false)]
    [InlineData("simple", 150D, true)]
    [InlineData("embedded-positioned", 50D, false)]
    [InlineData("embedded-logical", 150D, true)]
    public void SavedPdfPreservesSignedUnequalAdvancesAndTrackedPen(string encoding, double scale, bool zeroTotal) {
        byte[]? fontData = encoding == "simple" ? null :
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B', 'C', 'X');
        PdfTextShowCommand natural = Command("ABC", encoding, fontData);
        PdfTextShowCommand probe = Command("X", encoding, fontData);
        double[] expected = zeroTotal ? new[] { -1.5D, 2.25D, -0.75D } : new[] { -1.5D, 0.25D, -0.5D };
        double[] supplied = (double[])expected.Clone();
        PdfTextShowCommand adjusted = natural.WithJustificationAdvances(supplied);
        Array.Clear(supplied, 0, supplied.Length);
        Assert.True(adjusted.HasJustificationAdvances);
        Assert.Equal(expected.Sum(), adjusted.TotalJustificationAdvance);
        Assert.False(natural.HasJustificationAdvances);

        (byte[] controlPdf, _) = SaveCommandPdf(natural, probe, scale, fontData);
        (byte[] actualPdf, string content) = SaveCommandPdf(adjusted, probe, scale, fontData);
        double[] control = SavedOrigins(controlPdf, encoding + "-" + scale + "-" + zeroTotal + "-control");
        double[] actual = SavedOrigins(actualPdf, encoding + "-" + scale + "-" + zeroTotal);

        double cumulative = 0D;
        for (int index = 0; index < 3; index++) {
            Assert.InRange(Math.Abs(actual[index] - control[index] - cumulative), 0D, 0.002D);
            cumulative += expected[index];
        }
        Assert.InRange(Math.Abs(actual[3] - control[3] - cumulative), 0D, 0.002D);
        // The following X is positioned by a folded text-matrix request. A wrong
        // tracked sum, sign or scale moves X even if each earlier glyph paints correctly.
        var operations = new List<PdfContentOperation>();
        PdfContentStreamInterpreter.Interpret(content, 100, operations.Add);
        Assert.Single(operations, operation => operation.Name == "Tm");
        double[] movements = operations.Where(operation => operation.Name == "TJ")
            .SelectMany(operation => Assert.IsType<List<object>>(operation.Operands[0]))
            .OfType<double>().ToArray();
        Assert.Contains(movements, movement => movement > 0D);
        Assert.Contains(movements, movement => movement < 0D);
    }

    [Fact]
    public void AdvancesRequireExactFiniteGlyphValuesAndFiniteTotal() {
        var command = new PdfTextShowCommand("414243", advanceWidth1000: 2056D);
        Assert.Throws<ArgumentNullException>(() => command.WithJustificationAdvances(null!));
        Assert.Throws<ArgumentException>(() => command.WithJustificationAdvances(new double[2]));
        Assert.Throws<ArgumentException>(() => command.WithJustificationAdvances(new[] { 0D, double.NaN, 0D }));
        Assert.Throws<ArgumentException>(() => command.WithJustificationAdvances(new[] { 0D, double.PositiveInfinity, 0D }));
        Assert.Throws<ArgumentException>(() => command.WithJustificationAdvances(new[] { double.MaxValue, double.MaxValue, 0D }));
        Assert.Throws<ArgumentException>(() => command.WithJustificationAdvances(new[] { -double.MaxValue, -double.MaxValue, 0D }));
        Assert.False(command.HasJustificationAdvances);
        Assert.Equal(0D, command.TotalJustificationAdvance);
        Assert.Equal(0D, command.GetJustificationAdvance(1));
    }

    [Fact]
    public void ZeroAdvancesRetainTheSimpleShowAndReplacementIsIndependent() {
        var original = new PdfTextShowCommand("414243", advanceWidth1000: 2056D);
        PdfTextShowCommand adjusted = original.WithJustificationAdvances(new[] { -1D, 2D, 0D });
        PdfTextShowCommand cleared = adjusted.WithJustificationAdvances(new double[3]);
        var content = new StringBuilder();
        new ContentStreamBuilder(content).BeginText().Font("F1", 10D).TextMatrix(40D, 400D)
            .ShowText(cleared, 10D).EndText();

        Assert.False(cleared.HasJustificationAdvances);
        Assert.Equal(0D, cleared.TotalJustificationAdvance);
        Assert.Equal(-1D, adjusted.GetJustificationAdvance(0));
        Assert.Equal(1D, adjusted.TotalJustificationAdvance);
        var shows = new List<string>();
        PdfContentStreamInterpreter.Interpret(content.ToString(), 100, operation => {
            if (operation.Name is "Tj" or "TJ") shows.Add(operation.Name);
        });
        Assert.Equal(new[] { "Tj" }, shows);
    }

    private static PdfTextShowCommand Command(string text, string encoding, byte[]? fontData) {
        if (fontData == null) {
            double width = 0D;
            foreach (char character in text) {
                Assert.True(PdfStandardFontWidths.TryGetWidth1000(PdfStandardFont.Helvetica, character, out int glyphWidth));
                width += glyphWidth;
            }
            return new PdfTextShowCommand(BitConverter.ToString(Encoding.ASCII.GetBytes(text)).Replace("-", ""), advanceWidth1000: width);
        }
        PdfTrueTypeFontProgram font = PdfTrueTypeFontProgram.Parse(fontData, "SignedGlyph");
        PdfGlyphRun shaped = font.ShapeText(text, PdfTextShapingOptions.ForRendering("SignedGlyph"));
        PdfGlyphInfo[] glyphs = shaped.Glyphs.ToArray();
        if (encoding == "embedded-positioned" && glyphs.Length > 1) {
            PdfGlyphInfo glyph = glyphs[1];
            glyphs[1] = new PdfGlyphInfo(glyph.GlyphId, glyph.UnicodeText, glyph.TextIndex,
                glyph.NominalWidth1000, glyph.AdvanceWidth1000 + 50, 0, 0);
        }
        return font.ToTextShowCommand(text, new PdfGlyphRun(glyphs,
            Array.Empty<PdfTextEncodingDiagnostic>(), preserveGlyphUnicode: encoding == "embedded-logical"));
    }

    private static (byte[] Pdf, string Content) SaveCommandPdf(PdfTextShowCommand command,
        PdfTextShowCommand probe, double scale, byte[]? fontData) {
        var options = new PdfOptions { DefaultFontSize = 10D };
        if (fontData != null) options.UseFontFamily(new PdfEmbeddedFontFamily("SignedGlyph", fontData));
        byte[] template = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Text("ABCX")).ToBytes();
        string content = string.Empty;
        byte[] pdf = PdfDocumentObjectGraphRewriter.Rewrite(template, null, null, (objects, security) => {
            PdfDictionary page = Assert.Single(objects.Values.Select(value => value.Value).OfType<PdfDictionary>(),
                dictionary => dictionary.Get<PdfName>("Type")?.Name == "Page");
            PdfDictionary Resolve(PdfObject value) => Assert.IsType<PdfDictionary>(
                value is PdfReference reference ? objects[reference.ObjectNumber].Value : value);
            PdfDictionary fonts = Resolve(Resolve(page.Items["Resources"]).Items["Font"]);
            string resource = Assert.Single(fonts.Items).Key;
            var output = new StringBuilder();
            double next = 40D + command.AdvanceWidth1000!.Value * 10D / 1000D * scale / 100D
                + command.TotalJustificationAdvance + 2D;
            new ContentStreamBuilder(output).BeginText().Font(resource, 10D).HorizontalTextScaling(scale)
                .TextMatrix(40D, 400D).ShowText(command, 10D).TextMatrix(next, 400D).ShowText(probe, 10D).EndText();
            content = output.ToString();
            PdfReference contents = Assert.IsType<PdfReference>(page.Items["Contents"]);
            objects[contents.ObjectNumber] = new PdfIndirectObject(contents.ObjectNumber, contents.Generation,
                new PdfStream(new PdfDictionary(), Encoding.ASCII.GetBytes(content)));
            return security.InfoObjectNumber;
        });
        return (pdf, content);
    }

    private static double[] SavedOrigins(byte[] bytes, string name) {
        string? retained = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_ENGINE_PROOF_OUTPUT");
        string directory = string.IsNullOrWhiteSpace(retained) ? Path.GetTempPath() : retained!;
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "signed-glyph-" + name + (string.IsNullOrWhiteSpace(retained) ? "-" + Guid.NewGuid().ToString("N") : "") + ".pdf");
        try {
            File.WriteAllBytes(path, bytes);
            using var pdf = PdfPigDocument.Open(path);
            var letters = pdf.GetPage(1).Letters;
            Assert.Equal("ABCX", string.Concat(letters.Select(letter => letter.Value)));
            return letters.Select(letter => letter.StartBaseLine.X).ToArray();
        } finally {
            if (string.IsNullOrWhiteSpace(retained)) File.Delete(path);
        }
    }
}
