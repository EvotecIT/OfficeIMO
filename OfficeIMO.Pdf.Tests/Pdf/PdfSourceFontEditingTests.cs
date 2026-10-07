using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfSourceFontEditingTests {
    [Theory]
    [InlineData("reportlab")]
    [InlineData("rc4-40")]
    [InlineData("aes-128")]
    [InlineData("aes-256")]
    public void IndependentProducerFontsRemainEmbeddedAfterPartialReplacement(string producer) {
        string root = FindRepositoryRoot();
        bool encrypted = producer != "reportlab";
        var options = new PdfLoadOptions { Password = encrypted ? "owner" : null };
        PdfDocument source = PdfDocument.Load(File.ReadAllBytes(Path.Combine(root, "OfficeIMO.TestAssets", "PdfEditing", "source-font-" + producer + ".pdf")), options);
        PdfFontInfo before = Assert.Single(source.Resources.Fonts(new PdfFontInspectionOptions { IncludeEmbeddedProgramBytes = true }).Fonts, font => font.IsEmbedded);
        PdfTextEditResult edit = source.Text.ReplaceAll("alpha", "gamma");
        Assert.Empty(edit.Warnings);
        PdfFontInfo after = Assert.Single(edit.Document.Resources.Fonts(new PdfFontInspectionOptions { IncludeEmbeddedProgramBytes = true }).Fonts, font => font.IsEmbedded);
        Assert.Equal(before.BaseFontName, after.BaseFontName);
        Assert.Equal(before.EmbeddedProgramBytes, after.EmbeddedProgramBytes);
        Assert.Contains("gamma beta gamma Żółć", edit.Document.Reader.Text());
        PdfReadDocument read = PdfReadDocument.Open(edit.Document.ToBytes(), options);
        Assert.Equal(encrypted, read.Security.HasEncryption);
        Assert.Equal(PdfReadDocument.Open(source.ToBytes(), options).Security.EncryptionRevision, read.Security.EncryptionRevision);
        Assert.All(read.Pages[0].GetTextSpans(), span => {
            Assert.True(span.TryGetGlyphs(out IReadOnlyList<PdfTextGlyph> glyphs));
            Assert.Equal(span.Text, string.Concat(glyphs.Select(glyph => glyph.Text)));
        });
        if (encrypted) Assert.Throws<PdfPasswordRequiredException>(() => PdfReadDocument.Open(edit.Document.ToBytes()));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PartialReplacementRetainsEmbeddedFontAndNeighborText(bool encrypted) {
        PdfDocument source = CreateSource(encrypted);
        PdfFontInfo before = Assert.Single(source.Resources.Fonts(new PdfFontInspectionOptions { IncludeEmbeddedProgramBytes = true }).Fonts);
        PdfTextEditResult edit = source.Text.ReplaceAll("alpha", "gamma");
        Assert.Equal(1, edit.AffectedCount);
        Assert.Empty(edit.Warnings);
        PdfFontInfo after = Assert.Single(edit.Document.Resources.Fonts(new PdfFontInspectionOptions { IncludeEmbeddedProgramBytes = true }).Fonts);
        Assert.Equal(before.BaseFontName, after.BaseFontName);
        Assert.Equal(before.EmbeddedProgramBytes, after.EmbeddedProgramBytes);
        Assert.Equal(before.ToUnicodeMappingCount, after.ToUnicodeMappingCount);
        Assert.Contains("gamma beta gamma", edit.Document.Reader.Text());
        Assert.DoesNotContain("alpha", edit.Document.Reader.Text());
        Assert.Equal(encrypted, PdfReadDocument.Open(edit.Document.ToBytes(), new PdfLoadOptions { Password = encrypted ? "owner" : null }).Security.HasEncryption);
        Assert.All(PdfReadDocument.Open(edit.Document.ToBytes(), new PdfLoadOptions { Password = encrypted ? "owner" : null }).Pages[0].GetTextSpans(),
            span => Assert.Equal(before.BaseFontName, span.BaseFont));
    }

    [Fact]
    public void LocatedMovePreservesSourceFontAndUnmovedOrigins() {
        PdfDocument source = CreateSource(false);
        PdfTextSpan original = Assert.Single(PdfReadDocument.Open(source.ToBytes()).Pages[0].GetTextSpans());
        PdfTextMatch match = Assert.Single(source.Text.Find("beta"));
        PdfTextEditResult moved = source.Text.Move(match, 8, -24);
        Assert.Empty(moved.Warnings);
        PdfTextSpan[] spans = PdfReadDocument.Open(moved.Document.ToBytes()).Pages[0].GetTextSpans().ToArray();
        Assert.All(spans, span => Assert.Equal(original.BaseFont, span.BaseFont));
        PdfTextSpan beta = Assert.Single(spans, span => span.Text == "beta");
        Assert.Equal(original.Y - 24, beta.Y, 3);
        PdfTextSpan alpha = Assert.Single(spans, span => span.Text.Contains("alpha", StringComparison.Ordinal));
        Assert.Equal(original.X, alpha.X, 3);
        Assert.Equal(original.Y, alpha.Y, 3);
        Assert.Single(moved.Document.Resources.Fonts().Fonts);
    }

    [Fact]
    public void MissingSubsetGlyphRequiresExplicitSubstitutionAndLeavesSourceUntouched() {
        PdfDocument source = CreateSource(false);
        byte[] original = source.ToBytes();
        NotSupportedException failure = Assert.Throws<NotSupportedException>(() => source.Text.ReplaceAll("alpha", "zebra"));
        Assert.Contains("existing subset", failure.Message);
        Assert.Equal(original, source.ToBytes());
        PdfTextEditResult substituted = source.Text.ReplaceAll("alpha", "zebra", editOptions: new PdfTextEditOptions { Font = PdfStandardFont.Courier });
        Assert.Contains("zebra", substituted.Document.Reader.Text());
        Assert.Contains(substituted.Warnings, warning => warning.Contains("substituted", StringComparison.Ordinal));
        Assert.Contains(substituted.Document.Resources.Fonts().Fonts, font => font.BaseFontName == "Courier");
    }

    [Theory]
    [InlineData("Helvetica")]
    [InlineData("Courier")]
    [InlineData("Times-Roman")]
    public void EmbeddedFontNamedLikeStandardFontIsStillReused(string baseFontAlias) {
        PdfDocument source = CreateSource(false, baseFontAlias);
        PdfFontInfo before = Assert.Single(source.Resources.Fonts(new PdfFontInspectionOptions { IncludeEmbeddedProgramBytes = true }).Fonts);
        PdfTextEditResult edit = source.Text.ReplaceAll("alpha", "gamma");
        Assert.Empty(edit.Warnings);
        PdfFontInfo after = Assert.Single(edit.Document.Resources.Fonts(new PdfFontInspectionOptions { IncludeEmbeddedProgramBytes = true }).Fonts);
        Assert.Equal(before.EmbeddedProgramBytes, after.EmbeddedProgramBytes);
        Assert.Equal(before.ToUnicodeMappingCount, after.ToUnicodeMappingCount);
        Assert.Equal("Type0", after.Subtype);
        Assert.All(PdfReadDocument.Open(edit.Document.ToBytes()).Pages[0].GetTextSpans(), span => {
            Assert.Equal(baseFontAlias, span.BaseFont);
        });
        Assert.Contains("gamma beta gamma", edit.Document.Reader.Text());
    }

    [Theory]
    [InlineData("\u0628\u0627")]
    [InlineData("\u05D0\u05D1")]
    [InlineData("\u0915\u093F")]
    [InlineData("a\u0301")]
    [InlineData("a\u202Eb")]
    public void CoveredUnicodeMappingDoesNotAuthorizeUnshapedText(string text) {
        string entries = string.Join("\n", text.Select((character, index) =>
            "<" + (index + 1).ToString("X4") + "> <" + ((int)character).ToString("X4") + ">"));
        Assert.True(ToUnicodeCMap.TryParse(PdfEncoding.Latin1GetBytes(text.Length + " beginbfchar\n" + entries + "\nendbfchar"), out ToUnicodeCMap? cmap));
        Assert.True(cmap!.TryEncodeTextCodes(text, out _));
        var resource = new PdfFontResource("F1", "Embedded", "Identity-H", true, cmap, fontSubtype: "Type0");
        var sourceFont = new PdfSourceTextFont(resource, bytes => bytes.Length * 250D);
        NotSupportedException failure = Assert.Throws<NotSupportedException>(() => sourceFont.Encode(text));
        Assert.Contains("shaping", failure.Message);
    }

    [Fact]
    public void ExplicitStandardSubstitutionReportsEmbeddedFontWithMatchingName() {
        PdfDocument source = CreateSource(false, "Helvetica");
        PdfTextEditResult edit = source.Text.ReplaceAll("alpha", "gamma", editOptions: new PdfTextEditOptions { Font = PdfStandardFont.Helvetica });
        Assert.Contains(edit.Warnings, warning => warning.Contains("substituted", StringComparison.Ordinal));
        Assert.Contains("gamma", edit.Document.Reader.Text());
    }

    private static PdfDocument CreateSource(bool encrypted, string? baseFontAlias = null) {
        string root = FindRepositoryRoot();
        byte[] font = File.ReadAllBytes(Path.Combine(root, "OfficeIMO.TestAssets", "Fonts", "OfficeIMOBaselineSans-Regular.ttf"));
        var options = new PdfOptions().RegisterFontFamily(PdfStandardFont.Helvetica, new PdfEmbeddedFontFamily("Baseline", font));
        if (encrypted) options.SetEncryption(new PdfStandardEncryptionOptions("open") { OwnerPassword = "owner", Algorithm = PdfStandardEncryptionAlgorithm.Aes256 });
        byte[] source = PdfDocument.Create(pdf => pdf.Content(content => content.Paragraph(p => p.Text("alpha beta gamma"))), options).ToBytes();
        var loadOptions = new PdfLoadOptions { Password = encrypted ? "owner" : null };
        PdfReadDocument read = PdfReadDocument.Open(source, loadOptions);
        // A conventional producer's single Tj run isolates font reuse from positioned-glyph shaping.
        PdfFontResourceSet fonts = new PdfFontResourceCache().GetOrCreate(read.Pages[0].GetFontInspectionResources(), read.Objects);
        PdfFontResource resource = Assert.Single(fonts.Fonts.Values);
        Assert.True(resource.CMap!.TryEncodeText("alpha beta gamma", out string hex));
        int pageNumber = read.Pages[0].ObjectNumber;
        source = PdfDocumentObjectGraphRewriter.Rewrite(source, loadOptions, null, (objects, security) => {
            if (baseFontAlias is not null) {
                foreach (PdfIndirectObject item in objects.Values) {
                    if (item.Value is PdfDictionary dictionary && dictionary.Items.ContainsKey("BaseFont")) {
                        dictionary.Items["BaseFont"] = new PdfName(baseFontAlias);
                    }
                }
            }
            int streamNumber = objects.Keys.Max() + 1;
            objects[streamNumber] = new PdfIndirectObject(streamNumber, 0, new PdfStream(new PdfDictionary(),
                PdfEncoding.Latin1GetBytes("q BT /" + resource.ResourceName + " 14 Tf 1 0 0 1 50 700 Tm <" + hex + "> Tj ET Q\n")));
            ((PdfDictionary)objects[pageNumber].Value).Items["Contents"] = new PdfReference(streamNumber, 0);
            return security.InfoObjectNumber;
        });
        return PdfDocument.Load(source, new PdfLoadOptions { Password = encrypted ? "owner" : null });
    }

    private static string FindRepositoryRoot() =>
        RepositoryTestPaths.Find();
}
