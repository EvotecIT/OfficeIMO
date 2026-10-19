using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfEmbeddedCffType1RenderingTests {
    [Theory]
    [InlineData("fonts-type1")]
    [InlineData("fonts-type1c")]
    [InlineData("fonts-truetype")]
    public void EmbeddedProgramsPaintTheirOwnGlyphsAtPopplerPositions(string page) {
        OfficeImageExportResult exported = PdfRendererQualitySupport.Export(page + ".pdf");
        Assert.DoesNotContain(exported.Diagnostics, diagnostic =>
            diagnostic.Code is PdfRenderCapabilities.CffFontSubstitutionId or PdfRenderCapabilities.FontSubstitutionId);
        OfficeRasterImage actual = PdfRendererQualitySupport.Decode(exported.Bytes);
        OfficeRasterImage expected = PdfRendererQualitySupport.Baseline(page + ".poppler.png");

        var bands = PdfRendererQualitySupport.CompareInkBands(actual, expected);
        Assert.True(bands.Count >= 5, "Expected every text line of the fixture to paint.");
        foreach (var band in bands) {
            string message = $"{page} band {band.Top}-{band.Bottom}: expected {band.Expected}, actual {band.Actual}";
            Assert.InRange(band.Actual.Left - band.Expected.Left, -2, 2);
            Assert.InRange(band.Actual.Right - band.Expected.Right, -2, 2);
            Assert.InRange(band.Actual.Top - band.Expected.Top, -2, 2);
            Assert.InRange(band.Actual.Bottom - band.Expected.Bottom, -2, 2);
            Assert.True(Math.Abs(band.Actual.Darkness - band.Expected.Darkness) <= band.Expected.Darkness * 0.08D, message);
        }
    }

    [Fact]
    public void Type1ProgramConvertsSubroutinesFlexDivAndSeacToCff() {
        byte[] program = ReadFontProgram("fonts-type1.pdf", "FontFile");
        PdfType1FontConverter.Result? converted = PdfType1FontConverter.TryConvert(program);
        Assert.NotNull(converted);
        PdfCffFontProgram? cff = PdfCffFontProgram.TryParse(converted!.Cff);
        Assert.NotNull(cff);

        Assert.Equal(0.0005D, cff!.FontMatrix[0], 9);
        Assert.Equal("Aacute", converted.BuiltinEncoding[193]);
        Assert.Equal("H", converted.BuiltinEncoding[72]);
        Assert.True(cff.TryGetGlyph("D", out int d));
        Assert.Equal(1440D, cff.GetAdvance(d));
        Assert.True(cff.TryGetGlyph("Aacute", out int aacute));
        Assert.False(cff.IsEmptyGlyph(aacute));
        Assert.True(cff.TryGetGlyph("space", out int space));
        Assert.True(cff.IsEmptyGlyph(space));
    }

    [Fact]
    public void BareCffProgramExposesNamesBuiltinEncodingAndWidths() {
        byte[] program = ReadFontProgram("fonts-type1c.pdf", "FontFile3");
        PdfCffFontProgram? cff = PdfCffFontProgram.TryParse(program);
        Assert.NotNull(cff);
        Assert.False(cff!.IsCidKeyed);
        Assert.True(cff.TryGetGlyph("O", out int o));
        Assert.Equal(760D, cff.GetAdvance(o));
        Assert.Equal(o, cff.BuiltinEncoding['O']);
        Assert.True(cff.TryGetGlyphForCid(o, out int sameGlyph));
        Assert.Equal(o, sameGlyph);
    }

    [Fact]
    public void ExistingOpenTypeCidFixtureNoLongerSubstitutesItsCffOutlines() {
        string path = Path.Combine(VisualBaselineTestSupport.GetTestsProjectRoot(), "Pdf", "Fixtures", "Interoperability", "Redaction", "embedded-cff-rotation-0.pdf");
        OfficeImageExportResult exported = PdfReadDocument.Open(File.ReadAllBytes(path)).Pages[0]
            .ExportImage(OfficeImageExportFormat.Png, new PdfImageExportOptions { TargetDpi = 72 });
        Assert.DoesNotContain(exported.Diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.CffFontSubstitutionId);
    }

    [Fact]
    public void MalformedType1AndCffProgramsAreRejectedWithoutExceptions() {
        byte[] type1 = ReadFontProgram("fonts-type1.pdf", "FontFile");
        byte[] cff = ReadFontProgram("fonts-type1c.pdf", "FontFile3");
        var random = new Random(4711);
        for (int iteration = 0; iteration < 300; iteration++) {
            byte[] mutatedType1 = Mutate(type1, random);
            byte[] mutatedCff = Mutate(cff, random);
            PdfType1FontConverter.Result? converted = PdfType1FontConverter.TryConvert(mutatedType1);
            if (converted != null) PdfCffFontProgram.TryParse(converted.Cff);
            PdfCffFontProgram? parsed = PdfCffFontProgram.TryParse(mutatedCff);
            if (parsed != null) {
                for (int glyph = 0; glyph < parsed.GlyphCount; glyph++) {
                    parsed.GetAdvance(glyph);
                    parsed.IsEmptyGlyph(glyph);
                }
            }
        }
        Assert.Null(PdfType1FontConverter.TryConvert(new byte[] { 0x80, 0x01, 0xFF, 0xFF, 0xFF, 0x7F }));
        Assert.Null(PdfCffFontProgram.TryParse(new byte[] { 1, 0, 4, 4, 0xFF, 0xFF, 4 }));
        Assert.Null(PdfCffFontProgram.TryParse(Array.Empty<byte>()));
    }

    private static byte[] Mutate(byte[] source, Random random) {
        byte[] copy = (byte[])source.Clone();
        int edits = 1 + random.Next(8);
        for (int edit = 0; edit < edits; edit++) {
            int position = random.Next(copy.Length);
            copy[position] = (byte)random.Next(256);
        }
        if (random.Next(4) == 0) Array.Resize(ref copy, random.Next(copy.Length));
        return copy;
    }

    private static byte[] ReadFontProgram(string pdfFileName, string descriptorKey) {
        PdfReadDocument document = PdfReadDocument.Open(File.ReadAllBytes(PdfRendererQualitySupport.CorpusPath(pdfFileName)));
        foreach (PdfIndirectObject indirect in document.Objects.Values) {
            if (indirect.Value is not PdfDictionary dictionary ||
                dictionary.Get<PdfName>("Type")?.Name != "FontDescriptor" ||
                !dictionary.Items.TryGetValue(descriptorKey, out PdfObject? reference) ||
                PdfObjectLookup.Resolve(document.Objects, reference) is not PdfStream stream) continue;
            return StreamDecoder.Decode(stream, document.Objects);
        }
        throw new InvalidOperationException("The fixture has no " + descriptorKey + " program.");
    }
}
