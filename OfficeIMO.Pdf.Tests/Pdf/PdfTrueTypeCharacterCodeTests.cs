using System;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfTrueTypeCharacterCodeTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedGlyphCharactersRemainDistinctAcrossNestedFontContextsAndRepeatedSave(bool named) {
        byte[] data = ManagedTextShapingTestAssets.CreateFont(' ', 'A', 'B', 'C');
        var family = new PdfEmbeddedFontFamily("Shared Glyph", data);
        var options = new PdfOptions { CompressEmbeddedFonts = false, CompressContentStreams = false };
        if (named) options.RegisterNamedFontFamily(family); else options.UseFontFamily(family);
        var document = PdfDocument.Create(options).Panel(parent => parent
            .Paragraph(p => { if (named) p.FontFamily("Shared Glyph"); p.Text("AB"); })
            .Panel(child => child.Paragraph(p => { if (named) p.FontFamily("Shared Glyph"); p.Text("BC"); })));
        for (int save = 0; save < 2; save++) {
            byte[] pdf = document.ToBytes();
            using var independent = UglyToad.PdfPig.PdfDocument.Open(pdf);
            Assert.Equal("ABBC", string.Concat(independent.GetPage(1).Letters.Select(l => l.Value)));
            Assert.Equal("ABBC", string.Concat(PdfReadDocument.Open(pdf).Pages[0].GetTextSpans().Select(s => s.Text)));
        }
    }

    [Fact]
    public void LaterSelfReferentialLigatureCannotChangeOrdinaryCharacterMapping() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithSelfReferentialLigature('f', 'i', includeSpace: true);
        byte[] pdf = PdfDocument.Create(new PdfOptions().EmbedStandardFont(PdfStandardFont.Helvetica, data, "Shared Ligature"))
            .Paragraph(p => p.Text("f"))
            .Paragraph(p => p.Text("fi"))
            .Paragraph(p => p.Text("f"))
            .ToBytes();
        using var independent = UglyToad.PdfPig.PdfDocument.Open(pdf);
        Assert.Equal("ffif", string.Concat(independent.GetPage(1).Letters.Select(l => l.Value)));
        Assert.Equal("ffif", string.Concat(PdfReadDocument.Open(pdf).Pages[0].GetTextSpans().Select(s => s.Text)));
    }

    [Fact]
    public void SharedFontResourceMergesCharacterCodesIntroducedOnLaterPageSnapshots() {
        byte[] data = ManagedTextShapingTestAssets.CreateFont(' ', 'A', 'B', 'C');
        byte[] pdf = PdfDocument.Create(new PdfOptions().UseFontFamily(new PdfEmbeddedFontFamily("Page Glyph", data)))
            .Page(page => page.Content(content => content.Item(item => item.Paragraph(p => p.Text("A")))))
            .Page(page => page.Content(content => content.Item(item => item.Paragraph(p => p.Text("B")))))
            .Page(page => page.Content(content => content.Item(item => item.Paragraph(p => p.Text("C")))))
            .ToBytes();
        using var independent = UglyToad.PdfPig.PdfDocument.Open(pdf);
        var owned = PdfReadDocument.Open(pdf);
        Assert.Equal(3, independent.NumberOfPages);
        for (int page = 1; page <= 3; page++) {
            string expected = ((char)('A' + page - 1)).ToString();
            Assert.Equal(expected, Assert.Single(independent.GetPage(page).Letters).Value);
            Assert.Equal(expected, Assert.Single(owned.Pages[page - 1].GetTextSpans()).Text);
        }
        string raw = System.Text.Encoding.ASCII.GetString(pdf);
        Assert.Equal(1, raw.Split(new[] { " /FontFile2 " }, StringSplitOptions.None).Length - 1);
    }

    [Fact]
    public void CachedRunRestoresCharacterUsageAfterResetAndDocumentForkStaysIndependent() {
        var program = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFont('A', 'B'));
        var options = PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures);
        PdfGlyphRun first = program.ShapeText("AB", options);
        Assert.True(program.TryCreateAsciiTextShowCommand(first, out var command));
        byte[] firstMap = program.BuildCidToGlyphMap()!;
        Assert.Null(program.ForkForDocument().BuildCidToGlyphMap());
        program.ResetGlyphUsage();
        Assert.Null(program.BuildCidToGlyphMap());
        PdfGlyphRun second = program.ShapeText("AB", options);
        Assert.Same(first, second);
        Assert.True(program.TryCreateAsciiTextShowCommand(second, out var replayed));
        Assert.Equal(command.GlyphHex, replayed.GlyphHex);
        Assert.Equal(firstMap, program.BuildCidToGlyphMap());
        Assert.Equal(new[] { "A", "B" }, program.GetUsedAsciiCharacterMappings().Select(m => m.UnicodeText));
    }

    [Theory]
    [InlineData(65441, true)]
    [InlineData(65442, false)]
    [InlineData(65535, false)]
    public void SpareCharacterCodesRespectTheSixteenBitCidSpace(int glyphCount, bool compact) {
        var program = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithEmptyGlyphs(glyphCount, '~'));
        PdfGlyphRun run = program.ShapeText("~", PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.Equal(compact, program.TryCreateAsciiTextShowCommand(run, out var command));
        if (compact) {
            Assert.Equal("FFFF", command.GlyphHex);
            byte[] map = program.BuildCidToGlyphMap()!;
            Assert.Equal(131072, map.Length);
            Assert.Equal(1, map[map.Length - 1]);
        } else Assert.Null(program.BuildCidToGlyphMap());
    }

    [Fact]
    public void NonCanonicalAndSharedClusterRunsRetainLogicalIsolationWithoutPartialAliasUsage() {
        var program = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B'));
        var runs = new[] {
            new PdfGlyphRun(new[] { new PdfGlyphInfo(1, "A", 0, 500), new PdfGlyphInfo(1, "B", 1, 500) }, Array.Empty<PdfTextEncodingDiagnostic>(), preserveGlyphUnicode: true),
            new PdfGlyphRun(new[] { new PdfGlyphInfo(1, "A", 0, 500, 500, 0, 0, 0), new PdfGlyphInfo(2, "B", 1, 500, 500, 0, 0, 0, logicalClusterStart: 0) }, Array.Empty<PdfTextEncodingDiagnostic>(), preserveGlyphUnicode: true),
            new PdfGlyphRun(new[] { new PdfGlyphInfo(1, "A", 0, 500) }, Array.Empty<PdfTextEncodingDiagnostic>(), direction: OfficeTextDirection.RightToLeft, preserveGlyphUnicode: true),
            new PdfGlyphRun(new[] { new PdfGlyphInfo(1, "A", 0, 500, 510, 0, 0) }, Array.Empty<PdfTextEncodingDiagnostic>(), preserveGlyphUnicode: true)
        };
        foreach (PdfGlyphRun run in runs) {
            Assert.False(program.TryCreateAsciiTextShowCommand(run, out _));
            Assert.Null(program.BuildCidToGlyphMap());
            Assert.NotNull(run.ToTextShowCommand().LogicalGlyphs);
        }
    }

    [Fact]
    public void TwoByteAsciiCharacterCodesKeepZeroWordSpacingCountAndNominalWidths() {
        var program = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A'));
        PdfGlyphRun run = program.ShapeText(" A", PdfTextShapingOptions.ForRendering("Test", PdfTextShapingMode.OpenTypeLigatures));
        Assert.True(program.TryCreateAsciiTextShowCommand(run, out var command));
        Assert.Equal(0, command.WordSpaceCount);
        Assert.Equal(run.TotalAdvanceWidth1000, command.AdvanceWidth1000);
        Assert.All(program.GetUsedAsciiCharacterMappings(), mapping => Assert.Equal(500, program.GetGlyphWidth1000(mapping.GlyphId)));
        Assert.Null(command.LogicalGlyphs);
        Assert.Same(run.Glyphs, command.VisualGlyphs);
    }
}
