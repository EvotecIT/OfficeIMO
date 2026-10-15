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
    public void DistinctGlyphRunsShareOneShowOperatorWithoutActualText() {
        byte[] data = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A', 'B');
        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = false }
                .UseFontFamily(new PdfEmbeddedFontFamily("Distinct", data)))
            .Paragraph(p => p.Text("AB BA AB"))
            .ToBytes();
        string raw = System.Text.Encoding.ASCII.GetString(pdf);
        string content = raw.Substring(raw.IndexOf("BT\n", StringComparison.Ordinal));
        content = content.Substring(0, content.IndexOf("ET\n", StringComparison.Ordinal));
        Assert.DoesNotContain("ActualText", content, StringComparison.Ordinal);
        Assert.Equal(1, content.Split(new[] { " Tj\n" }, StringSplitOptions.None).Length - 1 +
            content.Split(new[] { " TJ\n" }, StringSplitOptions.None).Length - 1);
        using var independent = UglyToad.PdfPig.PdfDocument.Open(pdf);
        Assert.Equal("ABBAAB", string.Concat(independent.GetPage(1).Letters.Select(l => l.Value)).Replace(" ", string.Empty));
    }

    [Fact]
    public void SharedNominalGlyphCharactersUseDeterministicAliasCodes() {
        var program = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFont(' ', 'A', 'B'));
        IPdfGlyphTextPinner pinner = program.TextPinner;
        Assert.True(pinner.TryResolveCode(1, "B", out int b));
        Assert.True(pinner.TryResolveCode(1, " ", out int space));
        Assert.True(pinner.TryResolveCode(1, "A", out int a));
        Assert.Equal(1, space);
        Assert.Equal(program.GlyphCount + 'A', a);
        Assert.Equal(program.GlyphCount + 'B', b);
        Assert.False(pinner.TryResolveCode(1, "AB", out _));

        var fork = program.ForkForDocument();
        Assert.True(fork.TextPinner.TryResolveCode(1, "A", out int forkA));
        Assert.Equal(a, forkA);
        byte[] map = program.BuildCidToGlyphMap()!;
        Assert.Equal((program.GlyphCount + 'B' + 1) * 2, map.Length);
        Assert.Equal(1, map[(program.GlyphCount + 'A') * 2 + 1]);
        program.ResetGlyphUsage();
        Assert.Null(program.BuildCidToGlyphMap());
    }

    [Fact]
    public void UnmappedGlyphKeepsItsFirstPinnedText() {
        var program = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateFontWithEmptyGlyphs(4, 'A'));
        Assert.True(program.TextPinner.TryResolveCode(3, "fi", out int code));
        Assert.Equal(3, code);
        Assert.True(program.TextPinner.TryResolveCode(3, "fi", out _));
        Assert.False(program.TextPinner.TryResolveCode(3, "fl", out _));
        Assert.False(program.TextPinner.TryResolveCode(3, string.Empty, out _));
    }
}
