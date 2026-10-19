using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

/// <summary>Default text output: pair kerning, run-level show operators, compact subsets, bidi and line breaking.</summary>
public class PdfTextQualityTests {
    [Theory]
    [InlineData(0xFB00, "ff")]
    [InlineData(0xFB01, "fi")]
    [InlineData(0xFB02, "fl")]
    [InlineData(0xFB03, "ffi")]
    [InlineData(0xFB04, "ffl")]
    [InlineData(0xFB05, "st")]
    [InlineData(0xFB06, "st")]
    public void CanonicalLigatureGlyphTextRetainsItsSourceLetters(int scalar, string expected) {
        var canonical = new PdfCanonicalGlyphText(new System.Collections.Generic.Dictionary<int, int> { [scalar] = 1 });
        Assert.True(canonical.TryGet(1, out string text));
        Assert.Equal(expected, text);
    }

    [Fact]
    public void HardBreaksResolveEachParagraphDirectionIndependently() {
        const string later = "123 - abc";
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs("אבג".Concat(later).Select(character => (int)character).Distinct().ToArray());
        byte[] Create(string text) => PdfDocument.Create(new PdfOptions().UseFontFamily(new PdfEmbeddedFontFamily("Bidi", font)))
            .Paragraph(paragraph => paragraph.Text(text)).ToBytes();
        var combined = PdfReadDocument.Open(Create("אבג\n" + later)).Pages[0].GetTextSpans();
        var standalone = PdfReadDocument.Open(Create(later)).Pages[0].GetTextSpans();
        PdfTextSpan word = Assert.Single(combined, span => span.Text.Contains("abc", StringComparison.Ordinal));
        PdfTextSpan expected = Assert.Single(standalone, span => span.Text.Contains("abc", StringComparison.Ordinal));
        Assert.Equal(expected.X, word.X, 5);
    }

    [Fact]
    public void ScalarRightToLeftOutputHasTheSamePaintOrderWithAndWithoutFontUnitLayout() {
        const string text = "אבג";
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(text.Select(character => (int)character).ToArray());
        byte[] Glyphs(bool fontUnits) {
            byte[] pdf = PdfDocument.Create(new PdfOptions { UseFontUnitTextLayout = fontUnits, TextShapingMode = PdfTextShapingMode.UnicodeScalar }
                    .UseFontFamily(new PdfEmbeddedFontFamily("RTL", font)))
                .Paragraph(paragraph => paragraph.Text(text)).ToBytes();
            return PdfReadDocument.Open(pdf).Pages[0].GetTextSpans().SelectMany(span => span.GlyphBytes ?? Array.Empty<byte[]>()).SelectMany(bytes => bytes).ToArray();
        }
        byte[] expected = Glyphs(true);
        Assert.NotEmpty(expected);
        Assert.Equal(expected, Glyphs(false));
    }

    [Fact]
    public void StandardFontTextUsesAfmPairKerningUnlessDisabled() {
        double Width(bool kerning) {
            byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = false, TextKerning = kerning })
                .Paragraph(p => p.Text("AVATAR")).ToBytes();
            using var reader = PdfPigDocument.Open(pdf);
            var letters = reader.GetPage(1).Letters;
            return letters.Last().EndBaseLine.X - letters.First().StartBaseLine.X;
        }
        // Helvetica AFM: A-V -70, V-A -80, T-A -120 (and A-T -120) in 1/1000 em at 11 pt.
        Assert.Equal(-(70 + 80 + 120 + 120) * 11D / 1000D, Width(true) - Width(false), 2);
    }

    [Fact]
    public void EmbeddedFontTextUsesFontKerningInLayoutAndPaint() {
        byte[] font = ManagedTextShapingTestAssets.CreateFontWithKerning('A', 'V', -200, includeSpace: true);
        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = false }.EmbedStandardFont(PdfStandardFont.Helvetica, font, "Kerned"))
            .Paragraph(p => p.Text("AV AV")).ToBytes();
        string content = Encoding.ASCII.GetString(pdf);
        Assert.Matches(@"\[<0001> 200<0002", content);
        using var reader = PdfPigDocument.Open(pdf);
        var letters = reader.GetPage(1).Letters.Where(static letter => letter.Value == "V").ToArray();
        Assert.Equal(2, letters.Length);
    }

    [Fact]
    public void EmbeddedParagraphLinesUseOneShowOperatorWithoutPerGlyphTextObjects() {
        byte[] font = File.ReadAllBytes(PdfComplianceTestFonts.FindBundledTrueTypeFont()!);
        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = false }.UseFontFamily(new PdfEmbeddedFontFamily("Line", font)))
            .Paragraph(p => p.Text("Office files fly: affine kerning AVATAR, To, Ty and Wo stay compact."))
            .ToBytes();
        string raw = Encoding.ASCII.GetString(pdf);
        string content = raw.Substring(raw.IndexOf("BT\n", StringComparison.Ordinal));
        content = content.Substring(0, content.IndexOf("ET\n", StringComparison.Ordinal));
        Assert.Equal(1, Regex.Matches(content, @"\bT[jJ]\n").Count);
        Assert.DoesNotContain("ActualText", content, StringComparison.Ordinal);
        Assert.Equal("Office files fly: affine kerning AVATAR, To, Ty and Wo stay compact.",
            PdfReadDocument.Open(pdf).ExtractText().Trim());
    }

    [Fact]
    public void EmbeddedTrueTypeSubsetIsRenumberedTaggedAndMappedThroughCidToGidMap() {
        byte[] font = File.ReadAllBytes(PdfComplianceTestFonts.FindBundledTrueTypeFont()!);
        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressEmbeddedFonts = false }.UseFontFamily(new PdfEmbeddedFontFamily("Subset", font)))
            .Paragraph(p => p.Text("Subset tag"))
            .ToBytes();
        string raw = Encoding.ASCII.GetString(pdf);
        Match baseFont = Regex.Match(raw, @"/BaseFont /([A-Z]{6})\+Subset-Regular");
        Assert.True(baseFont.Success);
        Assert.Contains("/FontName /" + baseFont.Groups[1].Value + "+Subset-Regular", raw, StringComparison.Ordinal);
        Assert.Matches(@"/CIDToGIDMap \d+ 0 R", raw);
        PdfFontInfo inspected = Assert.Single(PdfDocument.Load(pdf).Resources.Fonts().Fonts, static candidate => candidate.IsEmbedded);
        PdfOpenTypeFontInfo program = Assert.IsType<PdfOpenTypeFontInfo>(inspected.EmbeddedOpenTypeInfo);
        Assert.InRange(program.GlyphCount, 2, 12);
        using var reader = PdfPigDocument.Open(pdf);
        Assert.Equal("Subset tag", reader.GetPage(1).Text);
    }

    [Fact]
    public void RepeatedSavesProduceIdenticalSubsetTags() {
        byte[] font = File.ReadAllBytes(PdfComplianceTestFonts.FindBundledTrueTypeFont()!);
        string Tag() {
            byte[] pdf = PdfDocument.Create(new PdfOptions().UseFontFamily(new PdfEmbeddedFontFamily("Stable", font)))
                .Paragraph(p => p.Text("Deterministic output")).ToBytes();
            return Regex.Match(Encoding.ASCII.GetString(pdf), @"/BaseFont /([A-Z]{6}\+)").Groups[1].Value;
        }
        string first = Tag();
        Assert.Equal(7, first.Length);
        Assert.Equal(first, Tag());
    }

    [Fact]
    public void ArabicParagraphIsShapedAndWordsArePlacedRightToLeft() {
        string? fontPath = FindRepositoryFile("Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/NotoSansArabic-Regular.ttf");
        Assert.NotNull(fontPath);
        const string text = "مرحبا بالعالم";
        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = false }
                .UseFontFamily(new PdfEmbeddedFontFamily("Arabic", File.ReadAllBytes(fontPath!))))
            .Paragraph(p => p.Text(text))
            .ToBytes();

        Assert.Equal(text, PdfReadDocument.Open(pdf).ExtractText().Trim());
        var spans = PdfReadDocument.Open(pdf).Pages[0].GetTextSpans().Where(static span => span.Text.Trim().Length > 0).ToArray();
        // Spans carry glyphs in paint (visual) order, as from other right-to-left producers.
        static string Visual(string logical) => new string(logical.Reverse().ToArray());
        var first = Assert.Single(spans, span => span.Text.Contains(Visual("مرحبا"), StringComparison.Ordinal));
        var second = Assert.Single(spans, span => span.Text.Contains(Visual("بالعالم"), StringComparison.Ordinal));
        // Logical first word stays first in content order but is painted to the right.
        Assert.True(first.X > second.X);
        Assert.True(Array.IndexOf(spans, first) < Array.IndexOf(spans, second));
        Assert.DoesNotContain("/ActualText", PdfTestContentText.ReadDecoded(pdf), StringComparison.Ordinal);
        // Joined (contextual) forms replace the isolated letters.
        var program = PdfTrueTypeFontProgram.Parse(File.ReadAllBytes(fontPath!));
        Assert.True(program.TryGetGlyphId(0x0645, out int isolatedMeem));
        string isolatedCode = isolatedMeem.ToString("X4");
        var shownCodes = Regex.Matches(string.Concat(Regex.Matches(Encoding.ASCII.GetString(pdf), @"stream\r?\n([\s\S]*?)endstream").Cast<Match>().Select(static m => m.Groups[1].Value).Where(static content => content.Contains(" Tf", StringComparison.Ordinal) && Regex.IsMatch(content, @"\bBT\b"))), "<([0-9A-F]+)>")
            .Cast<Match>()
            .SelectMany(static m => Enumerable.Range(0, m.Groups[1].Value.Length / 4).Select(i => m.Groups[1].Value.Substring(i * 4, 4)))
            .ToArray();
        Assert.NotEmpty(shownCodes);
        Assert.DoesNotContain(isolatedCode, shownCodes);
    }

    [Fact]
    public void WordsWrapAtLineBreakOpportunitiesWithinTheRemainingWidth() {
        byte[] pdf = PdfDocument.Create(new PdfOptions { PageWidth = 130, MarginLeft = 20, MarginRight = 20 })
            .Paragraph(p => p.Text("Use a state-of-the-art approach today"))
            .ToBytes();
        string[] lines = PdfReadDocument.Open(pdf).ExtractText().Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
        Assert.Contains(lines, static line => line.TrimEnd().EndsWith("-", StringComparison.Ordinal));
        Assert.Equal("Use a state-of-the-art approach today", string.Join(" ", lines.Select(static line => line.Trim())).Replace("- ", "-"));
    }

    [Fact]
    public void AutomaticHyphenationSplitsWordsAtLineEnds() {
        PdfOptions options = new PdfOptions { PageWidth = 160, MarginLeft = 20, MarginRight = 20 }.UseAutomaticHyphenation("en-US");
        byte[] pdf = PdfDocument.Create(options)
            .Paragraph(p => p.Text("The representation of international communication"))
            .ToBytes();
        string[] lines = PdfReadDocument.Open(pdf).ExtractText().Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
        Assert.Contains(lines, static line => Regex.IsMatch(line.TrimEnd(), @"\p{L}-$"));
    }

    private static string? FindRepositoryFile(string relativePath) {
        string candidate = Path.Combine(RepositoryTestPaths.Find(), relativePath.Replace('/', Path.DirectorySeparatorChar));
        return File.Exists(candidate) ? candidate : null;
    }
}
