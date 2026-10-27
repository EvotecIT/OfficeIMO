using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public class PdfDefaultFontEmbeddingTests {
    private static bool HasInstalledDefaultSansFont() =>
        PdfCreateOptions.EnumerateDefaultFontFamilyCandidates(PdfStandardFont.Helvetica)
            .Any(family => PdfEmbeddedFontFamily.TryFromSystem(family, out _));

    private static PdfDocument CreateStyledDocument(PdfCreateOptions? options = null) {
        return PdfDocument.Create(options).Compose(c => c.Page(page => page.Content(content => content
            .H1("Quarterly review")
            .Paragraph(p => p.Text("Regular, ").Bold("bold").Text(" and ").Italic("italic").Text(" text.")))));
    }

    [Fact]
    public void DefaultNativeOutputEmbedsTheThemeFontWithBoldAndItalicFaces() {
        if (!HasInstalledDefaultSansFont()) return;
        byte[] bytes = CreateStyledDocument(new PdfCreateOptions { EmbedDefaultFonts = true }).ToBytes();
        string pdf = PdfEncoding.Latin1GetString(bytes);

        Assert.DoesNotContain("/BaseFont /Helvetica", pdf);
        Assert.Contains("/FontFile2", pdf);
        using var document = PdfPigDocument.Open(bytes);
        var fonts = document.GetPages().SelectMany(page => page.Letters).Select(letter => letter.FontName).Distinct().ToArray();
        Assert.Contains(fonts, name => name.Contains("Bold", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(fonts, name => name.Contains("Italic", StringComparison.OrdinalIgnoreCase) || name.Contains("Oblique", StringComparison.OrdinalIgnoreCase));
        Assert.All(fonts, name => Assert.Matches("^[A-Z]{6}\\+", name));
    }

    [Fact]
    public void BoardroomHeadingsUseTheirOwnFamilyAndBodyKeepsItsFont() {
        var options = new PdfCreateOptions { Theme = OfficeTheme.Boardroom, EmbedDefaultFonts = true };
        PdfHeadingStyle heading = PdfTheme.FromOfficeTheme(OfficeTheme.Boardroom).HeadingStyles!.Level1!;
        Assert.Equal(OfficeTheme.Boardroom.Fonts.Heading.Family, heading.FontFamily);
        Assert.Equal(PdfStandardFont.TimesRoman, heading.Font);
        options.ApplyDefaultFontEmbedding();
        if (!options.HasNamedFontFamily(heading.FontFamily) || !HasInstalledDefaultSansFont()) return;

        using PdfDocument document = PdfDocument.Create(options).Compose(composition => composition.Content(content =>
            content.Title("Executive").H1("Quarterly review").Paragraph(p => p.Text("Regular, body text."))));
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var words = pdf.GetPage(1).GetWords().ToArray();
        string headingFont = Assert.Single(words, word => word.Text == "Quarterly").Letters[0].FontName;
        string bodyFont = Assert.Single(words, word => word.Text == "Regular,").Letters[0].FontName;
        string titleFont = Assert.Single(words, word => word.Text == "Executive").Letters[0].FontName;

        Assert.NotEqual(headingFont.Split('+').Last(), bodyFont.Split('+').Last());
        Assert.Contains(options.NamedFontFamilies.Values, family => headingFont.Contains(
            family.FamilyName.Replace(" ", string.Empty), StringComparison.OrdinalIgnoreCase) && titleFont.Contains(
            family.FamilyName.Replace(" ", string.Empty), StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void OptingOutWritesStandardFontsWithoutDiagnostics() {
        var report = new PdfConversionReport();
        var options = new PdfCreateOptions { EmbedDefaultFonts = false }.ReportDiagnosticsTo(report);
        byte[] bytes = CreateStyledDocument(options).ToBytes();
        string pdf = PdfEncoding.Latin1GetString(bytes);

        Assert.Contains("/BaseFont /Helvetica", pdf);
        Assert.DoesNotContain("/FontFile2", pdf);
        Assert.DoesNotContain(report.Warnings, w => w.Code == OfficeIMO.OfficeDiagnosticCodes.FontNotEmbedded);
        Assert.False(options.Clone().EmbedDefaultFonts);
    }

    [Fact]
    public void SectionsShareOneEmbeddedSubsetPerFace() {
        if (!HasInstalledDefaultSansFont()) return;
        PdfDocument document = PdfDocument.Create(new PdfCreateOptions { EmbedDefaultFonts = true });
        for (int index = 0; index < 4; index++) {
            int section = index;
            document.Section(page => page.Content(content => content.H2("Section " + section).Text("Body text for section " + section)));
        }
        byte[] bytes = document.ToBytes();
        string pdf = PdfEncoding.Latin1GetString(bytes);
        // Subsets carry only the glyphs the sections painted, not the whole installed face.
        Assert.True(bytes.Length < 100_000, "Generated PDF is " + bytes.Length + " bytes.");

        string[] baseFonts = Regex.Matches(pdf, @"/Type /Font\s+/Subtype /Type0\s+/BaseFont /([A-Z]{6}\+[^\s/]+)")
            .Cast<Match>().Select(match => match.Groups[1].Value).ToArray();
        Assert.NotEmpty(baseFonts);
        string[] faces = baseFonts.Select(name => name.Substring(7)).ToArray();
        Assert.Equal(faces.Distinct().Count(), faces.Length);
    }

    [Fact]
    public void UnusedFontProgramsContributeNoGlyphsWhenSectionsMergeUsage() {
        string path = Path.Combine(AppContext.BaseDirectory, "Typography", "Carlito-Regular.ttf");
        if (!File.Exists(path)) return;
        PdfTrueTypeFontProgram program = PdfTrueTypeFontProgram.Parse(File.ReadAllBytes(path), null);
        Assert.NotEmpty(program.GetGlyphToUnicodeMappings());
        Assert.Empty(program.GetUsedGlyphToUnicodeMappings());
        Assert.True(program.TryGetGlyphId('A', out int glyphId));
        program.RecordGlyphUsage(glyphId, "A");
        Assert.Single(program.GetUsedGlyphToUnicodeMappings());
    }

    [Fact]
    public void FallbackSetsKeepTheirSlotsWhenDefaultFontsAreEmbedded() {
        if (!HasInstalledDefaultSansFont()) return;
        var fallbacks = new PdfEmbeddedFontFallbackSet(
            new[] { new PdfEmbeddedFontFallbackCandidate("Emoji Fallback",
                ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 0x1F600)) },
            new[] { PdfStandardFont.TimesRoman });
        var options = new PdfCreateOptions { EmbedDefaultFonts = true }.RegisterEmbeddedFontFallbacks(fallbacks);
        byte[] bytes = PdfDocument.Create(options).Compose(c => c.Page(page => page.Content(content => content
            .Paragraph(p => p.Text("Plain text"))
            .Paragraph(p => p.Text("A\U0001F600B"))))).ToBytes();

        string extracted = PdfReadDocument.Open(bytes).ExtractText();
        Assert.Contains("Plain text", extracted, StringComparison.Ordinal);
        Assert.Contains("A\U0001F600B", extracted, StringComparison.Ordinal);
        Assert.Contains("Emoji", PdfEncoding.Latin1GetString(bytes), StringComparison.Ordinal);
    }

    [Fact]
    public void ExplicitlyRegisteredFontsAreNotReplaced() {
        string path = Path.Combine(AppContext.BaseDirectory, "Typography", "Carlito-Regular.ttf");
        if (!File.Exists(path)) return;
        var options = new PdfCreateOptions().UseFontFamily("Carlito", path);
        byte[] bytes = CreateStyledDocument(options).ToBytes();
        using var document = PdfPigDocument.Open(bytes);
        Assert.All(document.GetPages().SelectMany(page => page.Letters), letter => Assert.Contains("Carlito", letter.FontName));
    }
}
