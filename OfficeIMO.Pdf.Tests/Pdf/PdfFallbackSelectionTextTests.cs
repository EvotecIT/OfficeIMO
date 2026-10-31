using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Tests;
using Xunit;

namespace OfficeIMO.Pdf.Tests.Pdf;

public class PdfFallbackSelectionTextTests {
    private static byte[] Font() => File.ReadAllBytes(Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Drawing.Tests", "TestAssets", "OpenSans-Bold.ttf"));

    private static PdfEmbeddedFontFallbackCandidate Candidate(string name, byte[] font, params OfficeFontUnicodeRange[] ranges) =>
        new(name, font, new OfficeFontUnicodeRangeSet(ranges));

    [Fact]
    public void AddCoverageDrivenCandidates_SeparatedElements_SelectsFontsWithoutCombiningThem() {
        byte[] font = Font();
        var available = new[] {
            Candidate("A", font, new OfficeFontUnicodeRange('A', 'A')),
            Candidate("B", font, new OfficeFontUnicodeRange('B', 'B')),
            Candidate("Mark", font, new OfficeFontUnicodeRange(0x301, 0x301))
        };
        var selected = new List<PdfEmbeddedFontFallbackCandidate>();

        int covered = PdfOptions.AddCoverageDrivenCandidates(available, "AB\n\u0301", 6, selected, 0, 3);

        Assert.Equal(3, covered);
        Assert.Equal(new[] { "A", "B", "Mark" }, selected.Select(candidate => candidate.FontName));
        Assert.True(PdfTextDiagnostics.PlanEmbeddedFontFallbackText("AB\n\u0301", selected).IsFullyCovered);
    }

    [Fact]
    public void AddCoverageDrivenCandidates_WhitespaceOnlyFontsDoNotConsumeTheRequiredTextBudget() {
        byte[] font = Font();
        const string requiredText = " ABCDEF";
        var available = new[] { Candidate("Space", font, new OfficeFontUnicodeRange(' ', ' ')) }
            .Concat("ABCDEF".Select(character => Candidate(character.ToString(), font,
                new OfficeFontUnicodeRange(character, character))))
            .ToArray();
        var selected = new List<PdfEmbeddedFontFallbackCandidate>();

        PdfOptions.AddCoverageDrivenCandidates(available, requiredText, 6, selected, 0, 6);

        Assert.Equal(new[] { "A", "B", "C", "D", "E", "F" }, selected.Select(candidate => candidate.FontName));
        Assert.True(PdfTextDiagnostics.PlanEmbeddedFontFallbackText(requiredText, selected).IsFullyCovered);
    }

    [Fact]
    public void AddCoverageDrivenCandidates_WhitespaceLedElement_SelectsLaterCoveringFont() {
        byte[] font = Font();
        var available = new[] {
            Candidate("A", font, new OfficeFontUnicodeRange('A', 'A')),
            Candidate("Space and mark", font, new OfficeFontUnicodeRange(' ', ' '), new OfficeFontUnicodeRange(0x301, 0x301))
        };
        var selected = new List<PdfEmbeddedFontFallbackCandidate>();

        int covered = PdfOptions.AddCoverageDrivenCandidates(available, "A \u0301", 6, selected, 0, 3);

        Assert.Equal(3, covered);
        Assert.Equal(2, selected.Count);
        Assert.Equal(" \u0301", PdfTextDiagnostics.PlanEmbeddedFontFallbackText("A \u0301", selected).Segments[1].Text);
    }

    [Fact]
    public void AddCoverageDrivenCandidates_UncoveredCombiningElement_SelectsFontForWholeElement() {
        byte[] font = Font();
        var available = new[] {
            Candidate("A and B", font, new OfficeFontUnicodeRange('A', 'B')),
            Candidate("B and mark", font, new OfficeFontUnicodeRange('B', 'B'), new OfficeFontUnicodeRange(0x301, 0x301))
        };
        var selected = new List<PdfEmbeddedFontFallbackCandidate>();

        int covered = PdfOptions.AddCoverageDrivenCandidates(available, "AB\u0301", 6, selected, 0, 3);

        Assert.Equal(3, covered);
        Assert.Equal(new[] { "A and B", "B and mark" }, selected.Select(candidate => candidate.FontName));
        Assert.Equal("B\u0301", PdfTextDiagnostics.PlanEmbeddedFontFallbackText("AB\u0301", selected).Segments[1].Text);
    }

    [Fact]
    public void PlanEmbeddedFontFallbackText_FirstScalarCovered_DiagnosesMissingCombiningMark() {
        byte[] font = Font();
        var first = Candidate("B", font, new OfficeFontUnicodeRange('B', 'B'));

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(PdfTextDiagnostics.PlanEmbeddedFontFallbackText("B\u0301", new[] { first }).Diagnostics);

        Assert.Equal(1, diagnostic.Index);
        Assert.Equal("U+0301", diagnostic.CodePoint);
        Assert.Equal("\u0301", diagnostic.Text);
    }

    [Fact]
    public void PlanEmbeddedFontFallbackText_AfterSkippedWhitespace_ReportsLaterMissingMark() {
        byte[] font = Font();
        var candidate = Candidate("Letters and space", font,
            new OfficeFontUnicodeRange('A', 'B'), new OfficeFontUnicodeRange(' ', ' '));

        PdfTextFallbackPlan plan = PdfTextDiagnostics.PlanEmbeddedFontFallbackText("A \u0301B\u0301", new[] { candidate });

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(plan.Diagnostics);
        Assert.Equal("U+0301", diagnostic.CodePoint);
        Assert.Equal(4, diagnostic.Index);
        Assert.False(plan.IsFullyCovered);
        Assert.Throws<InvalidOperationException>(() => plan.ToTextRuns(new[] { PdfStandardFont.Helvetica }));
    }

    [Fact]
    public void PlanEmbeddedFontFallbackText_ScalarsCoveredByDifferentFonts_ExplainsElementCoverage() {
        byte[] font = Font();
        var candidates = new[] {
            Candidate("B", font, new OfficeFontUnicodeRange('B', 'B')),
            Candidate("Mark", font, new OfficeFontUnicodeRange(0x301, 0x301))
        };

        PdfTextEncodingDiagnostic diagnostic = Assert.Single(PdfTextDiagnostics.PlanEmbeddedFontFallbackText("B\u0301", candidates).Diagnostics);

        Assert.Equal(0, diagnostic.Index);
        Assert.Contains("No single embedded font fallback candidate covers the text element", diagnostic.Message);
    }
}
