using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using Letter = UglyToad.PdfPig.Content.Letter;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfWordSpaceExpansionTests {
    [Theory]
    [InlineData(PdfTextWhitespaceMode.Preserve, 1.2D)]
    [InlineData(PdfTextWhitespaceMode.Collapse, 1.2D)]
    [InlineData(PdfTextWhitespaceMode.Preserve, 12D)]
    [InlineData(PdfTextWhitespaceMode.Collapse, 12D)]
    public void SourceSpaceAdvancesOwnPreferredExpansionBeforeEqualRemainder(PdfTextWhitespaceMode whitespace, double adjustment) {
        var runs = new[] { new PdfTextRun("Alpha"), new PdfTextRun(" ", fontSize: 12D),
            new PdfTextRun("beta"), new PdfTextRun(" ", fontSize: 24D), new PdfTextRun("gamma") };
        string label = whitespace + "-" + adjustment.ToString("0.0", CultureInfo.InvariantCulture);
        List<Letter> natural = Letters(Render(runs, 300D, false, false, whitespace), label + "-natural");
        double width = Span(natural) + adjustment;
        List<Letter> actual = Letters(Render(runs, width, true, true, whitespace), label + "-native");
        double firstCapacity = Gap(natural, 4, 5) / 2D;
        double secondCapacity = Gap(natural, 8, 9) / 2D;
        double preferred = firstCapacity + secondCapacity;
        double expectedFirst = adjustment < preferred ? adjustment * firstCapacity / preferred
            : firstCapacity + (adjustment - preferred) / 2D;
        double expectedSecond = adjustment - expectedFirst;

        Assert.Equal("Alphabetagamma", string.Concat(actual.Select(letter => letter.Value)));
        Assert.Equal(expectedFirst, Gap(actual, 4, 5) - Gap(natural, 4, 5), 3);
        Assert.Equal(expectedSecond, Gap(actual, 8, 9) - Gap(natural, 8, 9), 3);
        Assert.Equal(width, Span(actual), 3);
        for (int index = 0; index < natural.Count; index++) Assert.Equal(natural[index].Width, actual[index].Width, 3);

        List<Letter> ordinary = Letters(Render(runs, width, false, true, whitespace), label + "-ordinary");
        Assert.Equal(adjustment / 2D, Gap(ordinary, 4, 5) - Gap(natural, 4, 5), 3);
        Assert.Equal(adjustment / 2D, Gap(ordinary, 8, 9) - Gap(natural, 8, 9), 3);
    }

    [Fact]
    public void PreservedSourceSpacesRetainIndividualEqualRemainderShares() {
        var runs = new[] { new PdfTextRun("Alpha"), new PdfTextRun("   ", fontSize: 8D),
            new PdfTextRun("beta"), new PdfTextRun(" ", fontSize: 16D), new PdfTextRun("gamma") };
        List<Letter> natural = Letters(Render(runs, 300D, false, false, PdfTextWhitespaceMode.Preserve), "group-natural");
        const double adjustment = 12D;
        List<Letter> actual = Letters(Render(runs, Span(natural) + adjustment, true, true, PdfTextWhitespaceMode.Preserve), "group-native");
        double firstPreferred = Gap(natural, 4, 5) / 2D, secondPreferred = Gap(natural, 8, 9) / 2D;
        double remainder = adjustment - firstPreferred - secondPreferred;
        Assert.True(remainder > 0D);
        Assert.Equal(firstPreferred + 3D * remainder / 4D, Gap(actual, 4, 5) - Gap(natural, 4, 5), 3);
        Assert.Equal(secondPreferred + remainder / 4D, Gap(actual, 8, 9) - Gap(natural, 8, 9), 3);
        Assert.Equal(Span(natural) + adjustment, Span(actual), 3);
        Assert.Equal("Alphabetagamma", string.Concat(actual.Select(letter => letter.Value)));
    }

    [Fact]
    public void ImportedWordExpansionAfterALeftTabMatchesTheSameIndentedWordField() {
        // Native Word preserves the body-word positions when an equal left-tab
        // field replaces a paragraph indent, including differently sized spaces.
        var runs = new[] { new PdfTextRun("Alpha"), new PdfTextRun(" ", fontSize: 12D),
            new PdfTextRun("beta"), new PdfTextRun(" ", fontSize: 24D),
            new PdfTextRun("gamma"), new PdfTextRun(" ", fontSize: 12D), new PdfTextRun("delta") };
        byte[] Render(bool anchored) {
            var style = new PdfParagraphStyle {
                UseWordJustificationSpacing = true, JustificationMode = PdfJustificationMode.Words,
                JustifyExplicitLineBreaks = true, JustifyLastLine = false,
                LeftIndent = anchored ? 0D : 36D, LineSpacing = PdfLineSpacing.Exactly(36D),
                TextWhitespaceMode = PdfTextWhitespaceMode.Preserve
            };
            if (anchored) style.AddTabStop(36D, PdfTabAlignment.Left);
            return PdfDocument.Create(new PdfOptions {
                PageWidth = 612D, PageHeight = 792D, MarginLeft = 72D, MarginRight = 360D,
                MarginTop = 72D, MarginBottom = 72D, DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 12D, CompressContentStreams = false
            }).Compose(document => document.Content(content => content.Paragraph(paragraph => {
                if (anchored) paragraph.Text("TAG").Tab();
                paragraph.Runs(runs).LineBreak().Text("NEXT");
            }, PdfAlign.Justify, style: style))).ToBytes();
        }
        List<List<Letter>> Lines(byte[] bytes) {
            using var pdf = PdfPigDocument.Open(bytes);
            return Assert.Single(pdf.GetPages()).Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
                .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 2)).OrderByDescending(line => line.Key)
                .Select(line => line.OrderBy(letter => letter.StartBaseLine.X).ToList()).ToList();
        }

        List<List<Letter>> ordinary = Lines(Render(false));
        List<List<Letter>> anchoredLines = Lines(Render(true));
        Assert.Equal(2, ordinary.Count);
        Assert.Equal(2, anchoredLines.Count);
        Assert.Equal("Alphabetagammadelta", string.Concat(ordinary[0].Select(letter => letter.Value)));
        Assert.Equal("TAGAlphabetagammadelta", string.Concat(anchoredLines[0].Select(letter => letter.Value)));
        List<Letter> body = anchoredLines[0].Skip("TAG".Length).ToList();
        Assert.Equal(ordinary[0].Count, body.Count);
        Assert.Equal(108D, body[0].StartBaseLine.X, 3);
        Assert.Equal(252D, body[body.Count - 1].EndBaseLine.X, 3);
        for (int index = 0; index < body.Count; index++) {
            Assert.Equal(ordinary[0][index].Value, body[index].Value);
            Assert.Equal(ordinary[0][index].StartBaseLine.X, body[index].StartBaseLine.X, 3);
            Assert.Equal(ordinary[0][index].EndBaseLine.X, body[index].EndBaseLine.X, 3);
            Assert.Equal(ordinary[0][index].StartBaseLine.Y, body[index].StartBaseLine.Y, 3);
        }
        Assert.Equal(72D, anchoredLines[0][0].StartBaseLine.X, 3);
        Assert.True(anchoredLines[0]["TAG".Length - 1].EndBaseLine.X < body[0].StartBaseLine.X);
        Assert.Equal("NEXT", string.Concat(ordinary[1].Select(letter => letter.Value)));
        Assert.Equal("NEXT", string.Concat(anchoredLines[1].Select(letter => letter.Value)));
        Assert.Equal(108D, ordinary[1][0].StartBaseLine.X, 3);
        Assert.Equal(72D, anchoredLines[1][0].StartBaseLine.X, 3);
        for (int index = 0; index < ordinary[1].Count; index++) {
            Assert.Equal(ordinary[1][index].StartBaseLine.X - 36D, anchoredLines[1][index].StartBaseLine.X, 3);
            Assert.Equal(ordinary[1][index].EndBaseLine.X - 36D, anchoredLines[1][index].EndBaseLine.X, 3);
            Assert.Equal(ordinary[1][index].StartBaseLine.Y, anchoredLines[1][index].StartBaseLine.Y, 3);
        }
    }

    private static byte[] Render(PdfTextRun[] runs, double width, bool native, bool lastLine, PdfTextWhitespaceMode whitespace) =>
        PdfDocument.Create(new PdfOptions { PageWidth = width + 40D, PageHeight = 300D,
            MarginLeft = 20D, MarginRight = 20D, MarginTop = 20D, MarginBottom = 20D,
            DefaultFont = PdfStandardFont.Helvetica, DefaultFontSize = 12D, CompressContentStreams = false })
        .Paragraph(p => p.Runs(runs), PdfAlign.Justify, style: new PdfParagraphStyle {
            UseWordJustificationSpacing = native, JustifyLastLine = lastLine, TextWhitespaceMode = whitespace }).ToBytes();

    private static double Span(IReadOnlyList<Letter> letters) => letters[letters.Count - 1].EndBaseLine.X - letters[0].StartBaseLine.X;
    private static double Gap(IReadOnlyList<Letter> letters, int left, int right) => letters[right].StartBaseLine.X - letters[left].EndBaseLine.X;
    private static List<Letter> Letters(byte[] bytes, string label) {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-compression-evidence");
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "expansion-" + label + ".pdf");
        File.WriteAllBytes(path, bytes);
        using var pdf = PdfPigDocument.Open(path);
        return Assert.Single(pdf.GetPages().SelectMany(page => page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 2)).Select(line => line.OrderBy(letter => letter.StartBaseLine.X).ToList())));
    }
}
