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
