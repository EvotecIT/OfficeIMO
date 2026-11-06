using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfLineGroupingGutterTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NarrowProseAnchorsDoNotSplitANearbyNumericRow(bool glyphFragments) {
        byte[] pdf = PdfDocument.Create(document => document.Content(content => content.AddParagraph("placeholder"))).ToBytes();
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];
        var context = new PdfUnderstandingPageContext(page, 1, new PdfTextLayoutOptions(), 10_000, 1_000);
        var words = new List<PdfUnderstandingWord>();
        foreach (double baseline in new[] { 700D, 680D }) {
            Add("one", 188D, 16.68D, baseline);
            Add("two", 207.46D, 16.67D, baseline);
            Add("six", 226.91D, 12.78D, baseline);
            Add("red", 250D, 14.45D, baseline);
            Add("and", 267.23D, 16.68D, baseline);
            Add("tan", 286.69D, 13.90D, baseline);
        }
        Add("61", 185D, 11.12D, 660D);
        Add("497.00", 199D, 33.36D, 660D);
        Add("8425.00", 250D, 38.92D, 660D);

        IReadOnlyList<PdfUnderstandingLine> lines = PdfUnderstandingPipelineOptions.Structured().LineGrouping!.GroupLines(context, words);

        Assert.Equal("61 497.00 8425.00", Assert.Single(lines, line => line.BaselineY == 660D).Text);

        void Add(string text, double x, double width, double baseline) {
            if (!glyphFragments) { words.Add(Word(text, x, width, baseline)); return; }
            for (int index = 0; index < text.Length; index++)
                words.Add(Word(text[index].ToString(), x + width * index / text.Length, width / text.Length, baseline));
        }
    }

    [Fact]
    public void DistantNarrowProseAnchorsDoNotSplitJustifiedTextAcrossALocalObstruction() {
        byte[] pdf = PdfDocument.Create(document => document.Content(content => content.AddParagraph("placeholder"))).ToBytes();
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];
        var context = new PdfUnderstandingPageContext(page, 1, new PdfTextLayoutOptions(), 10_000, 1_000);
        var words = new List<PdfUnderstandingWord>();
        foreach (double baseline in new[] { 700D, 680D, 500D }) {
            words.Add(Word("one", 188D, 16.68D, baseline));
            words.Add(Word("two", 207.46D, 16.67D, baseline));
            words.Add(Word("six", 226.91D, 12.78D, baseline));
            words.Add(Word("red", 250D, 14.45D, baseline));
            words.Add(Word("and", 267.23D, 16.68D, baseline));
            words.Add(Word("tan", 286.69D, 13.90D, baseline));
        }
        words.Add(Word("crossing", 242D, 25D, 480D));

        IReadOnlyList<PdfUnderstandingLine> lines = PdfUnderstandingPipelineOptions.Structured().LineGrouping!.GroupLines(context, words);

        Assert.Equal("one two six red and tan", Assert.Single(lines, line => line.BaselineY == 500D).Text);
    }

    private static PdfUnderstandingWord Word(string text, double x, double width, double baseline) =>
        new PdfUnderstandingWord(text, x, x + width, baseline, 10D, 0D, Array.Empty<PdfTextSpan>());
}
