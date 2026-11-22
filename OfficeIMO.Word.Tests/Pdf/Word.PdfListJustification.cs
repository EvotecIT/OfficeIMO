using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Letter = UglyToad.PdfPig.Content.Letter;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("WordTests")]
public sealed class WordPdfListJustificationTests {
    private const string Filler = "alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau upsilon phi chi psi";

    // A justified list item fills each of its lines to the right margin, as Word does, and leaves
    // the last line natural. A distributed item spreads the last line as well.
    [Theory]
    [InlineData(WordListStyle.Bulleted, 0, WordParagraphAlignment.Both)]
    [InlineData(WordListStyle.Numbered, 0, WordParagraphAlignment.Both)]
    [InlineData(WordListStyle.Bulleted, 1, WordParagraphAlignment.Both)]
    [InlineData(WordListStyle.Numbered, 0, WordParagraphAlignment.Distribute)]
    public void SaveAsPdf_JustifiedListItemFillsItsLines(WordListStyle listStyle, int level, WordParagraphAlignment alignment) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("RIGHTREF").ParagraphAlignment = WordParagraphAlignment.Right;
        WordList list = document.AddList(listStyle);
        list.AddItem("JLIST " + Filler + " JLISTEND", level).ParagraphAlignment = alignment;

        List<List<Letter>> lines = ConvertLines(document);
        double right = lines[0].Last().EndBaseLine.X;
        List<List<Letter>> item = lines.Skip(1).ToList();
        Assert.True(item.Count >= 3, $"The item wraps to {item.Count} lines.");
        foreach (List<Letter> line in item.Take(item.Count - 1))
            Assert.InRange(line.Last().EndBaseLine.X, right - .05D, right + .05D);
        double lastEnd = item.Last().Last().EndBaseLine.X;
        if (alignment == WordParagraphAlignment.Distribute) Assert.InRange(lastEnd, right - .05D, right + .05D);
        else Assert.True(lastEnd < right - 20D, $"The last line ends at {lastEnd:F2}.");
    }

    // A justified bulleted item followed by a right-to-left item: both keep their alignment in one list.
    [Fact]
    public void SaveAsPdf_ListItemsKeepJustificationAndDirection() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("RIGHTREF").ParagraphAlignment = WordParagraphAlignment.Right;
        WordList list = document.AddList(WordListStyle.Bulleted);
        list.AddItem("JLIST " + Filler + " JLISTEND").ParagraphAlignment = WordParagraphAlignment.Both;
        WordParagraph rightToLeft = list.AddItem("RTLITEM");
        rightToLeft._paragraph.ParagraphProperties!.BiDi = new DocumentFormat.OpenXml.Wordprocessing.BiDi();

        List<List<Letter>> lines = ConvertLines(document);
        double right = lines[0].Last().EndBaseLine.X;
        Assert.InRange(lines[1].Last().EndBaseLine.X, right - .05D, right + .05D);
        Assert.InRange(lines.Last().Last().EndBaseLine.X, right - .05D, right + .05D);
    }

    private static List<List<Letter>> ConvertLines(WordDocument document) {
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false,
            PageSize = new PdfPageSize(OfficeIMO.OfficeLength.Points(400), OfficeIMO.OfficeLength.Points(700)),
            Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(30)),
            FontFamily = "Helvetica"
        }).ToBytes(new PdfSaveOptions { AllowLoss = true }));
        return Assert.Single(pdf.GetPages()).Letters.Where(letter => letter.Value != " ")
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 1)).OrderByDescending(line => line.Key)
            .Select(line => line.OrderBy(letter => letter.StartBaseLine.X).ToList()).ToList();
    }
}
