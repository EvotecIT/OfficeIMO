using OfficeIMO.Pdf;
using System.Globalization;
using System.Text.RegularExpressions;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTaggedParagraphReadingOrderTests {
    [Theory]
    [InlineData("WrapObjectOrderLayout", 4, "sigma tau.")]
    [InlineData("WrapObjectOrderLargerColumns", 5, "ma tau.")]
    public void TaggedStoryRetainsParagraphOrderAcrossColumns(string fixtureName, int continuedParagraph, string tail) {
        using PdfDocument document = LoadPublisherFixture(fixtureName);
        PdfDocumentReadResult result = document.Read();
        PdfLogicalPage page = Assert.Single(result.Pages);

        Assert.Empty(page.Tables);
        Assert.Equal(8, page.Paragraphs.Count);
        Assert.Equal(Enumerable.Range(1, 8), Regex.Matches(result.Text, @"\bParagraph\s+([1-8])\.")
            .Cast<Match>().Select(match => int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture)));
        Assert.Equal("Paragraph 1. Alpha beta", page.Paragraphs[0].Lines[0].Text);
        PdfLogicalParagraph continued = page.Paragraphs[continuedParagraph - 1];
        Assert.Equal(tail, continued.Lines[continued.Lines.Count - 1].Text);
    }

    [Theory]
    [InlineData(true, PdfReadingDirection.Auto)]
    [InlineData(false, PdfReadingDirection.RightToLeft)]
    public void TaggedStoryRespectsExplicitReadingOrder(bool forceSingleColumn, PdfReadingDirection direction) {
        using PdfDocument document = LoadPublisherFixture("WrapObjectOrderLayout");
        PdfDocumentReadResult result = document.Read(new PdfLoadOptions {
            LayoutOptions = new PdfTextLayoutOptions { ForceSingleColumn = forceSingleColumn, ReadingDirection = direction }
        });

        int higherParagraph = result.Text.IndexOf("Paragraph 5.", StringComparison.Ordinal);
        int lowerParagraph = result.Text.IndexOf("Paragraph 4.", StringComparison.Ordinal);
        Assert.InRange(higherParagraph, 0, lowerParagraph - 1);
    }

    private static PdfDocument LoadPublisherFixture(string fixtureName) =>
        PdfDocument.Load(Path.Combine(RepositoryTestPaths.Find(), "OfficeIMO.Publisher.Tests", "Fixtures",
            "NativeComparison", fixtureName, "native-" + fixtureName.ToLowerInvariant() + ".pdf"));
}
