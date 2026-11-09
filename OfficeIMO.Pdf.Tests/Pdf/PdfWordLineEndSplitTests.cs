using System.Linq;
using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public class PdfWordLineEndSplitTests {
    private const string Address = "security-alerts@example.org";

    [Fact]
    public void WordOverflowingTheLineEndSplitsAfterItsHyphenByDefault() {
        string[] words = Words(new PdfCreateOptions());
        Assert.Contains("security-", words);
        Assert.Contains("alerts@example.org", words);
    }

    [Fact]
    public void WordThatFitsAFreshLineMovesThereWholeWhenSplittingIsOff() {
        string[] words = Words(new PdfCreateOptions { SplitWordsAtLineEnd = false });
        Assert.Contains(Address, words);
        Assert.True(new PdfCreateOptions { SplitWordsAtLineEnd = false }.Clone().SplitWordsAtLineEnd == false);
    }

    [Fact]
    public void WordLongerThanALineStillSplitsWhenSplittingIsOff() {
        string longWord = string.Join("-", Enumerable.Repeat("segment", 8));
        using PdfDocument document = CreateDocument(new PdfCreateOptions { SplitWordsAtLineEnd = false }, longWord);
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.True(pdf.GetPage(1).GetWords().Count() > 1);
    }

    // 170pt of text at 12pt Helvetica: "Contact security-" fits the first line, the whole phrase does not,
    // and the address alone fits a fresh line.
    private static string[] Words(PdfCreateOptions options) {
        using PdfDocument document = CreateDocument(options, "Contact " + Address);
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        return pdf.GetPage(1).GetWords().Select(word => word.Text).ToArray();
    }

    private static PdfDocument CreateDocument(PdfCreateOptions options, string text) {
        options.EmbedDefaultFonts = false;
        options.PageSetup.Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(250), global::OfficeIMO.OfficeLength.Points(400));
        options.PageSetup.Margins.Left = options.PageSetup.Margins.Right = global::OfficeIMO.OfficeLength.Points(40);
        PdfDocument document = PdfDocument.Create(options);
        document.AddParagraph(paragraph => {
            paragraph.RunDefaults.FontSize = global::OfficeIMO.OfficeLength.Points(12);
            paragraph.AddText(text);
        });
        return document;
    }
}
