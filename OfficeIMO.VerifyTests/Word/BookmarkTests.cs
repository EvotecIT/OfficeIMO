using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using VerifyXunit;
using Xunit;

namespace OfficeIMO.VerifyTests.Word;

public class BookmarkTests : VerifyTestBase {

    private static async Task DoTest(WordprocessingDocument document) {
        var result = await ToVerifyResult(document);
        await Verifier.Verify(result, GetSettings());
    }

    [Fact]
    public async Task BasicWordWithBookmarks() {
        using var document = WordDocument.Create();
        document.AddParagraph("Test 1").AddBookmark("Start");

        var paragraph = document.AddParagraph("This is text");
        foreach (var text in new[] { "text1", "text2", "text3" }) {
            paragraph = paragraph.AddText(text);
            paragraph.Bold = true;
            paragraph.Italic = true;
            paragraph.Underline = UnderlineValues.DashDotDotHeavy;
        }

        document.AddPageBreak();
        document.AddPageBreak();

        document.AddParagraph("Test 2").AddBookmark("Middle1");

        paragraph.AddText("OK baby");

        document.AddPageBreak();
        document.AddPageBreak();

        document.AddParagraph("Test 3").AddBookmark("Middle0");

        document.AddPageBreak();
        document.AddPageBreak();

        document.AddParagraph("Test 4").AddBookmark("EndOfDocument");

        document.Bookmarks[2].Remove();

        document.AddPageBreak();
        document.AddPageBreak();

        document.AddParagraph("Test 5");

        document.PageBreaks[7].Remove(includingParagraph: false);
        document.PageBreaks[6].Remove(true);

        document.Save();

        await DoTest(document._wordprocessingDocument);
    }
}
