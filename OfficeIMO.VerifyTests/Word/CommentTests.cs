using System;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using VerifyXunit;
using Xunit;

namespace OfficeIMO.VerifyTests.Word;

public class CommentTests : VerifyTestBase {

    private static async Task DoTest(WordprocessingDocument document) {
        var result = await ToVerifyResult(document);
        await Verifier.Verify(result, GetSettings());
    }

    [Fact]
    public async Task PlayingWithComments() {
        using var document = WordDocument.Create();
        document.AddParagraph("Test Section");

        document.Paragraphs[0].AddComment("Przemysław", "PK", "This is my comment");


        document.AddParagraph("Test Section - another line");

        document.Paragraphs[1].AddComment("Przemysław", "PK", "More comments");

        document.Save();

        await DoTest(document._wordprocessingDocument);
    }
}
