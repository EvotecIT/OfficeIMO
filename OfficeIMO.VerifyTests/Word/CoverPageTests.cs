using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using VerifyXunit;
using Xunit;

namespace OfficeIMO.VerifyTests.Word;

public class CoverPageTests : VerifyTestBase {

    private static async Task DoTest(WordprocessingDocument document) {
        var result = await ToVerifyResult(document);
        await Verifier.Verify(result, GetSettings());
    }

    [Fact]
    public async Task AddingCoverPage() {
        using var document = WordDocument.Create();
        document.Sections[0].PageSettings.PageSize = WordPageSize.A4;
        document.PageSettings.PageSize = WordPageSize.A4;
        document.BuiltinDocumentProperties.Title = "Cover Page Templates";
        document.BuiltinDocumentProperties.Subject = "How to use Cover Pages with TOC";
        document.ApplicationProperties.Company = "Evotec Services";
        document.Settings.UpdateFieldsOnOpen = true;

        document.AddCoverPage(CoverPageTemplate.IonDark);
        document.AddTableOfContent();
        document.AddPageBreak();

        var wordListToc = document.AddTableOfContentList(WordListStyle.Headings111);

        wordListToc.AddItem("Prepare document");
        document.AddParagraph("This is my test 1");
        wordListToc.AddItem("Make it shine");
        document.AddParagraph("This is my test 2");
        document.AddPageBreak();
        wordListToc.AddItem("More on the next page");

        document.Save();

        await DoTest(document._wordprocessingDocument);
    }

    [Fact]
    public async Task AddingCoverPage2() {
        using var document = WordDocument.Create();
        document.BuiltinDocumentProperties.Title = "Cover Page Templates";
        document.BuiltinDocumentProperties.Subject = "How to use Cover Pages with TOC";

        document.Settings.UpdateFieldsOnOpen = true;

        document.AddCoverPage(CoverPageTemplate.Austin);

        document.AddTableOfContent();

        document.AddPageBreak();

        var wordListToc = document.AddTableOfContentList(WordListStyle.Headings111);

        wordListToc.AddItem("Prepare document");

        document.AddParagraph("This is my test 1");

        wordListToc.AddItem("Make it shine");

        document.AddParagraph("This is my test 2");

        document.AddPageBreak();

        wordListToc.AddItem("More on the next page");

        document.TableOfContent.Update();

        document.Save();

        await DoTest(document._wordprocessingDocument);
    }
}
