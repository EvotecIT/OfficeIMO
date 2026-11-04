using OfficeIMO.Word;
using System.Threading.Tasks;
using VerifyXunit;
using Xunit;

namespace OfficeIMO.VerifyTests.Word;

/// <summary>
/// Validates cover page functionality within generated documents.
/// </summary>
public class CoverPageTests {
    [Fact]
    public async Task AddingCoverPage() {
        using var document = WordDocument.Create();
        document.Sections[0].PageSetup.Size.PageSize = WordPageSize.A4;
        document.PageSetup.Size.PageSize = WordPageSize.A4;
        document.Properties.Title = "Cover Page Templates";
        document.Properties.Subject = "How to use Cover Pages with TOC";
        document.ApplicationProperties.Company = "Evotec Services";
        document.Settings.UpdateFieldsOnOpen = true;

        document.AddCoverPage(WordCoverPageTemplate.IonDark);
        document.AddTableOfContent();
        document.AddPageBreak();

        var wordListToc = document.AddTableOfContentList(WordListStyle.Numbered);

        wordListToc.AddItem("Prepare document");
        document.AddParagraph("This is my test 1");
        wordListToc.AddItem("Make it shine");
        document.AddParagraph("This is my test 2");
        document.AddPageBreak();
        wordListToc.AddItem("More on the next page");

        await Verifier.Verify(document);
    }

    [Fact]
    public async Task AddingCoverPage2() {
        using var document = WordDocument.Create();
        document.Properties.Title = "Cover Page Templates";
        document.Properties.Subject = "How to use Cover Pages with TOC";

        document.Settings.UpdateFieldsOnOpen = true;

        document.AddCoverPage(WordCoverPageTemplate.Austin);

        document.AddTableOfContent();

        document.AddPageBreak();

        var wordListToc = document.AddTableOfContentList(WordListStyle.Numbered);

        wordListToc.AddItem("Prepare document");

        document.AddParagraph("This is my test 1");

        wordListToc.AddItem("Make it shine");

        document.AddParagraph("This is my test 2");

        document.AddPageBreak();

        wordListToc.AddItem("More on the next page");

        document.TableOfContent!.Update();

        await Verifier.Verify(document);
    }
}
