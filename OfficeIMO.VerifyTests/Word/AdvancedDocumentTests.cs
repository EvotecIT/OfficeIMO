using System;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using VerifyXunit;
using Xunit;

using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.VerifyTests.Word;

public class AdvancedDocumentTests : VerifyTestBase {

    private static async Task DoTest(WordprocessingDocument document) {
        var result = await ToVerifyResult(document);
        await Verifier.Verify(result, GetSettings());
    }

    [Fact]
    public async Task AdvancedWordCreate() {
        using var document = WordDocument.Create();
        // lets add some properties to the document
        document.BuiltinDocumentProperties.Title = "Cover Page Templates";
        document.BuiltinDocumentProperties.Subject = "How to use Cover Pages with TOC";
        document.ApplicationProperties.Company = "Evotec Services";

        // we force document to update fields on open, this will be used by TOC
        document.Settings.UpdateFieldsOnOpen = true;

        // lets add one of multiple added Cover Pages
        document.AddCoverPage(CoverPageTemplate.IonDark);

        // lets add Table of Content (1 of 2)
        document.AddTableOfContent(TableOfContentStyle.Template1);

        // lets add page break
        document.AddPageBreak();

        // lets create a list that will be binded to TOC
        var wordListToc = document.AddTableOfContentList(WordListStyle.Headings111);

        wordListToc.AddItem("How to add a table to document?");

        document.AddParagraph(
            "In the first paragraph I would like to show you how to add a table to the document using one of the 105 built-in styles:");

        // adding a table and modifying content
        var table = document.AddTable(5, 4, WordTableStyle.GridTable5DarkAccent5);
        table.Rows[3].Cells[2].Paragraphs[0].Text = "Adding text to cell";
        table.Rows[3].Cells[2].Paragraphs[0].Color = Color.Blue;
        ;
        table.Rows[3].Cells[3].Paragraphs[0].Text = "Different cell";

        document.AddParagraph("As you can see adding a table with some style, and adding content to it ").SetBold()
            .SetUnderline(UnderlineValues.Dotted).AddText("is not really complicated").SetColor(Color.OrangeRed);

        wordListToc.AddItem("How to add a list to document?");

        var paragraph = document
            .AddParagraph("Adding lists is similar to adding a table. Just define a list and add list items to it. ")
            .SetText("Remember that you can add anything between list items! ");
        paragraph.SetColor(Color.Blue)
            .SetText("For example TOC List is just another list, but defining a specific style.");

        var list = document.AddList(WordListStyle.Bulleted);
        list.AddItem("First element of list", 0);
        list.AddItem("Second element of list", 1);

        var paragraphWithHyperlink = document.AddHyperLink("Go to Evotec Blogs", new Uri("https://evotec.xyz"),
            true, "URL with tooltip");
        // you can also change the hyperlink text, uri later on using properties
        paragraphWithHyperlink.Hyperlink.Uri = new Uri("https://evotec.xyz/hub");
        paragraphWithHyperlink.ParagraphAlignment = JustificationValues.Center;

        list.AddItem("3rd element of list, but added after hyperlink", 0);
        list.AddItem("4th element with hyperlink ")
            .AddHyperLink("included.", new Uri("https://evotec.xyz/hub"), addStyle: true);

        document.AddParagraph();

        var listNumbered = document.AddList(WordListStyle.Heading1ai);
        listNumbered.AddItem("Different list number 1");
        listNumbered.AddItem("Different list number 2", 1);
        listNumbered.AddItem("Different list number 3", 1);
        listNumbered.AddItem("Different list number 4", 1);

        var section = document.AddSection();
        section.PageOrientation = PageOrientationValues.Landscape;
        section.PageSettings.PageSize = WordPageSize.A4;

        wordListToc.AddItem("Adding headers / footers");

        // lets add headers and footers
        document.AddHeadersAndFooters();

        // adding text to default header
        document.Header.Default.AddParagraph("Text added to header - Default");

        var section1 = document.AddSection();
        section1.PageOrientation = PageOrientationValues.Portrait;
        section1.PageSettings.PageSize = WordPageSize.A5;

        wordListToc.AddItem("Adding custom properties and page numbers to document");

        document.CustomDocumentProperties.Add("TestProperty", new WordCustomProperty {Value = DateTime.Today});
        document.CustomDocumentProperties.Add("MyName", new WordCustomProperty("Some text"));
        document.CustomDocumentProperties.Add("IsTodayGreatDay", new WordCustomProperty(true));

        // add page numbers
        document.Footer.Default.AddPageNumber(WordPageNumberStyle.PlainNumber);

        // add watermark
        document.Sections[0].AddWatermark(WordWatermarkStyle.Text, "Draft");

        document.Save();

        await DoTest(document._wordprocessingDocument);
    }
}
