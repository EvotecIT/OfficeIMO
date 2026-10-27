using System;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfFinalTableOfContentsPagesTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void NumberedHeading_InContentsHasAnEmittedDestination(bool columns, bool paragraphPath) {
        using WordDocument document = WordDocument.Create();
        document.AddTableOfContent(WordTableOfContentsStyle.Template1, 1, 1);
        document.AddPageBreak();
        if (columns) document.Sections[0].ColumnCount = 2;
        var list = document.AddList(WordListStyle.Numbered);
        WordParagraph heading = list.AddItem("Numbered findings");
        heading.Style = WordParagraphStyles.Heading1;
        if (paragraphPath) heading.Spacing = 20;
        document.AddParagraph("Numbered body marker");
        byte[] bytes = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes();
        var loaded = PdfDocument.Load(bytes);
        Assert.Contains("Numbered findings", loaded.Reader.TextByPage()[0]);
        Assert.NotEmpty(loaded.Read().GetLinksByDestinationName("officeimo-heading-numbered-findings"));
        Assert.Equal(2, Assert.Single(loaded.Reader.NamedDestinations(), destination => destination.Name == "officeimo-heading-numbered-findings").PageNumber);
    }

    [Fact]
    public void SaveAsPdf_OverflowBeforeHeading_ContentsMatchesActualHeadingPage() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Cover");
        document.AddPageBreak();
        document.AddTableOfContent(WordTableOfContentsStyle.Template1, 1, 1);
        document.AddPageBreak();
        document.AddParagraph("Executive summary").Style = WordParagraphStyles.Heading1;
        for (int index = 0; index < 80; index++) {
            WordParagraph paragraph = document.AddParagraph("Evidence row " + index + ": " + new string('x', 120));
            paragraph.FontSize = 9;
        }
        document.AddParagraph("Final findings").Style = WordParagraphStyles.Heading1;
        document.AddParagraph("Final body marker");
        byte[] bytes = document.ToPdf(new WordPdfExportOptions { IncludePageNumbers = false }).ToBytes();
        var pages = PdfDocument.Load(bytes).Reader.TextByPage();
        string contents = pages[1];
        int findingsPage = Enumerable.Range(2, pages.Count - 2).Single(index => pages[index].Contains("Final body marker")) + 1;
        Assert.True(findingsPage > 3);
        Assert.Matches(@"Executive summary[\s.]*3\b", contents);
        Assert.Matches(@"Final findings[\s.]*" + findingsPage + @"\b", contents);
        Assert.NotEmpty(PdfDocument.Load(bytes).Read().GetLinksByDestinationName("officeimo-heading-final-findings"));
    }
}
