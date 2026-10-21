using System;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPdfFinalTableOfContentsPagesTests {
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
        byte[] bytes = document.ToPdfDocument(new WordToPdfOptions { IncludePageNumbers = false }).ToBytes();
        var pages = PdfDocument.Load(bytes).Reader.TextByPage();
        string contents = pages[1];
        int findingsPage = Enumerable.Range(2, pages.Count - 2).Single(index => pages[index].Contains("Final body marker")) + 1;
        Assert.True(findingsPage > 3);
        Assert.Matches(@"Executive summary[\s.]*3\b", contents);
        Assert.Matches(@"Final findings[\s.]*" + findingsPage + @"\b", contents);
        Assert.NotEmpty(PdfDocument.Load(bytes).Read().GetLinksByDestinationName("officeimo-heading-final-findings"));
    }
}
