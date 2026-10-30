using OfficeIMO.Word;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordPageNumberStyle.Brackets2)]
    [InlineData(WordPageNumberStyle.Dots)]
    [InlineData(WordPageNumberStyle.RoundedRectangle)]
    [InlineData(WordPageNumberStyle.Tab)]
    [InlineData(WordPageNumberStyle.VerticalOutline1)]
    [InlineData(WordPageNumberStyle.VerticalOutline2)]
    public void SaveAsPdf_PageNumberInsideGroupedTextBoxIsRendered(WordPageNumberStyle style) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("First page");
        document.AddPageBreak();
        document.AddParagraph("Second page");
        document.AddHeadersAndFooters();
        document.Footer!.Default.AddPageNumber(style);

        using PdfPigDocument pdf = OpenAuditPdf(document);
        Assert.Equal(2, pdf.NumberOfPages);
        foreach (int pageNumber in new[] { 1, 2 }) {
            var page = pdf.GetPage(pageNumber);
            Assert.Contains(page.GetWords(), word => word.Text == pageNumber.ToString(System.Globalization.CultureInfo.InvariantCulture));
        }
    }
}
