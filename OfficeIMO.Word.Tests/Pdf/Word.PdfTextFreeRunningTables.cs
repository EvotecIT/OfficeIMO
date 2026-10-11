using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_TextFreeRunningTableRetainsBordersAndDrawsImageOnce(bool footer, bool image) {
        using WordDocument document = WordDocument.Create();
        WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
        WordTable table = story.AddTable(1, 2, WordTableStyle.TableGrid);
        if (image) table.Rows[0].Cells[0].Paragraphs[0].AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"), 32, 32);
        document.AddParagraph("BODY");
        string before = string.Concat(story.ChildElements.Select(element => element.OuterXml));
        byte[] bytes = document.ToPdfBytes(EmptyRunningOptions());
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Single(pdf.GetPages());
        Assert.Contains("BODY", pdf.GetPage(1).Text);
        Assert.Equal(image ? 1 : 0, pdf.GetPage(1).GetImages().Count());
        Assert.Contains(" RG", PdfOperatorSearchText.From(bytes));
        Assert.Equal(before, string.Concat(story.ChildElements.Select(element => element.OuterXml)));
    }
}
