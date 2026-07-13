using OfficeIMO.Excel;
using OfficeIMO.Excel.Html;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public class HtmlOfficeAdaptersStreamIO {
    [Fact]
    public async Task ExcelHtml_StreamAndAsyncApisUseUtf8WithoutBomAndLeaveStreamsOpen() {
        using ExcelDocument workbook = ExcelDocument.Create(new MemoryStream());
        workbook.AddWorksheet("Data").CellValue(1, 1, "Zażółć");
        using var htmlStream = new MemoryStream();

        await workbook.SaveAsHtmlAsync(htmlStream);

        Assert.True(htmlStream.CanWrite);
        byte[] htmlBytes = htmlStream.ToArray();
        Assert.False(HasUtf8Bom(htmlBytes));
        htmlStream.Position = 0;
        HtmlToExcelResult result = await htmlStream.ToExcelDocumentResultAsync();
        using ExcelDocument imported = result.Value;
        Assert.True(result.Succeeded);
        Assert.True(htmlStream.CanRead);
    }

    [Fact]
    public async Task PowerPointHtml_StreamAndAsyncApisUseUtf8WithoutBomAndLeaveStreamsOpen() {
        using PowerPointPresentation presentation = PowerPointPresentation.Create(new MemoryStream());
        presentation.AddSlide().AddTextBox("Zażółć");
        using var htmlStream = new MemoryStream();

        await presentation.SaveAsHtmlAsync(htmlStream);

        Assert.True(htmlStream.CanWrite);
        byte[] htmlBytes = htmlStream.ToArray();
        Assert.False(HasUtf8Bom(htmlBytes));
        htmlStream.Position = 0;
        HtmlToPowerPointResult result = await htmlStream.ToPowerPointPresentationResultAsync();
        using PowerPointPresentation imported = result.Value;
        Assert.True(result.Succeeded);
        Assert.True(htmlStream.CanRead);
    }

    [Fact]
    public async Task SharedHtmlTextIOIsUsedByWordMarkdownAndPdfStreams() {
        using WordDocument word = "<p>Zażółć</p>".ToWordDocument();
        using var wordHtml = new MemoryStream();
        await word.SaveAsHtmlAsync(wordHtml);
        Assert.False(HasUtf8Bom(wordHtml.ToArray()));
        Assert.True(wordHtml.CanWrite);

        wordHtml.Position = 0;
        HtmlToWordResult wordResult = await wordHtml.ToWordDocumentResultAsync();
        using WordDocument importedWord = wordResult.Value;
        Assert.True(wordHtml.CanRead);
        Assert.True(wordResult.Succeeded);

        byte[] bomHtml = Encoding.UTF8.GetPreamble()
            .Concat(Encoding.UTF8.GetBytes("<p>Stream marker</p>"))
            .ToArray();
        using var markdownStream = new MemoryStream(bomHtml);
        Assert.Contains("Stream marker", markdownStream.ToMarkdown(), StringComparison.Ordinal);
        Assert.True(markdownStream.CanRead);

        using var pdfStream = new MemoryStream(bomHtml);
        var pdfResult = await pdfStream.ToPdfDocumentResultAsync();
        Assert.True(pdfStream.CanRead);
        Assert.NotEmpty(pdfResult.ToBytes());
    }

    private static bool HasUtf8Bom(byte[] bytes) =>
        bytes.Length >= 3 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF;
}
