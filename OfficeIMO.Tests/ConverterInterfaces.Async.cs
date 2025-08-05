using OfficeIMO.Converters;
using OfficeIMO.Markdown;
using OfficeIMO.Html;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public class ConverterInterfacesAsync {
    [Fact]
    public async Task MarkdownConverters_Work_With_Interface_Async() {
        string markdown = "# Title\nSome text";
        using MemoryStream input = new MemoryStream(Encoding.UTF8.GetBytes(markdown));
        using MemoryStream wordStream = new MemoryStream();
        IWordConverter mdToWord = new MarkdownToWordConverter();
        await mdToWord.ConvertAsync(input, wordStream, new MarkdownToWordOptions());
        Assert.True(wordStream.Length > 0);

        wordStream.Position = 0;
        using MemoryStream markdownStream = new MemoryStream();
        IWordConverter wordToMd = new WordToMarkdownConverter();
        await wordToMd.ConvertAsync(wordStream, markdownStream, new WordToMarkdownOptions());
        string result = Encoding.UTF8.GetString(markdownStream.ToArray());
        Assert.Contains("Title", result);
    }

    [Fact]
    public async Task HtmlConverters_Work_With_Interface_Async() {
        string html = "<p>Hello <b>world</b></p>";
        using MemoryStream input = new MemoryStream(Encoding.UTF8.GetBytes(html));
        using MemoryStream wordStream = new MemoryStream();
        IWordConverter htmlToWord = new HtmlToWordConverter();
        await htmlToWord.ConvertAsync(input, wordStream, new HtmlToWordOptions());
        Assert.True(wordStream.Length > 0);

        wordStream.Position = 0;
        using MemoryStream htmlStream = new MemoryStream();
        IWordConverter wordToHtml = new WordToHtmlConverter();
        await wordToHtml.ConvertAsync(wordStream, htmlStream, new WordToHtmlOptions());
        string roundTrip = Encoding.UTF8.GetString(htmlStream.ToArray());
        Assert.Contains("<b>world</b>", roundTrip, System.StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public async Task PdfConverter_Works_With_Interface_Async() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Hello PDF");
        using MemoryStream docStream = new MemoryStream();
        document.Save(docStream);
        docStream.Position = 0;
        using MemoryStream pdfStream = new MemoryStream();
        IWordConverter converter = new WordPdfConverter();
        await converter.ConvertAsync(docStream, pdfStream, new PdfSaveOptions());
        Assert.True(pdfStream.Length > 0);
    }
}
