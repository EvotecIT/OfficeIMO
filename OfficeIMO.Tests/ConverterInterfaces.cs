using OfficeIMO.Converters;
using OfficeIMO.Markdown;
using OfficeIMO.Html;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using System;
using System.IO;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

public class ConverterInterfaces {
    [Fact]
    public void MarkdownConverters_Work_With_Interface() {
        string markdown = "# Title\nSome text";
        using MemoryStream input = new MemoryStream(Encoding.UTF8.GetBytes(markdown));
        using MemoryStream wordStream = new MemoryStream();
        IWordConverter mdToWord = new MarkdownToWordConverter();
        mdToWord.Convert(input, wordStream, new MarkdownToWordOptions());
        Assert.True(wordStream.Length > 0);

        wordStream.Position = 0;
        using MemoryStream markdownStream = new MemoryStream();
        IWordConverter wordToMd = new WordToMarkdownConverter();
        wordToMd.Convert(wordStream, markdownStream, new WordToMarkdownOptions());
        string result = Encoding.UTF8.GetString(markdownStream.ToArray());
        Assert.Contains("Title", result);
    }

    [Fact]
    public void HtmlConverters_Work_With_Interface() {
        string html = "<p>Hello <b>world</b></p>";
        using MemoryStream input = new MemoryStream(Encoding.UTF8.GetBytes(html));
        using MemoryStream wordStream = new MemoryStream();
        IWordConverter htmlToWord = new HtmlToWordConverter();
        htmlToWord.Convert(input, wordStream, new HtmlToWordOptions());
        Assert.True(wordStream.Length > 0);

        wordStream.Position = 0;
        using MemoryStream htmlStream = new MemoryStream();
        IWordConverter wordToHtml = new WordToHtmlConverter();
        wordToHtml.Convert(wordStream, htmlStream, new WordToHtmlOptions());
        string roundTrip = Encoding.UTF8.GetString(htmlStream.ToArray());
        Assert.Contains("<b>world</b>", roundTrip, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void PdfConverter_Works_With_Interface() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Hello PDF");
        using MemoryStream docStream = new MemoryStream();
        document.Save(docStream);
        docStream.Position = 0;
        using MemoryStream pdfStream = new MemoryStream();
        IWordConverter converter = new WordPdfConverter();
        converter.Convert(docStream, pdfStream, new PdfSaveOptions());
        Assert.True(pdfStream.Length > 0);
    }
}
