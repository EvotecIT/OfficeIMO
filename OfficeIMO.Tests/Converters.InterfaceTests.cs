using System.IO;
using System.Text;
using OfficeIMO.Word.Converters;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public class ConvertersInterfaceTests {
        [Fact]
        public void MarkdownConvertersImplementInterface() {
            IWordConverter toWord = new MarkdownToWordConverter();
            IWordConverter toMarkdown = new WordToMarkdownConverter();

            using MemoryStream markdownStream = new MemoryStream(Encoding.UTF8.GetBytes("# Test\n"));
            using MemoryStream docxStream = new MemoryStream();
            toWord.Convert(markdownStream, docxStream, new MarkdownToWordOptions());
            Assert.True(docxStream.Length > 0);

            docxStream.Position = 0;
            using MemoryStream mdOutput = new MemoryStream();
            toMarkdown.Convert(docxStream, mdOutput, new WordToMarkdownOptions());
            mdOutput.Position = 0;
            using StreamReader reader = new StreamReader(mdOutput);
            string result = reader.ReadToEnd();
            Assert.Contains("Test", result);
        }

        [Fact]
        public void HtmlConvertersImplementInterface() {
            IWordConverter toWord = new HtmlToWordConverter();
            IWordConverter toHtml = new WordToHtmlConverter();

            using MemoryStream htmlStream = new MemoryStream(Encoding.UTF8.GetBytes("<p>Hello</p>"));
            using MemoryStream docxStream = new MemoryStream();
            toWord.Convert(htmlStream, docxStream, new HtmlToWordOptions());
            Assert.True(docxStream.Length > 0);

            docxStream.Position = 0;
            using MemoryStream htmlOut = new MemoryStream();
            toHtml.Convert(docxStream, htmlOut, new WordToHtmlOptions());
            htmlOut.Position = 0;
            using StreamReader reader = new StreamReader(htmlOut);
            string result = reader.ReadToEnd();
            Assert.Contains("Hello", result);
        }

        [Fact]
        public void PdfConverterImplementsInterface() {
            IWordConverter toPdf = new WordPdfStreamConverter();

            using MemoryStream docxStream = new MemoryStream();
            using (WordDocument doc = WordDocument.Create(docxStream, autoSave: true)) {
                doc.AddParagraph("Hi");
            }
            docxStream.Position = 0;
            using MemoryStream pdfOut = new MemoryStream();
            toPdf.Convert(docxStream, pdfOut, new PdfSaveOptions());
            Assert.True(pdfOut.Length > 0);
        }
    }
}
