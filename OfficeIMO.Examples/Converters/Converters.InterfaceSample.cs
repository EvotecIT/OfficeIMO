using System.IO;
using System.Text;
using OfficeIMO.Word.Converters;
using OfficeIMO.Markdown;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Converters {
    public static class ConvertersInterfaceSample {
        public static void Example() {
            IWordConverter toWord = new MarkdownToWordConverter();
            using MemoryStream markdown = new MemoryStream(Encoding.UTF8.GetBytes("# Example\n"));
            using MemoryStream docx = new MemoryStream();
            toWord.Convert(markdown, docx, new MarkdownToWordOptions());

            docx.Position = 0;
            IWordConverter toPdf = new WordPdfStreamConverter();
            using MemoryStream pdf = new MemoryStream();
            toPdf.Convert(docx, pdf, new PdfSaveOptions());
        }
    }
}
