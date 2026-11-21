using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Examples.Word {
    internal static class BlockquoteWithQuoteStyleExample {
        public static void Example_BlockquoteWithQuoteStyle(string folderPath, bool openWord) {
            using WordDocument document = WordDocument.Create();
            // Blockquotes use the target document's Quote style when it defines one.
            document.Styles.AddParagraphStyle("Quote").FontFamily = "Arial";
            document.AddHtml("<blockquote>Quoted text</blockquote>", new HtmlWordExportOptions());
            string docPath = Path.Combine(folderPath, "BlockquoteWithQuoteStyle.docx");
            document.Save(docPath);
            Console.WriteLine($"✓ Created: {docPath}");
        }
    }
}

