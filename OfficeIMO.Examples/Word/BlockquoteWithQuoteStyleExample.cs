using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word {
    internal static class BlockquoteWithQuoteStyleExample {
        public static void Example_BlockquoteWithQuoteStyle(string folderPath, bool openWord) {
            var style = WordParagraphStyle.CreateFontStyle("Quote", "Arial");
            WordParagraphStyle.RegisterCustomStyle("Quote", style);

            string html = "<blockquote>Quoted text</blockquote>";
            using WordDocument document = html.LoadFromHtml(new HtmlToWordOptions());
            string docPath = Path.Combine(folderPath, "BlockquoteWithQuoteStyle.docx");
            document.Save(docPath);
            Console.WriteLine($"✓ Created: {docPath}");
        }
    }
}

