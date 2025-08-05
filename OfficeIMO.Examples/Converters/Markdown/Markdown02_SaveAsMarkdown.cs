using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word.Converters {
    internal static class Markdown02_SaveAsMarkdown {
        public static void Example(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating Word document and saving as Markdown");
            
            using var doc = WordDocument.Create();
            
            // Add content
            doc.AddParagraph("Main Title").Style = WordParagraphStyles.Heading1;
            doc.AddParagraph("This is a regular paragraph with some text.");
            
            doc.AddParagraph("Features").Style = WordParagraphStyles.Heading2;

            var paragraph = doc.AddParagraph("This has ");
            paragraph.AddText("bold text").Bold = true;
            paragraph.AddText(" and ");
            paragraph.AddText("italic text").Italic = true;
            paragraph.AddText(" plus a ");
            paragraph.AddHyperLink("link", new Uri("https://example.com"));
            paragraph.AddText(".");

            var list = doc.AddList(WordListStyle.Bulleted);
            list.AddItem("First item");
            list.AddItem("Second item");
            list.AddItem("Third item");
            
            var table = doc.AddTable(2, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Header 1";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Header 2";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Data 1";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "Data 2";

            string imagePath = Path.Combine(AppContext.BaseDirectory, "..", "Assets", "OfficeIMO.png");
            doc.AddParagraph().AddImage(imagePath);
            
            // Save as Markdown
            string outputPath = Path.Combine(folderPath, "SaveAsMarkdown.md");
            doc.SaveAsMarkdown(outputPath);
            
            Console.WriteLine($"✓ Created: {outputPath}");
            
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
            }
        }
    }
}