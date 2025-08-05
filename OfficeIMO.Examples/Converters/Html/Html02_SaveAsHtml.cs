using OfficeIMO.Word;
using OfficeIMO.Word.Html;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word.Converters {
    internal static class Html02_SaveAsHtml {
        public static void Example(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating Word document and saving as HTML");
            
            using var doc = WordDocument.Create();
            
            // Add content
            doc.AddParagraph("Main Title").Style = WordParagraphStyles.Heading1;
            doc.AddParagraph("This is a regular paragraph with some text.");
            
            doc.AddParagraph("Features").Style = WordParagraphStyles.Heading2;
            
            var paragraph = doc.AddParagraph("This has ");
            paragraph.AddText("bold text").Bold = true;
            paragraph.AddText(" and ");
            paragraph.AddText("italic text").Italic = true;
            
            var list = doc.AddList(WordListStyle.Bulleted);
            list.AddItem("First item");
            list.AddItem("Second item");
            list.AddItem("Third item");
            
            doc.AddParagraph("Table Example").Style = WordParagraphStyles.Heading3;
            
            var table = doc.AddTable(2, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Header 1";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Header 2";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Data 1";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "Data 2";
            
            var linkPara = doc.AddParagraph("Visit ");
            linkPara.AddHyperLink("GitHub", new Uri("https://github.com"));
            linkPara.AddText(" for more info.");
            
            // Save as HTML
            string outputPath = Path.Combine(folderPath, "SaveAsHtml.html");
            doc.SaveAsHtml(outputPath, new WordToHtmlOptions {
                IncludeFontStyles = true,
                IncludeListStyles = true
            });
            
            Console.WriteLine($"✓ Created: {outputPath}");
            
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(outputPath) { UseShellExecute = true });
            }
        }
    }
}