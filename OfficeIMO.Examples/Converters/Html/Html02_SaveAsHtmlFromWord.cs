using System;
using System.IO;
using OfficeIMO.Word;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Examples.Html {
    internal static partial class Html {
        public static void Example_Html02_SaveAsHtmlFromWord(string folderPath, bool openWord) {
            Console.WriteLine("[*] From Word → HTML with options");

            using var doc = WordDocument.Create();
            doc.AddParagraph("Main Title").Style = WordParagraphStyles.Heading1;
            doc.AddParagraph("This is a regular paragraph with some text.");

            doc.AddParagraph("Features").Style = WordParagraphStyles.Heading2;
            var p = doc.AddParagraph("This has ");
            p.AddText("bold").Bold = true;
            p.AddText(" and ");
            p.AddText("italic").Italic = true;

            var list = doc.AddList(WordListStyle.Bulleted);
            list.AddItem("First item");
            list.AddItem("Second item");
            list.AddItem("Third item");

            var table = doc.AddTable(2, 2);
            table.Rows[0].Cells[0].Paragraphs[0].Text = "Header 1";
            table.Rows[0].Cells[1].Paragraphs[0].Text = "Header 2";
            table.Rows[1].Cells[0].Paragraphs[0].Text = "Data 1";
            table.Rows[1].Cells[1].Paragraphs[0].Text = "Data 2";

            string htmlPath = Path.Combine(folderPath, "Html02_FromWord.html");
            doc.SaveAsHtml(htmlPath, new WordToHtmlOptions {
                IncludeFontStyles = true,
                IncludeListStyles = true
            });

            string docxPath = Path.Combine(folderPath, "Html02_FromWord.docx");
            doc.Save(docxPath);

            Console.WriteLine($"✓ Created: {htmlPath}");
            Console.WriteLine($"✓ Created: {docxPath}");
            if (openWord) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(htmlPath) { UseShellExecute = true });
            }
        }
    }
}

