using System;
using System.IO;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_SaveAsPdfRelative(string folderPath, bool openWord) {
            Console.WriteLine("[*] Saving document as PDF using relative path");
            string docPath = Path.Combine(folderPath, "SaveAsPdfRelative.docx");
            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("First paragraph");
                document.AddParagraph("Second paragraph").Bold = true;

                WordList list = document.AddList(WordListStyle.Bulleted);
                list.AddItem("Bullet 1");
                list.AddItem("Bullet 2");

                WordTable table = document.AddTable(2, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "A1";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "B1";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "A2";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "B2";

                string current = Directory.GetCurrentDirectory();
                Directory.SetCurrentDirectory(folderPath);
                document.SaveAsPdf("output.pdf");
                Directory.SetCurrentDirectory(current);
            }
        }
    }
}