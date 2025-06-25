using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

internal static partial class Paragraphs {
        internal static void Example_InsertParagraphAt(string folderPath, bool openWord) {
            Console.WriteLine("[*] Inserting paragraph at a specific index");
            string filePath = Path.Combine(folderPath, "Example-InsertParagraphAt.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Paragraph 1");
                document.AddParagraph("Paragraph 3");

                // Insert new paragraph at index 1
                var inserted = document.InsertParagraphAt(1);
                inserted.Text = "Paragraph 2";

                document.Save(openWord);
            }
        }
    }
