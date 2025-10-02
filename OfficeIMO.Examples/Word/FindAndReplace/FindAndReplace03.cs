using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class FindAndReplace {
        internal static void Example_FindAndReplace03(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document - Find & Replace with new lines");
            string filePath = Path.Combine(folderPath, "Find and Replace with New Lines.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Before KEY After");
                paragraph.SetBold();

                document.FindAndReplace("KEY", "Line1\nLine2");

                Console.WriteLine($"Paragraph text: {paragraph.Text.Replace(Environment.NewLine, "\\n")}");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Console.WriteLine($"Reloaded paragraph text: {document.Paragraphs[0].Text.Replace(Environment.NewLine, "\\n")}");
                document.Save(openWord);
            }
        }
    }
}
