using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class PageBreaks {
        internal static void Example_PageBreaks1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with page breaks and removing them");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with some page breaks1.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var paragraph = document.AddParagraph("Test 1");
                paragraph.Text = "Test 2";

                document.AddPageBreak();

                document.AddPageBreak();

                var paragraph1 = document.AddParagraph("Test 1");
                paragraph1.Text = "Test 3";


                document.Save(openWord);
            }
        }
    }
}
