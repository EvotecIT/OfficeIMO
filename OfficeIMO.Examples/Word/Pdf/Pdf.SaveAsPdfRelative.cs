using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Pdf {
        public static void Example_SaveAsPdfRelative(string folderPath, bool openWord) {
            Console.WriteLine("[*] Saving document as PDF using relative path");
            string docPath = Path.Combine(folderPath, "SaveAsPdfRelative.docx");
            using (WordDocument document = WordDocument.Create(docPath)) {
                document.AddParagraph("Hello PDF");
                string current = Directory.GetCurrentDirectory();
                Directory.SetCurrentDirectory(folderPath);
                document.SaveAsPdf("output.pdf");
                Directory.SetCurrentDirectory(current);
            }
        }
    }
}