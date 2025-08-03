using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class CompareDocuments {
        internal static void Example_BasicComparison(string folderPath, bool openWord) {
            Console.WriteLine("[*] Comparing documents");

            string sourcePath = Path.Combine(folderPath, "CompareSource.docx");
            string targetPath = Path.Combine(folderPath, "CompareTarget.docx");

            using (WordDocument doc = WordDocument.Create(sourcePath)) {
                doc.AddParagraph("Hello");
                doc.Save(false);
            }

            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Hello World");
                doc.Save(false);
            }

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            result.Save(openWord);
        }
    }
}

