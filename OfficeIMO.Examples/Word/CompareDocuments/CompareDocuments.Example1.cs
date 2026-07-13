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
                doc.Save();
            }

            using (WordDocument doc = WordDocument.Create(targetPath)) {
                doc.AddParagraph("Hello World");
                doc.Save();
            }

            using WordDocument result = WordDocumentComparer.Compare(sourcePath, targetPath);
            result.Save();
            if (openWord) result.OpenInApplication();

            WordComparisonResult structuredResult = WordDocumentComparer.CompareStructure(sourcePath, targetPath);
            foreach (WordComparisonFinding finding in structuredResult.Findings) {
                Console.WriteLine($"{finding.ChangeKind} {finding.Scope} at {finding.Location}: {finding.Message}");
            }
        }
    }
}

