using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class DocumentVariablesExamples {
        public static void Example_BasicDocumentVariables(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with variables");
            string filePath = Path.Combine(folderPath, "DocumentWithVariables.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.SetDocumentVariable("Author", "OfficeIMO");
                document.SetDocumentVariable("Year", DateTime.Now.Year.ToString());
                document.Save(openWord);
            }
            using (WordDocument document = WordDocument.Load(filePath, false)) {
                Console.WriteLine($"Author: {document.GetDocumentVariable("Author")}");
                Console.WriteLine($"Year: {document.GetDocumentVariable("Year")}");
                if (document.HasDocumentVariables) {
                    foreach (var pair in document.DocumentVariables) {
                        Console.WriteLine($"{pair.Key} -> {pair.Value}");
                    }
                }
            }
        }
    }
}
