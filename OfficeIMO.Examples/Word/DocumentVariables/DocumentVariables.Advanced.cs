using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class DocumentVariablesExamples {
        public static void Example_AdvancedDocumentVariables(string folderPath, bool openWord) {
            Console.WriteLine("[*] Working with document variables");
            string filePath = Path.Combine(folderPath, "AdvancedDocumentWithVariables.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Title = "Variables";
                document.SetDocumentVariable("Project", "OfficeIMO");
                document.SetDocumentVariable("Version", "1.0");
                document.SetDocumentVariable("Date", DateTime.Today.ToShortDateString());
                document.Save();
            }
            using (WordDocument document = WordDocument.Load(filePath, false)) {
                document.SetDocumentVariable("Version", "2.0");
                if (document.HasDocumentVariables) {
                    document.RemoveDocumentVariableAt(0);
                }
                document.Save(openWord);
            }
        }
    }
}
