using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Macros {
        public static void Example_AddMacroToExistingDocx(string templatesPath, string folderPath, bool openWord) {
            Console.WriteLine("[*] Adding macro to existing document");
            string docPath = Path.Combine(templatesPath, "BasicDocument.docx");
            string macroPath = Path.Combine(templatesPath, "vbaProject.bin");
            string filePath = Path.Combine(folderPath, "DocumentWithMacroFromDocx.docm");

            File.Copy(docPath, filePath, true);
            using (WordDocument document = WordDocument.Load(filePath)) {
                document.AddMacro(macroPath);
                document.Save(openWord);
            }
        }
    }
}
