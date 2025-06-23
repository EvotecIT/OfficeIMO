using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Macros {
        public static void Example_CreateDocmWithMacro(string templatesPath, string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating macro-enabled document");
            string filePath = Path.Combine(folderPath, "DocumentWithMacro.docm");
            string macroPath = Path.Combine(templatesPath, "vbaProject.bin");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Document with macro");
                document.AddMacro(macroPath);
                document.Save(openWord);
            }
        }
    }
}
