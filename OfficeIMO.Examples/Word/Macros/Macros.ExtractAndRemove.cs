using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Macros {
        public static void Example_ExtractAndRemoveMacro(string templatesPath, string folderPath, bool openWord) {
            Console.WriteLine("[*] Extracting and removing macro");
            string macroDocPath = Path.Combine(folderPath, "DocumentWithMacro.docm");
            string macroPath = Path.Combine(folderPath, "ExtractedMacro.bin");

            using (WordDocument document = WordDocument.Load(macroDocPath)) {
                document.SaveMacros(macroPath);
                document.RemoveMacros();
                document.Save(openWord);
            }
        }
    }
}
