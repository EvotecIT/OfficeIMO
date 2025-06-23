using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Macros {
        public static void Example_ListMacros(string templatesPath, string folderPath, bool openWord) {
            Console.WriteLine("[*] Listing macro modules");
            string macroDocPath = Path.Combine(folderPath, "DocumentWithMacro.docm");

            using (WordDocument document = WordDocument.Load(macroDocPath)) {
                foreach (var macro in document.Macros) {
                    Console.WriteLine($"Found macro: {macro.Name}");
                }
                document.Save(openWord);
            }
        }
    }
}
