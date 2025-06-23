using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class Macros {
        public static void Example_ListAndRemoveMacro(string templatesPath, string folderPath, bool openWord) {
            Console.WriteLine("[*] Listing and removing macro modules");
            string macroDocPath = Path.Combine(folderPath, "DocumentWithMacro.docm");

            using (WordDocument document = WordDocument.Load(macroDocPath)) {
                foreach (var name in document.Macros) {
                    Console.WriteLine($"Found macro: {name}");
                    document.RemoveMacro(name);
                }
                document.Save(openWord);
            }
        }
    }
}
