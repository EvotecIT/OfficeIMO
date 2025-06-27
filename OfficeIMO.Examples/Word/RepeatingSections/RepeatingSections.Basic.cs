using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class RepeatingSections {
        internal static void Example_BasicRepeatingSection(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a repeating section control");
            string filePath = Path.Combine(folderPath, "DocumentWithRepeatingSection.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph().AddRepeatingSection("Section", "RS", "RSTag");
                document.Save(openWord);
            }
        }
    }
}
