using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class SmartArt {
        internal static void Example_AddBasicSmartArt(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with a SmartArt diagram");
            string filePath = System.IO.Path.Combine(folderPath, "SmartArtBasic.docx");
            using WordDocument document = WordDocument.Create(filePath);
            document.AddSmartArt(SmartArtType.BasicProcess);
            document.Save(openWord);
        }
    }
}
