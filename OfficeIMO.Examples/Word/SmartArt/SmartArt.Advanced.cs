using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class SmartArt {
        internal static void Example_AddAdvancedSmartArt(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with multiple SmartArt diagrams");
            string filePath = System.IO.Path.Combine(folderPath, "SmartArtAdvanced.docx");
            using WordDocument document = WordDocument.Create(filePath);
            document.AddSmartArt(SmartArtType.Hierarchy);
            document.AddParagraph("Between diagrams");
            document.AddSmartArt(SmartArtType.Cycle);
            document.AddSmartArt(SmartArtType.PictureOrgChart);
            document.Save(openWord);
        }
    }
}
