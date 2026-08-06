using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class SmartArt {
        internal static void Example_AddAdvancedSmartArt(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with multiple SmartArt diagrams");
            string filePath = System.IO.Path.Combine(folderPath, "SmartArtAdvanced.docx");
            using WordDocument document = WordDocument.Create(filePath);
            document.AddSmartArt(WordSmartArtType.Hierarchy);
            document.AddParagraph("Between diagrams");
            document.AddSmartArt(WordSmartArtType.Cycle);
            document.AddSmartArt(WordSmartArtType.PictureOrgChart);
            document.Save();
            if (openWord) document.OpenInApplication();
            OfficeIMO.Examples.Utils.Validation.ValidateDoc(filePath);
        }
    }
}
