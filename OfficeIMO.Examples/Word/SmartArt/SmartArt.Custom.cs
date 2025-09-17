using System;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class SmartArt {
        internal static void Example_AddCustomSmartArt1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with Custom SmartArt 1");
            string filePath = System.IO.Path.Combine(folderPath, "SmartArtCustom1.docx");
            using WordDocument document = WordDocument.Create(filePath);
            document.AddSmartArt(SmartArtType.CustomSmartArt1);
            document.Save(openWord);
            OfficeIMO.Examples.Utils.Validation.ValidateDoc(filePath);
        }

        internal static void Example_AddCustomSmartArt2(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with Custom SmartArt 2");
            string filePath = System.IO.Path.Combine(folderPath, "SmartArtCustom2.docx");
            using WordDocument document = WordDocument.Create(filePath);
            document.AddSmartArt(SmartArtType.CustomSmartArt2);
            document.Save(openWord);
            OfficeIMO.Examples.Utils.Validation.ValidateDoc(filePath);
        }
    }
}

