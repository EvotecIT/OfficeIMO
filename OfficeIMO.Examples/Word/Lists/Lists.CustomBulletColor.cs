using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word {
    internal static partial class Lists {
        internal static void Example_CustomBulletColor(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating document with custom bullet color");
            string filePath = System.IO.Path.Combine(folderPath, "Document custom bullet color.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                var bulletList = document.AddCustomBulletList(WordListLevelKind.BulletDiamondSymbol, "Wingdings", Color.DarkRed, fontSize: 14);
                bulletList.AddItem("Red bullet item");
                bulletList.AddItem("Another item");

                var second = document.AddCustomBulletList('o', "Calibri", "0000ff", fontSize: 12);
                second.AddItem("Blue bullet item");
                second.AddItem("Second blue item");
                document.Save(openWord);
            }
        }
    }
}
