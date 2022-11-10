using System;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.Word;
using SixLabors.ImageSharp;

namespace OfficeIMO.Examples.Word {
    internal static partial class BordersAndMargins {

        internal static void Example_BasicPageBorders2(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with page borders 2");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with page borders 2.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.Background.SetColor(Color.DarkSeaGreen);

                document.AddParagraph("Section 0");

                document.Sections[0].SetBorders(WordBorder.Box);

                document.AddSection();
                document.Sections[1].SetBorders(WordBorder.Shadow);

                Console.WriteLine(document.Sections[1].Borders.Type);

                document.Save(openWord);
            }
        }
    }
}
