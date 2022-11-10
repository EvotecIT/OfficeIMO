using System;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word {
    internal static partial class BordersAndMargins {

        internal static void Example_BasicPageBorders1(string folderPath, bool openWord) {
            Console.WriteLine("[*] Creating standard document with page borders 1");
            string filePath = System.IO.Path.Combine(folderPath, "Basic Document with page borders 1.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section 0");

                document.Sections[0].Borders.LeftStyle = BorderValues.PalmsColor;
                document.Sections[0].Borders.LeftColor = SixLabors.ImageSharp.Color.Aqua;
                document.Sections[0].Borders.LeftSpace = 24;
                document.Sections[0].Borders.LeftSize = 24;

                document.Sections[0].Borders.RightStyle = BorderValues.BabyPacifier;
                document.Sections[0].Borders.RightColor = SixLabors.ImageSharp.Color.Red;
                document.Sections[0].Borders.RightSize = 12;

                document.Sections[0].Borders.TopStyle = BorderValues.SharksTeeth;
                document.Sections[0].Borders.TopColor = SixLabors.ImageSharp.Color.GreenYellow;
                document.Sections[0].Borders.TopSize = 10;

                document.Sections[0].Borders.BottomStyle = BorderValues.Thick;
                document.Sections[0].Borders.BottomColor = SixLabors.ImageSharp.Color.Blue;
                document.Sections[0].Borders.BottomSize = 15;

                document.Save(openWord);
            }
        }

    }
}
