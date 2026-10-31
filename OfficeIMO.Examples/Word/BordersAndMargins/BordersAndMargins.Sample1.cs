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

                document.Sections[0].Borders.LeftStyle = WordBorderStyle.PalmsColor;
                document.Sections[0].Borders.LeftColor = OfficeIMO.OfficeColor.Aqua;
                document.Sections[0].Borders.LeftSpace = OfficeLength.Points(24);
                document.Sections[0].Borders.LeftSize = OfficeLength.Points(3);

                document.Sections[0].Borders.RightStyle = WordBorderStyle.BabyPacifier;
                document.Sections[0].Borders.RightColor = OfficeIMO.OfficeColor.Red;
                document.Sections[0].Borders.RightSize = OfficeLength.Points(1.5);

                document.Sections[0].Borders.TopStyle = WordBorderStyle.SharksTeeth;
                document.Sections[0].Borders.TopColor = OfficeIMO.OfficeColor.GreenYellow;
                document.Sections[0].Borders.TopSize = OfficeLength.Points(1.25);

                document.Sections[0].Borders.BottomStyle = WordBorderStyle.Thick;
                document.Sections[0].Borders.BottomColor = OfficeIMO.OfficeColor.Blue;
                document.Sections[0].Borders.BottomSize = OfficeLength.Points(1.875);

                document.Save();
                if (openWord) document.OpenInApplication();
            }
        }

    }
}
