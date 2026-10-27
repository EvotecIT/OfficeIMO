using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal partial class Sections
    {
        internal static void Example_Word_Fluent_Sections_OrientationBackground(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with section orientation and background color via fluent API");
            string filePath = Path.Combine(folderPath, "Fluent_Sections_OrientationBackground.docx");
            using (var document = WordDocument.Create(filePath))
            {
                {
                    var section1 = document.Sections.Last();
                    section1 = document.AddSection();
                    section1.PageSettings.Orientation = OfficePageOrientation.Landscape;
                    document.Background.SetColorHex("FFD700");
                    var paragraph2 = section1.AddParagraph();
                    paragraph2 = paragraph2.AddText("Landscape section with background");
                    document.Save();
                }
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}