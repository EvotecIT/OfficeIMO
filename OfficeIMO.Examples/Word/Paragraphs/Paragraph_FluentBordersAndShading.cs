using System;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = OfficeIMO.OfficeColor;

namespace OfficeIMO.Examples.Word;
internal static partial class Paragraphs
{
    internal static void Example_Word_Fluent_Paragraph_BordersAndShading(string folderPath, bool openWord)
    {
        Console.WriteLine("[*] Creating document with fluent bordered and shaded paragraph");
        string filePath = Path.Combine(folderPath, "Fluent_Paragraph_BordersAndShading.docx");
        using (var document = WordDocument.Create(filePath))
        {
            {
                var paragraph1 = document.AddParagraph();
                paragraph1 = paragraph1.AddText("Border and shading");
                {
                    (paragraph1.Borders).LeftStyle = WordBorderStyle.Thick;
                    (paragraph1.Borders).LeftColor = Color.Red;
                    (paragraph1.Borders).LeftSize = OfficeLength.Points(3);
                }

                paragraph1.ShadingFillColor = Color.LightGray;
                document.Save();
            }
        }

        if (openWord)
            ExampleFileLauncher.Open(filePath);
    }
}