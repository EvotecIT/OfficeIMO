using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word
{
    internal static partial class TextExamples
    {
        internal static void Example_RunFormatting(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with run-level formatting");
            string filePath = Path.Combine(folderPath, "TextRunFormatting.docx");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    var paragraph1 = document.AddParagraph();
                    paragraph1 = paragraph1.AddText("Outline ");
                    paragraph1.SetOutline();
                    paragraph1 = paragraph1.AddText("Shadow ");
                    paragraph1.SetShadow();
                    paragraph1 = paragraph1.AddText("Emboss ");
                    paragraph1.SetEmboss();
                    paragraph1 = paragraph1.AddText("SmallCaps ");
                    paragraph1.SetSmallCaps();
                    paragraph1 = paragraph1.AddText("Combined");
                    paragraph1.SetOutline();
                    paragraph1.SetShadow();
                    paragraph1.SetEmboss();
                    paragraph1.SetSmallCaps();
                    document.Save();
                }
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}