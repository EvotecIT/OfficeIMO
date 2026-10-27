using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.IO;

namespace OfficeIMO.Examples.Word
{
    internal static partial class FluentDocument
    {
        public static void Example_FluentTextBuilder(string folderPath, bool openWord)
        {
            Console.WriteLine("[*] Creating document with fluent text");
            string filePath = Path.Combine(folderPath, "FluentTextBuilder.docx");
            using (WordDocument document = WordDocument.Create(filePath))
            {
                {
                    var paragraph1 = document.AddParagraph();
                    paragraph1 = paragraph1.AddText("Hello");
                    paragraph1 = paragraph1.AddText(" World");
                    paragraph1.SetBold();
                    paragraph1.SetItalic();
                    paragraph1.SetColorHex("#ff0000");
                    paragraph1 = paragraph1.AddText(" Formatting");
                    paragraph1.SetUnderline(WordUnderlineStyle.Single);
                    paragraph1.SetHighlight(WordHighlightColor.Yellow);
                    paragraph1.SetFontSize(18);
                    paragraph1.SetFontFamily("Arial");
                    paragraph1.SetCapsStyle(WordCapsStyle.SmallCaps);
                    paragraph1.SetStrike();
                    document.Save();
                }
            }

            if (openWord)
                ExampleFileLauncher.Open(filePath);
        }
    }
}