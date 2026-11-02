using System;
using System.IO;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word;
internal static partial class Paragraphs
{
    internal static void Example_Word_Fluent_Paragraph_TextAndFormatting(string folderPath, bool openWord)
    {
        Console.WriteLine("[*] Creating document with fluent paragraph text + formatting");
        string filePath = Path.Combine(folderPath, "Fluent_Paragraph_TextAndFormatting.docx");
        using (var document = WordDocument.Create(filePath))
        {
            {
                var paragraph1 = document.AddParagraph();
                paragraph1 = paragraph1.AddText("Hello");
                paragraph1 = paragraph1.AddText(" World");
                paragraph1.SetBold();
                paragraph1.SetItalic();
                paragraph1.SetColorHex("#ff0000");
                paragraph1 = paragraph1.AddText("!");
                paragraph1.SetBold();
                document.Save();
            }
        }

        if (openWord)
            ExampleFileLauncher.Open(filePath);
    }
}