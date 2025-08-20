using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;

namespace OfficeIMO.Examples.Word;

internal static partial class Paragraphs {

    internal static void Example_AddFormattedText(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating document with AddFormattedText");
        string filePath = System.IO.Path.Combine(folderPath, "AddFormattedText.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            var paragraph = document.AddParagraph(string.Empty);
            paragraph.AddFormattedText("Bold", bold: true);
            paragraph.AddFormattedText(" Italic", italic: true);
            paragraph.AddFormattedText(" Underlined", underline: UnderlineValues.Single);
            document.Save(openWord);
        }
    }
}
