using System;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Color = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Examples.Word;

internal static partial class Paragraphs {

    internal static void Example_BasicParagraphs(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating standard document with paragraphs");
        string filePath = System.IO.Path.Combine(folderPath, "Basic Document with some paragraphs.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            var paragraph = document.AddParagraph("Basic paragraph - Page 4");
            paragraph.ParagraphAlignment = JustificationValues.Center;
            paragraph.Color = SixLabors.ImageSharp.Color.Blue;
            paragraph.AddText(" This is continutation in the same line");
            paragraph.AddBreak(BreakValues.TextWrapping);
            paragraph.AddText(" This is continuation, in another line").SetUnderline(UnderlineValues.Double).SetFontSize(15).SetColor(Color.Yellow).SetHighlight(HighlightColorValues.DarkGreen);

            var paragraph1 = document.AddParagraph("Here's another paragraph ").AddText(" which continues here, but will continue in another line ").AddBreak(BreakValues.TextWrapping).AddText("to confirm that breaks with TextWrapping is working properly");

            Console.WriteLine("+ Color: " + paragraph.Color);
            Console.WriteLine("+ Color 0: " + document.Paragraphs[0].Color);
            Console.WriteLine("+ Color 1: " + document.Paragraphs[1].Color);
            Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
            Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
            Console.WriteLine("+ Sections: " + document.Sections.Count);

            // primary section (for the whole document)
            Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);

            document.Save();
        }

        using (WordDocument document = WordDocument.Load(filePath)) {
            Console.WriteLine("+ Color 0: " + document.Paragraphs[0].Color);
            Console.WriteLine("+ Color 1: " + document.Paragraphs[1].Color);
            Console.WriteLine("+ Paragraphs: " + document.Paragraphs.Count);
            Console.WriteLine("+ PageBreaks: " + document.PageBreaks.Count);
            Console.WriteLine("+ Sections: " + document.Sections.Count);

            // primary section (for the whole document)
            Console.WriteLine("+ Paragraphs section 0: " + document.Sections[0].Paragraphs.Count);
            document.Save(openWord);
        }
    }

}
