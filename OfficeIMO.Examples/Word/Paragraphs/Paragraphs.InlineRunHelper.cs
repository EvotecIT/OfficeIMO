using System;
using System.IO;
using OfficeIMO.Word;

internal static partial class Paragraphs {
    internal static void Example_InlineRunHelper(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating document with inline runs");
        string filePath = Path.Combine(folderPath, "InlineRunHelper.docx");
        using (WordDocument document = WordDocument.Create(filePath)) {
            var paragraph = document.AddParagraph();
            InlineRunHelper.AddInlineRuns(paragraph, "Hello **world** and *universe*. Visit https://example.com");
            document.Save(openWord);
        }
    }
}
