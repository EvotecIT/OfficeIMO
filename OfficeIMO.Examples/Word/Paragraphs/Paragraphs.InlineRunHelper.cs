using System;
using System.IO;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Examples.Word;

internal static partial class Paragraphs {
    internal static void Example_InlineRunHelper(string folderPath, bool openWord) {
        Console.WriteLine("[*] Creating document with inline runs");
        string filePath = Path.Combine(folderPath, "InlineRunHelper.docx");
        using (var document = OfficeIMO.Markdown.MarkdownReader
            .Parse("Hello **world** and *universe*. Visit <https://example.com>")
            .ToWordDocument()) {
            document.Save(filePath);
            if (openWord) document.OpenInApplication();
        }
    }
}
